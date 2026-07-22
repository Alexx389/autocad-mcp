"""COM-based dispatch trigger for AutoCAD.

Replaces the old ``PostMessage(WM_CHAR)`` approach, which silently does nothing
on AutoCAD 2024+: the command line is a WPF control, so synthesized character
messages posted to the MDIClient are never routed to the command interpreter.

Two paths, chosen at connect time:

* **Direct** — the server runs unelevated: call ``SendCommand`` in-process.
* **Helper** — the server runs elevated: an elevated process cannot see the
  interactive user's Running Object Table (``GetActiveObject`` fails with
  ``MK_E_UNAVAILABLE``), so a medium-integrity helper is spawned with the
  shell's token and fires on our behalf via a trigger file.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
from pathlib import Path

import structlog

log = structlog.get_logger()

TRIGGER_NAME = "autocad_mcp_trigger"
STATUS_NAME = "autocad_mcp_helper_status"
HELPER_PATH = Path(__file__).with_name("com_helper.py")

# Ordered by preference: version-specific ProgIDs are what AutoCAD actually
# registers; the generic one is often absent.
PROGID_CANDIDATES = (
    "AutoCAD.Application.25.1",  # 2026
    "AutoCAD.Application.25.0",  # 2025
    "AutoCAD.Application.24.3",  # 2024
    "AutoCAD.Application",
)


def is_elevated() -> bool:
    """True if this process runs at high integrity."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def resolve_progid() -> str:
    """Pick the first AutoCAD ProgID that is actually registered."""
    try:
        import winreg

        for progid in PROGID_CANDIDATES:
            try:
                winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid))
                return progid
            except OSError:
                continue
    except ImportError:
        pass
    return PROGID_CANDIDATES[0]


def _shell_token():
    """Duplicate the desktop shell's primary token (medium integrity)."""
    import ctypes

    import win32api
    import win32con
    import win32process
    import win32security

    # SeImpersonatePrivilege is held but disabled by default for admins.
    self_tok = win32security.OpenProcessToken(
        win32api.GetCurrentProcess(),
        win32con.TOKEN_ADJUST_PRIVILEGES | win32con.TOKEN_QUERY,
    )
    luid = win32security.LookupPrivilegeValue(None, "SeImpersonatePrivilege")
    win32security.AdjustTokenPrivileges(self_tok, False, [(luid, win32con.SE_PRIVILEGE_ENABLED)])

    hwnd = ctypes.windll.user32.GetShellWindow()
    if not hwnd:
        raise RuntimeError("no shell window — explorer.exe is not running")
    _, pid = win32process.GetWindowThreadProcessId(hwnd)

    hproc = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
    htok = win32security.OpenProcessToken(hproc, win32con.TOKEN_DUPLICATE | win32con.TOKEN_QUERY)
    return win32security.DuplicateTokenEx(
        htok,
        win32security.SecurityImpersonation,
        win32con.TOKEN_ALL_ACCESS,
        win32security.TokenPrimary,
        None,
    )


def _spawn_unelevated(cmdline: str, cwd: str) -> int:
    """Start a process at medium integrity using the shell's token.

    ``CreateProcessAsUser`` needs SeAssignPrimaryToken, which even admins lack;
    ``CreateProcessWithTokenW`` only needs SeImpersonate.
    """
    import ctypes
    from ctypes import wintypes

    class STARTUPINFOW(ctypes.Structure):
        _fields_ = [
            ("cb", wintypes.DWORD),
            ("lpReserved", wintypes.LPWSTR),
            ("lpDesktop", wintypes.LPWSTR),
            ("lpTitle", wintypes.LPWSTR),
            ("dwX", wintypes.DWORD),
            ("dwY", wintypes.DWORD),
            ("dwXSize", wintypes.DWORD),
            ("dwYSize", wintypes.DWORD),
            ("dwXCountChars", wintypes.DWORD),
            ("dwYCountChars", wintypes.DWORD),
            ("dwFillAttribute", wintypes.DWORD),
            ("dwFlags", wintypes.DWORD),
            ("wShowWindow", wintypes.WORD),
            ("cbReserved2", wintypes.WORD),
            ("lpReserved2", ctypes.POINTER(ctypes.c_byte)),
            ("hStdInput", wintypes.HANDLE),
            ("hStdOutput", wintypes.HANDLE),
            ("hStdError", wintypes.HANDLE),
        ]

    class PROCESS_INFORMATION(ctypes.Structure):
        _fields_ = [
            ("hProcess", wintypes.HANDLE),
            ("hThread", wintypes.HANDLE),
            ("dwProcessId", wintypes.DWORD),
            ("dwThreadId", wintypes.DWORD),
        ]

    STARTF_USESHOWWINDOW = 0x00000001
    SW_HIDE = 0
    CREATE_NO_WINDOW = 0x08000000

    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    pi = PROCESS_INFORMATION()

    fn = ctypes.windll.advapi32.CreateProcessWithTokenW
    fn.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPCWSTR,
        wintypes.LPWSTR,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.LPCWSTR,
        ctypes.POINTER(STARTUPINFOW),
        ctypes.POINTER(PROCESS_INFORMATION),
    ]

    token = _shell_token()
    buf = ctypes.create_unicode_buffer(cmdline)
    if not fn(int(token), 0, None, buf, CREATE_NO_WINDOW, None, cwd, ctypes.byref(si), ctypes.byref(pi)):
        raise ctypes.WinError(ctypes.windll.kernel32.GetLastError())

    ctypes.windll.kernel32.CloseHandle(pi.hProcess)
    ctypes.windll.kernel32.CloseHandle(pi.hThread)
    return int(pi.dwProcessId)


class ComDispatchTrigger:
    """Fires ``(c:mcp-dispatch)`` at a running AutoCAD."""

    def __init__(self, ipc_dir: Path):
        self._ipc_dir = Path(ipc_dir)
        self._progid = resolve_progid()
        self._elevated = is_elevated()
        self._app = None  # cached COM object (direct mode only)
        self._helper_pid: int | None = None
        self._lock = threading.Lock()

    @property
    def mode(self) -> str:
        return "helper" if self._elevated else "direct"

    @property
    def progid(self) -> str:
        return self._progid

    def info(self) -> dict:
        return {
            "mode": self.mode,
            "progid": self._progid,
            "elevated": self._elevated,
            "helper_pid": self._helper_pid,
        }

    def fire(self) -> None:
        """Trigger one dispatch. Raises on unrecoverable failure."""
        with self._lock:
            if self._elevated:
                self._fire_via_helper()
            else:
                self._fire_direct()

    # --- direct (unelevated server) ---

    def _fire_direct(self) -> None:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        if self._app is None:
            self._app = win32com.client.GetActiveObject(self._progid)
        try:
            self._app.ActiveDocument.SendCommand("(c:mcp-dispatch) " + chr(13))
        except Exception:
            # Stale COM reference (drawing closed, AutoCAD restarted) — retry once.
            self._app = win32com.client.GetActiveObject(self._progid)
            self._app.ActiveDocument.SendCommand("(c:mcp-dispatch) " + chr(13))

    # --- helper (elevated server) ---

    @staticmethod
    def _pid_alive(pid: int | None) -> bool:
        if pid is None:
            return False
        import ctypes

        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if not h:
            return False
        code = ctypes.c_ulong()
        ok = ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(h)
        return bool(ok) and code.value == 259  # STILL_ACTIVE

    def _status_pid(self) -> int | None:
        """PID the helper reported for itself.

        The venv's python.exe is a launcher stub that re-execs the real
        interpreter, so the PID returned by CreateProcessWithTokenW is not the
        PID that ends up running the helper. The status file is authoritative.
        """
        try:
            first = (self._ipc_dir / STATUS_NAME).read_text(encoding="utf-8").split()[0]
            return int(first)
        except (OSError, ValueError, IndexError):
            return None

    def _helper_running(self) -> bool:
        if self._pid_alive(self._status_pid()):
            return True
        return self._pid_alive(self._helper_pid)

    def ensure_helper(self) -> None:
        """Start the medium-integrity helper if it isn't already running."""
        if not self._elevated or self._helper_running():
            return
        # Drop any status left by a dead helper so we never read a stale PID.
        try:
            (self._ipc_dir / STATUS_NAME).unlink()
        except OSError:
            pass
        cmdline = (
            f'"{sys.executable}" "{HELPER_PATH}" "{self._ipc_dir}" '
            f"{os.getpid()} {self._progid}"
        )
        self._helper_pid = _spawn_unelevated(cmdline, str(HELPER_PATH.parent))
        log.info("com_helper_started", pid=self._helper_pid, progid=self._progid)

    def helper_status(self) -> str | None:
        try:
            return (self._ipc_dir / STATUS_NAME).read_text(encoding="utf-8").strip()
        except OSError:
            return None

    def _fire_via_helper(self) -> None:
        self.ensure_helper()
        trigger = self._ipc_dir / TRIGGER_NAME
        trigger.write_text(str(time.time()), encoding="utf-8")

    def shutdown(self) -> None:
        """Stop the helper, if we started one."""
        for pid in (self._status_pid(), self._helper_pid):
            if not self._pid_alive(pid):
                continue
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"], capture_output=True, check=False
                )
            except OSError:
                pass
        self._helper_pid = None
