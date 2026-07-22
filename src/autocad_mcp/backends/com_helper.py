"""Standalone, medium-integrity COM trigger helper.

Launched as a plain script (never imported as part of the package) by
``com_trigger.py`` when the MCP server itself runs elevated. An elevated
process cannot see the interactive user's Running Object Table, so
``GetActiveObject`` fails with MK_E_UNAVAILABLE; this helper runs at medium
integrity and does the ``SendCommand`` on the server's behalf.

Protocol: poll for a trigger file; when it appears, delete it and fire
``(c:mcp-dispatch)`` at AutoCAD. Exits when the parent process goes away.

Usage: python com_helper.py <ipc_dir> <parent_pid> [progid]
"""

import os
import sys
import time

POLL = 0.02  # seconds
PARENT_CHECK_EVERY = 50  # poll iterations
RECONNECT_BACKOFF = 2.0  # seconds between COM reconnect attempts


def parent_alive(pid):
    import ctypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return False
    code = ctypes.c_ulong()
    ok = ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
    ctypes.windll.kernel32.CloseHandle(h)
    return bool(ok) and code.value == 259  # STILL_ACTIVE


def main():
    ipc_dir = sys.argv[1]
    parent_pid = int(sys.argv[2])
    progid = sys.argv[3] if len(sys.argv) > 3 else "AutoCAD.Application"

    trigger = os.path.join(ipc_dir, "autocad_mcp_trigger")
    status = os.path.join(ipc_dir, "autocad_mcp_helper_status")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()

    def write_status(text):
        try:
            with open(status, "w", encoding="utf-8") as f:
                f.write(f"{os.getpid()} {text}")
        except OSError:
            pass

    app = None
    last_attempt = 0.0
    i = 0
    write_status("starting")

    while True:
        i += 1
        if i % PARENT_CHECK_EVERY == 0 and not parent_alive(parent_pid):
            break

        if not os.path.exists(trigger):
            time.sleep(POLL)
            continue

        # Claim the trigger before firing so a slow SendCommand can't double-fire.
        try:
            os.remove(trigger)
        except OSError:
            time.sleep(POLL)
            continue

        if app is None and time.time() - last_attempt > RECONNECT_BACKOFF:
            last_attempt = time.time()
            try:
                app = win32com.client.GetActiveObject(progid)
                write_status("connected")
            except Exception as e:  # noqa: BLE001 — surface any COM failure to the server
                write_status(f"com_error {type(e).__name__}: {e}")
                app = None

        if app is None:
            continue

        try:
            app.ActiveDocument.SendCommand("(c:mcp-dispatch) " + chr(13))
            write_status(f"fired at {time.time():.3f}")
        except Exception as e:  # noqa: BLE001 — AutoCAD may have closed; reconnect next time
            write_status(f"send_error {type(e).__name__}: {e}")
            app = None

    try:
        os.remove(status)
    except OSError:
        pass


if __name__ == "__main__":
    main()
