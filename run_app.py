"""
run_app.py — Stealth Launcher
Starts Streamlit with no visible console window, then opens the browser.
This is the entry-point that PyInstaller will compile into the .exe.
"""

import os
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

# ── ΚΡΙΣΙΜΟ: Αποτρέπει το infinite loop στο PyInstaller frozen exe ──────────
# Όταν το exe καλεί τον εαυτό του ως subprocess, βάζουμε env variable
# ώστε να ξέρει ότι είναι το child process και να μην ξανακάνει launch
if os.environ.get("EXCELFORMATTER_CHILD") == "1":
    # Είμαστε το child process — τρέχουμε κατευθείαν το Streamlit app
    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit", "run",
        str(Path(__file__).parent / "app.py"),
        "--server.port=8501",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())

# ── Port & URL ──────────────────────────────────────────────────────────────
PORT = 8501
APP_URL = f"http://localhost:{PORT}"

# ── Resolve paths whether running as .py or frozen .exe ─────────────────────
if getattr(sys, "frozen", False):
    # Running inside a PyInstaller bundle
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

APP_SCRIPT = BASE_DIR / "app.py"

# ── Streamlit flags ──────────────────────────────────────────────────────────
STREAMLIT_ARGS = [
    "streamlit", "run", str(APP_SCRIPT),
    f"--server.port={PORT}",
    "--server.headless=true",          # don't auto-open from Streamlit itself
    "--server.fileWatcherType=none",   # no hot-reload needed in production
    "--global.developmentMode=false",
    "--browser.gatherUsageStats=false",
]

# ── Windows-only creation flags ──────────────────────────────────────────────
CREATE_NO_WINDOW = 0x08000000  # subprocess constant — hides the console


def find_streamlit_exe() -> list:
    """
    Στο PyInstaller onedir, καλούμε το streamlit module
    μέσω του ίδιου του exe με την παράμετρο __streamlit__
    """
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent

        # Ψάχνουμε streamlit.exe μέσα στους πιθανούς φακέλους
        candidates = [
            base / "_internal" / "Scripts" / "streamlit.exe",
            base / "Scripts" / "streamlit.exe",
            base / "_internal" / "streamlit.exe",
            base / "streamlit.exe",
        ]
        for c in candidates:
            if c.exists():
                return [str(c)]

        # Τελευταία λύση: χρησιμοποιούμε το runpy μέσα από το bundled exe
        # Το PyInstaller έχει όλα τα modules — καλούμε streamlit ως module
        return [sys.executable, "-c",
                "import sys; from streamlit.web import cli as stcli; sys.exit(stcli.main())"]

    # Dev mode
    return [sys.executable, "-m", "streamlit"]


def launch_streamlit() -> subprocess.Popen:
    env = os.environ.copy()
    env["EXCELFORMATTER_CHILD"] = "1"   # ← το child ξέρει να μην κάνει loop
    env["STREAMLIT_SERVER_PORT"] = str(PORT)
    env["STREAMLIT_SERVER_HEADLESS"] = "true"
    env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    env["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"

    # Καλούμε τον εαυτό μας (το ίδιο exe) με το CHILD flag
    cmd = [sys.executable]

    proc = subprocess.Popen(
        cmd,
        cwd=str(BASE_DIR),
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW,
        close_fds=True,
    )
    return proc


def wait_for_server(timeout: int = 15) -> bool:
    """Poll localhost until Streamlit is accepting connections."""
    import socket
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("localhost", PORT), timeout=1):
                return True
        except OSError:
            time.sleep(0.4)
    return False


def main():
    proc = launch_streamlit()

    # Give Streamlit a moment to bind the port, then open the browser
    ready = wait_for_server(timeout=20)

    if ready:
        webbrowser.open(APP_URL)
    else:
        # Server didn't respond in time — open anyway (Streamlit may still be loading)
        webbrowser.open(APP_URL)

    # Keep the launcher alive so the child process isn't orphaned on some OSes.
    # The user closes the app by closing the browser tab; the process ends when
    # they kill it from Task Manager or a system-tray wrapper.
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()


if __name__ == "__main__":
    main()