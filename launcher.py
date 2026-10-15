import os
import sys
import time
import socket
import threading
import webbrowser
from pathlib import Path

# PyInstaller --onedir stores bundled application files in _internal.
APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
APP_PATH = APP_DIR / "dashboard.py"
HOST = "127.0.0.1"
PORT = 8501
URL = f"http://{HOST}:{PORT}"

os.chdir(APP_DIR)

def open_browser_when_ready():
    # Wait until Streamlit's local server accepts connections.
    deadline = time.time() + 90
    while time.time() < deadline:
        try:
            with socket.create_connection((HOST, PORT), timeout=1):
                time.sleep(1)
                webbrowser.open(URL, new=2)
                return
        except OSError:
            time.sleep(1)
    # Last attempt gives a useful URL even if startup is unusually slow.
    try:
        webbrowser.open(URL, new=2)
    except Exception:
        pass

threading.Thread(target=open_browser_when_ready, daemon=True).start()

sys.argv = [
    "streamlit",
    "run",
    str(APP_PATH),
    "--server.headless=true",
    "--global.developmentMode=false",
    f"--server.port={PORT}",
    f"--server.address={HOST}",
    f"--browser.serverAddress={HOST}",
    f"--browser.serverPort={PORT}",
    "--browser.gatherUsageStats=false",
]

from streamlit.web import cli as stcli
stcli.main()
