"""
OpenOrder Desktop Application

Launches the FastAPI backend in a background thread and opens
a native desktop window using pywebview. No browser needed,
no terminal, no server management — just double-click and go.
"""

import sys
import os
import logging
import threading
import time
import traceback
from datetime import datetime

# When running as a PyInstaller bundle, __file__ is inside _internal/
# We need to figure out the right base directory for all paths.
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
    os.chdir(APP_DIR)
else:
    BASE_DIR = os.path.dirname(__file__)
    APP_DIR = BASE_DIR

# Set up logging — always log to file, console only in dev
LOG_DIR = os.path.join(APP_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"openorder-{datetime.now().strftime('%Y-%m-%d')}.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ]
)
logger = logging.getLogger("OpenOrder")

# Also capture unhandled exceptions
def _exception_handler(exc_type, exc_value, exc_tb):
    logger.error("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))

sys.excepthook = _exception_handler

# Ensure the backend module is importable
sys.path.insert(0, os.path.join(BASE_DIR, "backend") if not getattr(sys, 'frozen', False)
                else BASE_DIR)


def start_server():
    """Start the FastAPI server in a background thread."""
    try:
        # In windowed mode, sys.stdout/stderr are None — uvicorn's logger
        # crashes on sys.stdout.isatty(). Redirect to devnull.
        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w")
        if sys.stderr is None:
            sys.stderr = open(os.devnull, "w")

        import uvicorn
        from app.main import app
        logger.info("Starting server on port 8316")
        uvicorn.run(app, host="127.0.0.1", port=8316, log_level="warning")
    except Exception:
        logger.error("Server failed to start:\n%s", traceback.format_exc())


def wait_for_server(timeout=10):
    """Wait until the server is responding."""
    import urllib.request
    start = time.time()
    while time.time() - start < timeout:
        try:
            urllib.request.urlopen("http://127.0.0.1:8316/api/health")
            logger.info("Server is ready")
            return True
        except Exception:
            time.sleep(0.2)
    logger.error("Server did not respond within %d seconds", timeout)
    return False


def main():
    logger.info("OpenOrder starting (frozen=%s, base=%s, app=%s)",
                getattr(sys, 'frozen', False), BASE_DIR, APP_DIR)

    # Start server in background thread
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    # Wait for it to be ready
    if not wait_for_server():
        logger.error("Exiting — server failed to start")
        sys.exit(1)

    try:
        import webview
        logger.info("Opening webview window")

        # Find icon path — bundled or dev
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "resources", "images", "openorder.ico")
        else:
            icon_path = os.path.join(BASE_DIR, "resources", "images", "openorder.ico")

        window = webview.create_window(
            "OpenOrder",
            "http://127.0.0.1:8316",
            width=1100,
            height=850,
            min_size=(800, 600),
        )

        # Start the webview event loop (blocks until window is closed)
        webview.start(icon=icon_path if os.path.exists(icon_path) else None)
        logger.info("Window closed, shutting down")
    except Exception:
        logger.error("Webview failed:\n%s", traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
