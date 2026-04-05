"""
OpenOrder Desktop Application

Launches the FastAPI backend in a background thread and opens
a native desktop window using pywebview. No browser needed,
no terminal, no server management — just double-click and go.
"""

import sys
import os
import threading
import time

# When running as a PyInstaller bundle, __file__ is inside _internal/
# We need to figure out the right base directory for all paths.
if getattr(sys, 'frozen', False):
    # Running as exe — base is the _internal directory
    BASE_DIR = sys._MEIPASS
    # Set working directory to exe's folder (for output/, hymnal-json/, etc.)
    os.chdir(os.path.dirname(sys.executable))
else:
    # Running as script — base is the project root
    BASE_DIR = os.path.dirname(__file__)

# Ensure the backend module is importable
sys.path.insert(0, os.path.join(BASE_DIR, "backend") if not getattr(sys, 'frozen', False)
                else BASE_DIR)


def start_server():
    """Start the FastAPI server in a background thread."""
    import uvicorn
    from app.main import app

    uvicorn.run(app, host="127.0.0.1", port=8316, log_level="warning")


def wait_for_server(timeout=10):
    """Wait until the server is responding."""
    import urllib.request
    start = time.time()
    while time.time() - start < timeout:
        try:
            urllib.request.urlopen("http://127.0.0.1:8316/api/health")
            return True
        except Exception:
            time.sleep(0.2)
    return False


def main():
    # Start server in background thread
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    # Wait for it to be ready
    if not wait_for_server():
        print("ERROR: Server failed to start within 10 seconds.")
        sys.exit(1)

    # Open native window
    import webview

    window = webview.create_window(
        "OpenOrder",
        "http://127.0.0.1:8316",
        width=1100,
        height=850,
        min_size=(800, 600),
    )

    # Start the webview event loop (blocks until window is closed)
    webview.start()


if __name__ == "__main__":
    main()
