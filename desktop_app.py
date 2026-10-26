import sys
import os
import time
import socket
import threading
import logging
from app import app, get_target_file_from_args

# Silence unnecessary werkzeug log spam in desktop window mode
log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)

def get_free_port():
    """Try port 5000 first; if in use, find an available ephemeral port."""
    for port in [5000, 5001, 5050]:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                return port
        except OSError:
            continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]

def wait_for_server(port, timeout=10.0):
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                if s.connect_ex(('127.0.0.1', port)) == 0:
                    return True
        except Exception:
            pass
        time.sleep(0.1)
    return False

def debug_log(msg):
    try:
        base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
        with open(os.path.join(base, 'debug_runtime.log'), 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%X')}] {msg}\n")
    except Exception:
        pass

def main():
    debug_log(f"main() started, sys.argv: {sys.argv}")
    import webview
    debug_log("webview imported")

    # Detect file passed via Open With or command line
    target_file = get_target_file_from_args()
    if target_file:
        app.config['INITIAL_FILE_PATH'] = target_file
        debug_log(f"Open with target file detected: {target_file}")
    else:
        debug_log("No target file specified in sys.argv")

    port = get_free_port()
    debug_log(f"port selected: {port}")

    # Run Flask in background daemon thread
    def run_flask():
        try:
            debug_log("Flask starting...")
            app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)
        except Exception as e:
            debug_log(f"Flask failed: {e}")

    server_thread = threading.Thread(target=run_flask, daemon=True)
    server_thread.start()

    # Wait for the local server to be ready
    ready = wait_for_server(port)
    debug_log(f"Server ready status: {ready}")

    # Path to icon if present
    if getattr(sys, 'frozen', False):
        bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    else:
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
    
    icon_path = os.path.join(bundle_dir, 'app_icon.ico')
    if not os.path.exists(icon_path):
        icon_path = None
    debug_log(f"icon_path: {icon_path}")

    # Create native Windows desktop application window
    debug_log("Creating window...")
    window_title = f"{os.path.basename(target_file)} - Excel Web Pro" if target_file else "Excel Web Pro"
    window = webview.create_window(
        title=window_title,
        url=f'http://127.0.0.1:{port}',
        width=1400,
        height=860,
        min_size=(960, 640),
        background_color='#1E1E1E',
        text_select=True,
        zoomable=True,
        confirm_close=False
    )
    debug_log(f"Window object: {window}")

    debug_log("Calling webview.start()...")
    webview.start(debug=False, icon=icon_path)
    debug_log("webview.start() ended normally")
    sys.exit(0)

if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    try:
        main()
    except Exception as e:
        import traceback
        err_file = os.path.join(
            os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__),
            'app_error.log'
        )
        with open(err_file, 'w', encoding='utf-8') as f:
            traceback.print_exc(file=f)
        sys.exit(1)
