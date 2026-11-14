"""Open the local workbench, starting its server only when necessary."""
from pathlib import Path
import ctypes
import json
import os
import socket
import subprocess
import sys
import time
import urllib.request

ROOT = Path(__file__).resolve().parent
URL = 'http://127.0.0.1:8776'
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def open_desktop_window():
    import webview

    controls = WindowControls()
    window = webview.create_window(
        'AuditFlow · 香港审计数据填列工作台', URL, js_api=controls,
        width=1440, height=960, min_size=(800, 600), frameless=True,
        easy_drag=False, shadow=True, text_select=True, background_color='#faf9f6',
    )
    controls._window = window
    window.events.maximized += controls._on_maximized
    window.events.restored += controls._on_restored
    webview.settings['ALLOW_DOWNLOADS'] = True
    webview.settings['DRAG_REGION_DIRECT_TARGET_ONLY'] = True
    webview.start(gui='edgechromium', icon=str(ROOT / 'assets/audit-a-20261011.ico'))


class WindowControls:
    def __init__(self):
        self._window = None
        self._maximized = False

    def close(self):
        self._window.destroy()

    def minimize(self):
        self._window.minimize()

    def resize(self, width, height):
        if not self._maximized:
            self._window.resize(max(800, int(width)), max(600, int(height)))

    def toggle_maximize(self):
        if self._maximized:
            self._window.restore()
        else:
            self._window.maximize()

    def _on_maximized(self):
        self._maximized = True

    def _on_restored(self):
        self._maximized = False


def is_ready():
    try:
        with socket.create_connection(('127.0.0.1', 8776), timeout=1):
            pass
    except OSError:
        return False
    try:
        with OPENER.open(URL + '/api/session', timeout=2) as response:
            if json.load(response).get('application') == 'HK_Audit_MVP':
                return True
    except (OSError, ValueError):
        pass
    raise RuntimeError('端口 8776 已被其他程序占用，无法打开 AuditFlow。')


def main():
    if not is_ready():
        frozen = getattr(sys, 'frozen', False)
        log_dir = Path(os.environ['LOCALAPPDATA']) / 'AuditFlow' if frozen else ROOT
        log_dir.mkdir(parents=True, exist_ok=True)
        command = ([sys.executable, '--serve'] if frozen else
                   [sys.executable, '-X', 'utf8', '-u', str(ROOT / 'server.py')])
        with (log_dir / 'launcher.log').open('ab') as log:
            subprocess.Popen(
                command + ['--port', '8776'],
                cwd=ROOT, stdin=subprocess.DEVNULL, stdout=log, stderr=log,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        for _ in range(40):
            if is_ready():
                break
            time.sleep(0.25)
        else:
            raise RuntimeError('启动失败，请查看软件目录中的 launcher.log。')
    open_desktop_window()


if __name__ == '__main__':
    try:
        if '--serve' in sys.argv:
            sys.argv.remove('--serve')
            import server
            server.main()
        else:
            main()
    except Exception as error:
        if sys.platform == 'win32':
            ctypes.windll.user32.MessageBoxW(None, str(error), 'AuditFlow', 0x10)
        else:
            print(error, file=sys.stderr)
        sys.exit(1)
