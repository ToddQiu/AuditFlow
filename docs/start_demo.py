"""Open the reproducible manual demo without changing the production database."""
from pathlib import Path
import ctypes
import json
import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser

DOCS = Path(__file__).resolve().parent
URL = 'http://127.0.0.1:8777'
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def is_ready():
    try:
        with socket.create_connection(('127.0.0.1', 8777), timeout=1):
            pass
    except OSError:
        return False
    try:
        with OPENER.open(URL + '/api/session', timeout=2) as response:
            if json.load(response).get('environment') == 'auditflow_demo':
                return True
    except (OSError, ValueError):
        pass
    raise RuntimeError('端口 8777 已被其他程序占用，无法打开演示工作区。')


def main():
    if not is_ready():
        with (DOCS / 'demo_launcher.log').open('ab') as log:
            subprocess.Popen(
                [sys.executable, '-X', 'utf8', '-u', str(DOCS / 'run_demo.py')],
                cwd=DOCS.parent, stdin=subprocess.DEVNULL, stdout=log, stderr=log,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        for _ in range(40):
            if is_ready():
                break
            time.sleep(0.25)
        else:
            raise RuntimeError('演示启动失败，请查看 docs/demo_launcher.log。')
    if not webbrowser.open(URL, new=2):
        raise RuntimeError('浏览器未能打开，请访问 ' + URL)


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        if sys.platform == 'win32':
            ctypes.windll.user32.MessageBoxW(None, str(error), 'AuditFlow Demo', 0x10)
        else:
            print(error, file=sys.stderr)
        sys.exit(1)
