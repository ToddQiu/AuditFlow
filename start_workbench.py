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
    candidates = [Path(os.environ.get(key, fallback)) / 'Microsoft/Edge/Application/msedge.exe'
                  for key, fallback in [('ProgramFiles(x86)', r'C:\Program Files (x86)'),
                                        ('ProgramFiles', r'C:\Program Files')]]
    edge = next((path for path in candidates if path.is_file()), None)
    if edge is None:
        raise RuntimeError('AuditFlow 独立窗口需要 Microsoft Edge，请安装后重试。')
    profile = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData/Local'))) / 'AuditFlow/desktop-profile'
    subprocess.Popen([str(edge), '--app=' + URL, '--user-data-dir=' + str(profile),
                      '--no-first-run', '--no-default-browser-check', '--window-size=1440,960'],
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


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
        with (ROOT / 'launcher.log').open('ab') as log:
            subprocess.Popen(
                [sys.executable, '-X', 'utf8', '-u', str(ROOT / 'server.py'), '--port', '8776'],
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
        main()
    except Exception as error:
        if sys.platform == 'win32':
            ctypes.windll.user32.MessageBoxW(None, str(error), 'AuditFlow', 0x10)
        else:
            print(error, file=sys.stderr)
        sys.exit(1)
