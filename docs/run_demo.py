"""Run the manual's examples with storage separate from the audit workspace."""
from pathlib import Path
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse
import sys

DOCS = Path(__file__).resolve().parent
APP = DOCS.parent
sys.path.insert(0, str(APP))
import server


def configure():
    examples = DOCS / 'examples'
    server.DB = DOCS / 'demo_workspace.sqlite3'
    server.LIBRARY_ROOT = DOCS / 'demo_library'
    server.OUTPUT = DOCS / 'demo_outputs'
    server.OUTPUT.mkdir(exist_ok=True)
    server.RESEARCH = DOCS / 'demo_reading'
    server.RESEARCH.mkdir(exist_ok=True)
    server.KNOWN = {
        str(path.resolve()): {
            'path': str(path.resolve()), 'relative': '演示资料/' + path.name,
            'ext': path.suffix.lower(), 'size': path.stat().st_size,
        }
        for path in examples.iterdir()
        if path.suffix.lower() in {'.xlsx', '.docx'}
    }
    if len(server.KNOWN) < 4:
        raise RuntimeError('演示资料不完整，请保留 docs/examples 内的四个文件。')


class DemoHandler(server.Handler):
    def do_GET(self):
        if not self.valid_host():
            return self.send(403, {'error': '只接受本机地址访问。'})
        path = urlparse(self.path).path
        if path == '/api/session':
            return self.send(200, {
                'application': 'HK_Audit_MVP', 'environment': 'auditflow_demo',
                'token': server.TOKEN, 'state': server.state(),
            })
        if path in {'/', '/index.html'}:
            html = (APP / 'web' / 'index.html').read_text(encoding='utf-8')
            banner = '<p style="margin:0;padding:10px 24px;background:#244c93;color:white">AuditFlow 演示工作区 · 仅使用说明书样例 · 正式项目数据分开保存</p>'
            html = html.replace('<main>', banner + '<main>', 1)
            return self.send(200, html.encode('utf-8'), 'text/html; charset=utf-8')
        return super().do_GET()


if __name__ == '__main__':
    configure()
    httpd = ThreadingHTTPServer(('127.0.0.1', 8777), DemoHandler)
    print('AuditFlow demo http://127.0.0.1:8777/', flush=True)
    httpd.serve_forever()
