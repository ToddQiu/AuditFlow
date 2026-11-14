import json, os, pathlib, shutil, subprocess, tempfile, time, urllib.request
from urllib.parse import urlencode
root = pathlib.Path(__file__).resolve().parents[1]
opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
with tempfile.TemporaryDirectory(prefix='auditflow-package-') as temp:
    env = dict(os.environ, LOCALAPPDATA=temp)
    library = pathlib.Path(temp) / 'AuditFlow/data/客户资料'
    library.mkdir(parents=True)
    workbook = library / 'sample.xlsx'
    shutil.copyfile(root / 'docs/examples/AuditFlow_Demo_Working.xlsx', workbook)
    proc = subprocess.Popen([str(root / 'dist/AuditFlow/AuditFlow.exe'), '--serve', '--port', '18776'], env=env, creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        base = 'http://127.0.0.1:18776'
        def get(path):
            with opener.open(base + path, timeout=15) as response:
                return json.load(response)
        for _ in range(40):
            try:
                session = get('/api/session')
                break
            except OSError:
                if proc.poll() is not None: raise RuntimeError('Packaged server exited')
                time.sleep(.25)
        else: raise RuntimeError('Packaged server did not start')
        assert session['application'] == 'HK_Audit_MVP'
        get('/api/library')
        data = get('/api/workbook?' + urlencode({'path': str(workbook)}))
        assert data['sheets']
        for name in ('/', '/desktop.css', '/desktop.js', '/manual.html'):
            with opener.open(base + name) as response: assert response.status == 200
        print('PASS: bundled server, isolated storage, Excel parsing, desktop assets, manual')
    finally:
        proc.terminate()
        proc.wait(timeout=15)
