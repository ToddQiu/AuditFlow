import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import start_workbench as launcher


class LauncherTests(unittest.TestCase):
    def test_frozen_launcher_starts_bundled_server(self):
        with tempfile.TemporaryDirectory() as temp, \
             patch.dict(launcher.os.environ, {'LOCALAPPDATA': temp}), \
             patch.object(launcher.sys, 'frozen', True, create=True), \
             patch.object(launcher, 'is_ready', side_effect=[False, True]), \
             patch.object(launcher.subprocess, 'Popen') as spawn, \
             patch.object(launcher, 'open_desktop_window'):
            launcher.main()
            self.assertTrue((Path(temp) / 'AuditFlow/launcher.log').exists())
        self.assertEqual(spawn.call_args.args[0], [launcher.sys.executable, '--serve', '--port', '8776'])

    def test_running_app_is_reused(self):
        with patch.object(launcher, 'is_ready', return_value=True), \
             patch.object(launcher.subprocess, 'Popen') as spawn, \
             patch.object(launcher, 'open_desktop_window') as window:
            launcher.main()
        spawn.assert_not_called()
        window.assert_called_once_with()

    def test_missing_app_starts_and_waits(self):
        with tempfile.TemporaryDirectory() as temp, \
             patch.object(launcher, 'ROOT', Path(temp)), \
             patch.object(launcher, 'is_ready', side_effect=[False, False, True]), \
             patch.object(launcher.subprocess, 'Popen') as spawn, \
             patch.object(launcher.time, 'sleep') as sleep, \
             patch.object(launcher, 'open_desktop_window'):
            launcher.main()
        spawn.assert_called_once()
        self.assertEqual(spawn.call_args.args[0][-2:], ['--port', '8776'])
        sleep.assert_called_once_with(0.25)

    def test_other_application_is_rejected(self):
        with patch.object(launcher.socket, 'create_connection'), \
             patch.object(launcher.OPENER, 'open', return_value=io.BytesIO(b'{"application":"other"}')):
            with self.assertRaisesRegex(RuntimeError, '8776'):
                launcher.is_ready()

    def test_matching_application_is_ready(self):
        with patch.object(launcher.socket, 'create_connection'), \
             patch.object(launcher.OPENER, 'open', return_value=io.BytesIO(b'{"application":"HK_Audit_MVP"}')):
            self.assertTrue(launcher.is_ready())

    def test_closed_port_is_not_ready(self):
        with patch.object(launcher.socket, 'create_connection', side_effect=ConnectionRefusedError), \
             patch.object(launcher.OPENER, 'open') as request:
            self.assertFalse(launcher.is_ready())
        request.assert_not_called()

    def test_desktop_window_uses_custom_frame_and_supports_downloads(self):
        import webview
        with patch.object(webview, 'create_window') as create, \
             patch.object(webview, 'start') as start, \
             patch.object(webview, 'settings', dict(webview.settings)):
            launcher.open_desktop_window()
            self.assertTrue(webview.settings['ALLOW_DOWNLOADS'])
            self.assertTrue(webview.settings['DRAG_REGION_DIRECT_TARGET_ONLY'])
        self.assertEqual(create.call_args.args[1], launcher.URL)
        self.assertTrue(create.call_args.kwargs['frameless'])
        self.assertFalse(create.call_args.kwargs['easy_drag'])
        self.assertEqual(start.call_args.kwargs['gui'], 'edgechromium')

    def test_window_controls_follow_native_state(self):
        controls = launcher.WindowControls()
        controls._window = Mock()
        controls.toggle_maximize()
        controls._window.maximize.assert_called_once_with()
        controls._on_maximized()
        controls.toggle_maximize()
        controls._window.restore.assert_called_once_with()
        controls._on_restored()
        self.assertFalse(controls._maximized)
        controls.minimize()
        controls._window.minimize.assert_called_once_with()
        controls.close()
        controls._window.destroy.assert_called_once_with()

    def test_resize_obeys_minimum_and_maximized_state(self):
        controls = launcher.WindowControls()
        controls._window = Mock()
        controls.resize(300, 200)
        controls._window.resize.assert_called_once_with(800, 600)
        controls._on_maximized()
        controls.resize(1200, 900)
        controls._window.resize.assert_called_once()

    def test_failed_runtime_does_not_open_default_browser(self):
        import webview
        with patch.object(webview, 'start', side_effect=RuntimeError('WebView2')), \
             patch.object(webview, 'create_window'), \
             patch.object(launcher.subprocess, 'Popen') as spawn:
            with self.assertRaisesRegex(RuntimeError, 'WebView2'):
                launcher.open_desktop_window()
        spawn.assert_not_called()


if __name__ == '__main__':
    unittest.main()
