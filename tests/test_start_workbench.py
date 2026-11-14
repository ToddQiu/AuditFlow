import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import start_workbench as launcher


class LauncherTests(unittest.TestCase):
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

    def test_desktop_window_uses_app_mode_and_separate_profile(self):
        with patch.object(launcher.Path, 'is_file', return_value=True), \
             patch.object(launcher.subprocess, 'Popen') as spawn:
            launcher.open_desktop_window()
        args = spawn.call_args.args[0]
        self.assertIn('--app=' + launcher.URL, args)
        self.assertTrue(any(arg.startswith('--user-data-dir=') and arg.endswith('desktop-profile') for arg in args))

    def test_missing_window_runtime_does_not_open_default_browser(self):
        with patch.object(launcher.Path, 'is_file', return_value=False), \
             patch.object(launcher.subprocess, 'Popen') as spawn:
            with self.assertRaisesRegex(RuntimeError, 'Microsoft Edge'):
                launcher.open_desktop_window()
        spawn.assert_not_called()


if __name__ == '__main__':
    unittest.main()
