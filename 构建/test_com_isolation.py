# -*- coding: utf-8 -*-
import sys
import types
import unittest
from unittest.mock import MagicMock, patch


class ComIsolationTest(unittest.TestCase):
    def test_unavailable_isolated_instance_does_not_close_user_wps(self):
        import report_engine

        user_document = MagicMock()
        user_document.Name = "正在编辑的用户文档.docx"
        user_app = MagicMock()
        user_app.Documents.Count = 1
        user_app.Documents.Open.return_value = MagicMock()

        pythoncom = types.ModuleType("pythoncom")
        pythoncom.CoInitialize = MagicMock()
        pythoncom.CoUninitialize = MagicMock()
        win32 = types.ModuleType("win32com.client")
        win32.gencache = MagicMock()
        win32.gencache.EnsureDispatch.side_effect = [user_app, user_app]
        win32.DispatchEx = MagicMock(side_effect=RuntimeError("DispatchEx unavailable"))
        win32com = types.ModuleType("win32com")
        win32com.client = win32

        with patch.dict(sys.modules, {
            "pythoncom": pythoncom,
            "win32com": win32com,
            "win32com.client": win32,
        }):
            report_engine.update_fields_com("generated-report.docx")

        user_document.Close.assert_not_called()
        user_app.Quit.assert_not_called()
        self.assertEqual(win32.DispatchEx.call_args_list, [
            unittest.mock.call("KWps.Application"),
            unittest.mock.call("Word.Application"),
        ])
        win32.gencache.EnsureDispatch.assert_not_called()
        self.assertEqual(user_document.Name, "正在编辑的用户文档.docx")

    def test_refresh_quits_only_its_isolated_instance(self):
        import report_engine

        user_document = MagicMock()
        isolated_doc = MagicMock()
        isolated_app = MagicMock()
        isolated_app.Documents.Open.return_value = isolated_doc

        pythoncom = types.ModuleType("pythoncom")
        pythoncom.CoInitialize = MagicMock()
        pythoncom.CoUninitialize = MagicMock()
        win32 = types.ModuleType("win32com.client")
        win32.DispatchEx = MagicMock(return_value=isolated_app)
        win32.gencache = MagicMock()
        win32com = types.ModuleType("win32com")
        win32com.client = win32

        with patch.dict(sys.modules, {
            "pythoncom": pythoncom,
            "win32com": win32com,
            "win32com.client": win32,
        }):
            report_engine.update_fields_com("generated-report.docx")

        isolated_doc.Close.assert_called_once_with(False)
        isolated_app.Quit.assert_called_once_with()
        user_document.Close.assert_not_called()


if __name__ == "__main__":
    unittest.main()
