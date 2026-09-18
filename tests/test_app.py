from io import BytesIO
from pathlib import Path
import unittest
from unittest.mock import patch
import pandas as pd
from streamlit.testing.v1 import AppTest
from test_processing import pdf, archive


class AppTests(unittest.TestCase):
    def test_initial_page(self):
        app = AppTest.from_file(str(Path(__file__).parents[1] / 'app.py')).run()
        self.assertFalse(app.exception)
        self.assertEqual(app.title[0].value, 'KPJ PDF Auto Renamer')

    def test_upload_process_and_invalidate(self):
        excel = BytesIO()
        pd.DataFrame({'No. KPJ': ['00123456789'], 'Nama Baru': ['pegawai']}).to_excel(excel, index=False)
        excel.name = 'pegawai.xlsx'
        source = BytesIO(archive([('sub/input.pdf', pdf('00123456789'))]))
        source.name = 'input.zip'
        def upload(label, **kwargs):
            return source if 'ZIP' in label else excel
        with patch('streamlit.file_uploader', side_effect=upload):
            app = AppTest.from_file(str(Path(__file__).parents[1] / 'app.py')).run()
            self.assertFalse(app.exception)
            self.assertEqual(app.selectbox[1].value, 'No. KPJ')
            app.button[1].click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['result']['stats']['Berhasil di-rename'], 1)
            self.assertEqual(len(app.get('download_button')), 2)
            app.run()
            self.assertEqual(len(app.get('download_button')), 2)
            source = BytesIO(archive([('different.pdf', pdf('99999999999'))]))
            source.name = 'different.zip'
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get('download_button')), 0)
