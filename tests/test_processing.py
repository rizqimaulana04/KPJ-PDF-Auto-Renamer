from io import BytesIO
import unittest
from unittest.mock import patch
import zipfile
import fitz
import pandas as pd
from openpyxl import load_workbook
from utils import (normalize_kpj, detect_column, safe_filename, build_mapping,
                   process_zip, read_pdf_kpjs, read_workbook, KPJ_RE, report_excel)


def pdf(text):
    with fitz.open() as doc:
        page = doc.new_page()
        page.insert_text((50, 80), text)
        return doc.tobytes()


def archive(entries):
    stream = BytesIO()
    with zipfile.ZipFile(stream, 'w') as z:
        for name, payload in entries:
            z.writestr(name, payload)
    return stream.getvalue()


class ProcessingTests(unittest.TestCase):
    def test_normalization(self):
        for value in ['26133068531', '"26133068531"', '26133068531.0', '26133 068 531']:
            self.assertEqual(normalize_kpj(value), '26133068531')
        self.assertEqual(normalize_kpj('00123456789'), '00123456789')
        for value in ['123', '1261330685319', '', None, 'abc26133068531']:
            self.assertIsNone(normalize_kpj(value))
        self.assertEqual(KPJ_RE.findall('1261330685319 26133068531'), ['26133068531'])

    def test_headers_and_excel(self):
        for name in ['NO KPJ', 'No. KPJ', 'No_KPJ', 'Nomor_KPJ', 'KPJ', 'Nomor KPJ BPJS']:
            self.assertEqual(detect_column(['pegawai', name], 'kpj'), 1)
        self.assertEqual(detect_column(['New Filename', 'id'], 'name'), 0)
        self.assertIsNone(detect_column(['pegawai'], 'kpj'))
        data = BytesIO()
        pd.DataFrame({'KPJ': ['00123456789']}).to_excel(data, index=False)
        self.assertEqual(read_workbook(data.getvalue(), 'x.xlsx', 'Sheet1').iloc[0, 0], '00123456789')

    def test_sanitization(self):
        for name in ['../../a', '..\\a', 'CON', 'a/b:*?"<>|', '文' * 200]:
            result = safe_filename(name)
            self.assertNotRegex(result, r'[\\/:*?"<>|]')
            self.assertNotIn('..', result)
            self.assertLessEqual(len(result.encode('utf-8')), 184)

    def test_end_to_end_preserves_every_pdf(self):
        frame = pd.DataFrame({'KPJ': ['00123456789', '26133068531', '26133068531'],
                              'Nama Baru': ['same.pdf', 'same.pdf', 'same.pdf']})
        mapping, issues, duplicates, conflicts = build_mapping(frame, 'KPJ', 'Nama Baru')
        entries = [('sub/a.pdf', pdf('KPJ 00123456789')), ('b.pdf', pdf('KPJ 26133068531')),
                   ('c.pdf', pdf('KPJ 26133068531')), ('d.pdf', pdf('KPJ 99999999999')),
                   ('empty.pdf', pdf('No number')), ('broken.pdf', b'broken'),
                   ('ambiguous.pdf', pdf('00123456789 26133068531'))]
        result = process_zip(archive(entries), mapping, issues, duplicates, conflicts)
        self.assertEqual(result['stats']['Total PDF'], 7)
        self.assertEqual(result['stats']['Berhasil di-rename'], 3)
        self.assertEqual(result['stats']['Gagal membaca KPJ'], 2)
        with zipfile.ZipFile(BytesIO(result['zip'])) as z:
            names = z.namelist()
            self.assertEqual(len([n for n in names if n.endswith('.pdf')]), 7)
            self.assertIn('Hasil_Rename/same_3.pdf', names)
            for entry, row in zip(entries, result['rows']):
                self.assertEqual(z.read(row['Nama File Baru']), entry[1])
            self.assertEqual(z.read('laporan_mapping.xlsx'), result['report'])

    def test_conflicts(self):
        frame = pd.DataFrame({'k': ['26133068531'] * 3, 'n': ['a', 'b', 'a']})
        mapping, issues, duplicate, conflicts = build_mapping(frame, 'k', 'n')
        self.assertNotIn('26133068531', mapping)
        result = process_zip(archive([('x.pdf', pdf('26133068531'))]), mapping, issues, duplicate, conflicts)
        self.assertEqual(result['rows'][0]['Status Matching'], 'DUPLIKAT')
        self.assertTrue(result['rows'][0]['Nama File Baru'].startswith('Tidak_Ditemukan/'))

    def test_fallback(self):
        data = pdf('26133068531')
        with patch('utils.fitz.open', side_effect=RuntimeError('test fallback')):
            found, note, failed = read_pdf_kpjs(data)
        self.assertEqual(found, ['26133068531'])
        self.assertIn('Fallback', note)
        self.assertFalse(failed)

    def test_zip_validation(self):
        with self.assertRaises(ValueError):
            process_zip(archive([('a.txt', b'none')]), {}, [], set(), set())
        with self.assertRaises(zipfile.BadZipFile):
            process_zip(b'broken zip', {}, [], set(), set())
        with patch('utils.MAX_FILE_BYTES', 1), self.assertRaises(ValueError):
            process_zip(archive([('x.pdf', pdf('26133068531'))]), {}, [], set(), set())

    def test_excel_formula_not_executable(self):
        report = report_excel([{'Nama File Awal': '=1+1', 'KPJ Terdeteksi': '00123456789'}], [], {})
        book = load_workbook(BytesIO(report))
        self.assertEqual(book['Hasil Mapping']['A2'].data_type, 's')
        self.assertEqual(book['Hasil Mapping']['B2'].value, '00123456789')


if __name__ == '__main__':
    unittest.main()
