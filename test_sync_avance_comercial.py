import io
import unittest

import openpyxl

from sync_avance_comercial import COLUMNS, parse_pivot


class PivotHeadersTest(unittest.TestCase):
    def workbook(self, amount_header, extra_header=None):
        workbook = openpyxl.Workbook()
        headers = list(COLUMNS.values())
        headers[6] = amount_header
        row = ['21/09/2026', 'Vendedor', 'Talca', 'PERSONA', 'EQUIPO', 'Equipo', 119990, 1]
        if extra_header:
            headers.append(extra_header)
            row.append(119990)
        workbook.active.append(headers)
        workbook.active.append(row)
        stream = io.BytesIO()
        workbook.save(stream)
        workbook.close()
        stream.seek(0)
        return stream

    def test_both_export_formats_preserve_values(self):
        for header in ('Suma de Subtotal Bruto', 'Subtotal Bruto'):
            with self.subTest(header=header):
                data = parse_pivot(self.workbook(header))
                self.assertEqual(data['asOf'], '21/09/2026')
                self.assertEqual(data['rows'][0]['m'], 119990)
                self.assertEqual(data['rows'][0]['q'], 1)

    def test_ambiguous_or_duplicate_amount_is_rejected(self):
        for extra in ('Subtotal Bruto', 'Suma de Subtotal Bruto'):
            with self.subTest(extra=extra):
                with self.assertRaisesRegex(ValueError, 'duplicadas o alias ambiguos'):
                    parse_pivot(self.workbook('Subtotal Bruto', extra))

    def test_unknown_amount_is_rejected(self):
        with self.assertRaises(ValueError):
            parse_pivot(self.workbook('Subtotal Neto'))


if __name__ == '__main__':
    unittest.main()
