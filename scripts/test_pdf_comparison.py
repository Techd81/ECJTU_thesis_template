"""Ensure the fidelity gate detects layout changes, including missing pages."""
from pathlib import Path
import tempfile
import unittest

import pymupdf as fitz

from compare_word_pdf import compare


class PdfComparisonTests(unittest.TestCase):
    def test_identity_and_visible_shift(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, x in [('reference', 30), ('shifted', 31)]:
                with fitz.open() as pdf:
                    page = pdf.new_page(width=150, height=150)
                    page.insert_text((x, 50), 'Same content')
                    pdf.save(root / (name + '.pdf'))
            reference = root / 'reference.pdf'
            self.assertTrue(compare(reference, reference, root / 'identity', 72)['passed'])
            shifted = compare(reference, root / 'shifted.pdf', root / 'shift', 72)
            self.assertFalse(shifted['passed'])
            geometry = shifted['pages'][0]['text_geometry']
            self.assertTrue(geometry['text_sequence_equal'])
            self.assertAlmostEqual(geometry['median_absolute_dx_pt'], 1)
            with fitz.open(reference) as pdf:
                pdf.new_page(width=150, height=150)
                pdf.save(root / 'extra.pdf')
            extra = compare(reference, root / 'extra.pdf', root / 'extra', 72)
            self.assertFalse(extra['passed'])
            self.assertEqual(extra['pages'][1]['missing_in'], 'reference')


if __name__ == '__main__':
    unittest.main()
