import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts' / 'generate.py'
EXAMPLE = ROOT / 'examples' / 'internal-update.json'

class GeneratorTests(unittest.TestCase):
    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(SCRIPT), *map(str, args)], capture_output=True, text=True)

    def test_both_formats_read_back(self):
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader
        with tempfile.TemporaryDirectory() as d:
            ppt, pdf = Path(d) / 'a.pptx', Path(d) / 'a.pdf'
            result = self.run_cli(EXAMPLE, '--pptx', ppt, '--pdf', pdf)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)['slides'], 4)
            pres = Presentation(ppt)
            self.assertEqual(len(pres.slides), 4)
            texts = ' '.join(shape.text for slide in pres.slides for shape in slide.shapes if shape.has_text_frame)
            self.assertIn('Point d’équipe'.replace('’', "'"), texts)
            self.assertIn('Identifier les points de blocage', texts)
            pages = PdfReader(str(pdf)).pages
            self.assertEqual(len(pages), 4)
            self.assertIn('Identifier les points de blocage', pages[2].extract_text())

    def test_quote_layout_in_both_formats(self):
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader
        with tempfile.TemporaryDirectory() as d:
            spec, ppt, pdf = (Path(d) / name for name in ('deck.json', 'deck.pptx', 'deck.pdf'))
            spec.write_text(json.dumps({'slides': [{'type': 'quote', 'title': 'Une idée à partager', 'subtitle': 'Équipe projet'}]}), encoding='utf8')
            result = self.run_cli(spec, '--pptx', ppt, '--pdf', pdf)
            self.assertEqual(result.returncode, 0, result.stderr)
            text = ' '.join(s.text for s in Presentation(ppt).slides[0].shapes if s.has_text_frame)
            self.assertIn('Une idée à partager', text)
            self.assertIn('Équipe projet', text)
            self.assertIn('Une idée à partager', PdfReader(str(pdf)).pages[0].extract_text())

    def test_cards_layout_in_both_formats_and_validation(self):
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader
        with tempfile.TemporaryDirectory() as d:
            spec, ppt, pdf = (Path(d) / name for name in ('deck.json', 'deck.pptx', 'deck.pdf'))
            slide = {'type': 'cards', 'title': 'Trois priorités', 'bullets': ['Coopérer', 'Tester', 'Partager']}
            spec.write_text(json.dumps({'slides': [slide]}), encoding='utf8')
            result = self.run_cli(spec, '--pptx', ppt, '--pdf', pdf)
            self.assertEqual(result.returncode, 0, result.stderr)
            text = ' '.join(s.text for s in Presentation(ppt).slides[0].shapes if s.has_text_frame)
            for item in slide['bullets']:
                self.assertIn(item, text)
                self.assertIn(item, PdfReader(str(pdf)).pages[0].extract_text())
            slide['bullets'].append('Quatrième')
            spec.write_text(json.dumps({'slides': [slide]}), encoding='utf8')
            result = self.run_cli(spec, '--pdf', pdf)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('bullets', result.stderr)

    def test_output_modes(self):
        with tempfile.TemporaryDirectory() as d:
            for flag, name in [('--pptx', 'only.pptx'), ('--pdf', 'only.pdf')]:
                p = Path(d) / name
                result = self.run_cli(EXAMPLE, flag, p)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertTrue(p.stat().st_size > 100)

    def test_invalid_spec_and_path_collision(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'invalid.json'
            p.write_text('{"slides":[{"title":"X","bullets":[1]}]}', encoding='utf8')
            result = self.run_cli(p, '--pdf', Path(d) / 'bad.pdf')
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('bullets', result.stderr)
            result = self.run_cli(p, '--pdf', p)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('distinct', result.stderr)

    def test_rejects_overflow(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'long.json'
            p.write_text(json.dumps({'slides': [{'title': 'W' * 80}]}), encoding='utf8')
            result = self.run_cli(p, '--pdf', Path(d) / 'bad.pdf')
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('exceeds its box', result.stderr)

    def test_cli_requires_output(self):
        self.assertNotEqual(self.run_cli(EXAMPLE).returncode, 0)

if __name__ == '__main__': unittest.main()
