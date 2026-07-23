import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from cid_evolucao import infer_cid_from_evolution  # noqa: E402


class CidEvolutionTests(unittest.TestCase):
    def test_explicit_cid_has_priority(self):
        self.assertEqual(infer_cid_from_evolution("diagnóstico diverso. CID R31"), "R31")

    def test_common_pathologies(self):
        cases = {
            "PO de apendicectomia por apendicite": "K35.9",
            "anemia ferropriva secundária a menorragia": "D50.0",
            "fratura transtrocantérica após queda": "S72.1",
            "hipotensão liquórica com fístula liquórica": "G96.0",
            "adenocarcinoma de pulmão metastático": "C34.9",
            "neutropenia febril": "D70",
            "internação para ablação de TRN": "I47.1",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(infer_cid_from_evolution(text), expected)

    def test_unknown_pathology_remains_blank(self):
        self.assertEqual(infer_cid_from_evolution("paciente estável, sem queixas"), "")


if __name__ == "__main__":
    unittest.main()
