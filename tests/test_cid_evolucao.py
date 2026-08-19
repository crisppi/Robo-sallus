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
            "queda da própria altura com trauma e TCE": "S09.9",
            "coronariano previamente, submetido a ATC de TCE/DA": "I25.1",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(infer_cid_from_evolution(text), expected)

    def test_unknown_pathology_remains_blank(self):
        self.assertEqual(infer_cid_from_evolution("paciente estável, sem queixas"), "")

    def test_explicit_diagnoses_found_in_daily_evolutions(self):
        cases = {
            "Diagnóstico atual: fratura-luxação traumática do úmero proximal esquerdo.": "S42.2",
            "Paciente com diagnóstico de pneumomia em lobo inferior esquerdo.": "J18.9",
            "Impressão: IRA KDIGO 1 por sepse foco urinário.": "N17.9",
            "Pielonefrite à direita, sem agente etiológico isolado.": "N10",
            "Paciente em acompanhamento por suboclusão intestinal por brida.": "K56.5",
            "EDA: úlcera bulbar ativa com sinal de sangramento recente.": "K26.4",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(infer_cid_from_evolution(text), expected)

    def test_recent_uncoded_evolutions(self):
        cases = {
            "1o DIH por dor torácica A/E reentrante": "R07.4",
            "Programação de marca-passo para exame de colonoscopia": "Z45.0",
            "RM evidenciou formação cística multilobulada, provável natureza gangliônica": "M67.4",
            "Submetida a hemorroidectomia por prolapso hemorroidário": "K64.8",
            "TC de tórax demonstrando bronquite inflamatória": "J40",
            "Déficit neurológico agudo compatível com AVC lacunar": "I63.9",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(infer_cid_from_evolution(text), expected)


if __name__ == "__main__":
    unittest.main()
