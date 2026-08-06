from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from preencher_evolucoes_excel import extract  # noqa: E402


class PreencherEvolucoesExcelTests(unittest.TestCase):
    def test_full_peripherally_inserted_central_catheter_name_maps_to_picc(self):
        values = extract(
            "Cateter venoso central de inserção periférica à esquerda, com extremidade na veia cava superior.",
            4,
        )

        self.assertEqual(values["Exame Físico - Acesso venoso? *"], "Sim")
        self.assertEqual(values["Exame Físico - Qual o acesso venoso? * (cond.)"], "Central")
        self.assertEqual(values["Exame Físico - Detalhamento do acesso central * (cond.)"], "PICC")

    def test_explicit_antibiotics_populate_selected_antibiotic_column(self):
        values = extract("# Recebe\nATB: Mero + Vanco", 7)

        expected = "Meropenem; Vancomicina"
        self.assertEqual(values["Conduta Clínica - Selecione os antibióticos em uso * (cond.)"], expected)
        self.assertEqual(values["Conduta Clínica - Antibiótico selecionado * (cond.)"], expected)

    def test_effective_discharge_populates_auditor_columns(self):
        values = extract(
            "Paciente estável. Conduta: alta hospitalar hoje.",
            3,
            evolution_date=date(2026, 7, 23),
            seed="ALTA01",
        )

        self.assertEqual(values["Parecer do Auditor - Paciente permanece internado? *"], "Não")
        self.assertEqual(
            values["Parecer do Auditor - Selecione o desfecho assistencial * (cond.)"],
            "Alta melhorada",
        )
        self.assertEqual(
            values["Parecer do Auditor - Data do desfecho * (cond.)"],
            "23/07/2026",
        )
        hour = values["Parecer do Auditor - Hora do desfecho * (cond.)"]
        self.assertGreaterEqual(hour, "08:00")
        self.assertLessEqual(hour, "12:30")

    def test_explicit_icu_accommodation_is_extracted(self):
        values = extract("Paciente permanece em UTI, estável, em monitorização.", 2)

        self.assertEqual(values["Dados da Internação - Acomodação *"], "UTI")


if __name__ == "__main__":
    unittest.main()
