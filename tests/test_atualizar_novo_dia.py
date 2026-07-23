from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from atualizar_novo_dia import generate_clinical_base  # noqa: E402


class AtualizarNovoDiaTests(unittest.TestCase):
    def test_reuses_evolution_dates_it_and_fills_derived_columns(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Preenchimento"
        headers = [
            "Nome",
            "Iniciais",
            "Senha",
            "Dias internado",
            "ID internação",
            "Dados da Internação - CID de internação *",
            "Dados da Internação - CID ajustado *",
            "Dados da Internação - Tempo de existência da doença *",
            "Dados da Internação - Nomenclatura do tempo de existência da doença *",
            "Exame Físico - Estado geral *",
            "Exame Físico - PA Sistólica max (mmHg) *",
            "Exame Físico - PA Diastólica max (mmHg) *",
            "Exame Físico - FC máx. (bpm) *",
            "Data da evolução",
            "evolucao",
        ]
        sheet.append(headers)
        sheet.append(
            [
                "Paciente Teste",
                "PT",
                "ABC123",
                4,
                99,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                "21/07/2026",
                "CID J18.9\nBEG\nPA 120x80\nFC 88",
            ]
        )

        patients = [
            {
                "nomeCompleto": "Paciente Teste",
                "nomeIniciais": "PT",
                "senha": "ABC123",
                "diasInternados": 5,
                "idInternacao": 99,
            }
        ]

        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "base.xlsx"
            stats = generate_clinical_base(
                patients,
                Path("modelo_nao_usado.xlsx"),
                output,
                previous_workbook=workbook,
                evolution_date=date(2026, 7, 22),
            )
            result = load_workbook(output)

        result_sheet = result["Preenchimento"]
        columns = {str(cell.value): cell.column for cell in result_sheet[1] if cell.value}
        self.assertEqual(result_sheet.cell(2, columns["evolucao"]).value, "CID J18.9\nBEG\nPA 120x80\nFC 88")
        self.assertEqual(result_sheet.cell(2, columns["Data da evolução"]).value, "22/07/2026")
        self.assertEqual(result_sheet.cell(2, columns["Dados da Internação - CID de internação *"]).value, "J18.9")
        self.assertEqual(result_sheet.cell(2, columns["Dados da Internação - CID ajustado *"]).value, "J18.9")
        self.assertEqual(result_sheet.cell(2, columns["Exame Físico - Estado geral *"]).value, "BEG – Bom Estado Geral")
        self.assertEqual(result_sheet.cell(2, columns["Exame Físico - PA Sistólica max (mmHg) *"]).value, 120)
        self.assertEqual(result_sheet.cell(2, columns["Exame Físico - PA Diastólica max (mmHg) *"]).value, 80)
        self.assertEqual(result_sheet.cell(2, columns["Exame Físico - FC máx. (bpm) *"]).value, 88)
        self.assertEqual(result_sheet.cell(2, columns["Dados da Internação - Tempo de existência da doença *"]).value, 5)
        self.assertEqual(stats["evolucoes_reaproveitadas"], 1)
        self.assertEqual(stats["linhas_derivadas"], 1)


if __name__ == "__main__":
    unittest.main()
