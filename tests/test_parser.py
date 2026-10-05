# -*- coding: utf-8 -*-
from pathlib import Path
from openpyxl import Workbook

from cutlist.parser import read_input


def test_parser_ptbr_and_qty(tmp_path: Path):
    path = tmp_path / "entrada.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Entrada"
    ws.append(["Painel", "Marca da peça", "Perfil / material", "Função da peça",
               "Qtd. (peças)", "Corte por peça (mm)", "Comprimento total (m)", "Peso (kg)"])
    ws.append(["Painel: 1P1"])
    ws.append(["1P1", "A", "G90-1.25", "GUIA", 1, "6.044", "6,04", "9,93"])
    ws.append(["1P1", "D", "89S-41-0.8", "MONTANTE", 3, 858, "2,574", "3,03"])
    ws.append(["Subtotal do painel: 1P1 — 4 peças"])
    wb.save(path)

    pieces, subtotals, alerts = read_input(str(path))
    assert len(pieces) == 4
    assert pieces[0].length_mm == 6044
    assert pieces[1].length_mm == 858
    assert subtotals["1P1"]["pieces"] == 4
