# -*- coding: utf-8 -*-
import math

from cutlist.models import Piece, Parameters
from cutlist.optimizer import exact_pack, can_fit


def piece(i, panel, mark, length, profile="89S-41-0.8"):
    return Piece(i, panel, mark, profile, "MONTANTE", length)


def params(kerf=0):
    return Parameters(
        bar_lengths={"89S-41-0.8": [3000, 6000]},
        kerf_mm=kerf,
        useful_scrap_min_mm=100,
        optimization_mode="Por painel",
        weight_per_meter={"89S-41-0.8": 1.0},
        optimize_profile={"89S-41-0.8": True},
    )


def test_case_1_17x858_minimum_18m_and_426_scrap():
    ps = [piece(i, "1P1", "D" if i > 1 else "C", 858) for i in range(1, 18)]
    bars = exact_pack(ps, "89S-41-0.8", [3000, 6000], 0)
    assert sum(b.length_mm for b in bars) == 18000
    assert sum(b.remainder_mm for b in bars) == 3414
    assert any(b.length_mm == 3000 and len(b.pieces) == 3 and b.remainder_mm == 426 for b in bars)


def test_case_2_2698_2698_568_leaves_36():
    ps = [
        piece(1, "T2", "A", 2698),
        piece(2, "T2", "B", 2698),
        piece(3, "T2", "C", 568),
    ]
    bars = exact_pack(ps, "89S-41-0.8", [3000, 6000], 0)
    assert any(b.length_mm == 6000 and sorted(p.length_mm for p in b.pieces) == [568, 2698, 2698] for b in bars)
    assert sum(b.remainder_mm for b in bars) == 36


def test_case_3_302_scrap_fits_298_with_zero_kerf_but_not_4():
    b = exact_pack(
        [piece(1, "T3", "A", 2698), piece(2, "T3", "B", 298)],
        "89S-41-0.8", [3000], 0
    )
    assert len(b) == 1
    assert b[0].remainder_mm == 4

    b4 = exact_pack(
        [piece(1, "T3", "A", 2698), piece(2, "T3", "B", 298)],
        "89S-41-0.8", [3000], 4
    )
    assert len(b4) == 2


def test_case_4_500_pieces_processes():
    ps = [piece(i, f"P{i//20}", f"M{i}", 858 + (i % 5) * 7) for i in range(500)]
    bars = exact_pack(ps[:20], "89S-41-0.8", [3000, 6000], 0)
    assert all(b.used_mm <= b.length_mm for b in bars)

    # A rotina usada para grandes grupos é FFD; aqui validamos que a
    # construção de peças e o encaixe permanecem válidos em escala.
    from cutlist.optimizer import ffd_pack
    bars500 = ffd_pack(ps, "89S-41-0.8", [3000, 6000], 0)
    assert len(bars500) > 0
    assert sum(len(b.pieces) for b in bars500) == 500
    assert all(b.used_mm <= b.length_mm for b in bars500)
