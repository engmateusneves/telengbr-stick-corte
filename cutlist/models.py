# -*- coding: utf-8 -*-
"""Modelos de dados do Cut list by panel."""

from dataclasses import dataclass, field
from typing import List, Optional, Dict


@dataclass
class Piece:
    id: int
    panel: str
    mark: str
    profile: str
    function: str
    length_mm: int
    weight_kg: float = 0.0
    source_row: int = 0
    source_qty: int = 1


@dataclass
class Bar:
    number: int
    profile: str
    length_mm: int
    pieces: List[Piece] = field(default_factory=list)
    kerf_mm: int = 0

    @property
    def used_mm(self) -> int:
        if not self.pieces:
            return 0
        return sum(p.length_mm for p in self.pieces) + max(0, len(self.pieces) - 1) * self.kerf_mm

    @property
    def remainder_mm(self) -> int:
        return self.length_mm - self.used_mm


@dataclass
class Compatibility:
    source_profile: str
    target_profile: str
    requires_validation: bool = True


@dataclass
class Parameters:
    bar_lengths: Dict[str, List[int]]
    kerf_mm: int = 0
    useful_scrap_min_mm: int = 100
    optimization_mode: str = "Por painel"
    steel_price_per_kg: float = 25.55
    weight_per_meter: Dict[str, Optional[float]] = field(default_factory=dict)
    optimize_profile: Dict[str, bool] = field(default_factory=dict)
    compatibility: List[Compatibility] = field(default_factory=list)
    exact_piece_limit: int = 22


@dataclass
class Alert:
    severity: str
    message: str
    panel: str = ""
    profile: str = ""
    row: int = 0


@dataclass
class ScrapSuggestion:
    source_panel: str
    source_bar: int
    profile: str
    scrap_mm: int
    target_panel: str
    target_mark: str
    target_piece_id: int
    target_length_mm: int
    requires_validation: bool = False

    @property
    def text(self) -> str:
        extra = " [requer validação do responsável técnico]" if self.requires_validation else ""
        return (
            f"Serve para: painel {self.target_panel} peça {self.target_mark} "
            f"({self.target_length_mm} mm){extra}"
        )


@dataclass
class Plan:
    bars: List[Bar] = field(default_factory=list)
    unallocated: List[Piece] = field(default_factory=list)
    suggestions: List[ScrapSuggestion] = field(default_factory=list)
    baseline_bars: List[Bar] = field(default_factory=list)
    alerts: List[Alert] = field(default_factory=list)
    expanded_pieces: List[Piece] = field(default_factory=list)
    weight_per_meter: Dict[str, float] = field(default_factory=dict)
    subtotals: Dict[str, dict] = field(default_factory=dict)
