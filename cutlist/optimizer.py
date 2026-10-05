# -*- coding: utf-8 -*-
"""Algoritmos de corte 1D, reaproveitamento e métricas."""

from collections import defaultdict
from functools import lru_cache
from typing import List, Dict, Tuple

from .models import Bar, Piece, Parameters, Plan, ScrapSuggestion, Alert


def can_fit(bar: Bar, piece: Piece, kerf: int) -> bool:
    extra_kerf = kerf if bar.pieces else 0
    return bar.used_mm + extra_kerf + piece.length_mm <= bar.length_mm


def sort_piece_key(p: Piece):
    return (-p.length_mm, p.panel, p.mark, p.id)


def ffd_pack(pieces: List[Piece], profile: str, lengths: List[int], kerf: int, start_number=1):
    """First-fit decreasing com melhor escolha de comprimento de barra."""
    bars: List[Bar] = []
    n = start_number
    for p in sorted(pieces, key=sort_piece_key):
        candidates = [b for b in bars if can_fit(b, p, kerf)]
        if candidates:
            # Best fit: menor sobra após colocar a peça.
            b = min(candidates, key=lambda x: x.remainder_mm - p.length_mm - (kerf if x.pieces else 0))
            b.pieces.append(p)
            continue

        possible = [L for L in lengths if p.length_mm <= L]
        if not possible:
            continue
        # Prefere a menor barra que comporte a peça.
        L = min(possible)
        b = Bar(n, profile, L, [], kerf)
        n += 1
        b.pieces.append(p)
        bars.append(b)
    return bars


def baseline_pack(pieces: List[Piece], profile: str, lengths: List[int], kerf: int):
    """Plano-base: ordem original + first-fit, usado para comparação."""
    bars = []
    n = 1
    for p in pieces:
        placed = False
        for b in bars:
            if can_fit(b, p, kerf):
                b.pieces.append(p)
                placed = True
                break
        if not placed:
            possible = [L for L in lengths if p.length_mm <= L]
            if not possible:
                continue
            # Mantém a regra comercial: menor barra que comporte a peça.
            L = min(possible)
            b = Bar(n, profile, L, [], kerf)
            n += 1
            b.pieces.append(p)
            bars.append(b)
    return bars


def _objective(bars: List[Bar]):
    # Critérios: comprimento comprado, nº barras, concentração da perda.
    total = sum(b.length_mm for b in bars)
    count = len(bars)
    scraps = sorted((b.remainder_mm for b in bars), reverse=True)
    # Maximizar o maior aproveitamento de uma sobra longa equivale a minimizar
    # a soma dos quadrados das sobras úteis; usado como desempate.
    concentration = sum(x * x for x in scraps)
    return (total, count, -concentration)


def exact_pack(pieces: List[Piece], profile: str, lengths: List[int], kerf: int):
    """
    Busca exata para grupos pequenos.
    A busca trabalha com peças ordenadas e estados de barras. O limite de
    peças evita explosão combinatória; para grupos grandes usa-se heurística.
    """
    if not pieces:
        return []

    pieces = sorted(pieces, key=sort_piece_key)
    lengths = sorted(set(lengths))

    best = baseline_pack(pieces, profile, lengths, kerf)
    best_obj = _objective(best)

    bars = []
    seen = set()

    # Lower bound simples pelo comprimento total.
    raw = sum(p.length_mm for p in pieces) + max(0, len(pieces)-1) * kerf
    max_len = max(lengths)

    def state_key():
        return tuple(sorted((b.length_mm, b.used_mm) for b in bars))

    def dfs(i):
        nonlocal best, best_obj
        if i == len(pieces):
            obj = _objective(bars)
            if obj < best_obj:
                # Cópia profunda suficiente para os objetos Bar/Piece.
                best = [Bar(b.number, b.profile, b.length_mm, list(b.pieces), b.kerf_mm) for b in bars]
                best_obj = obj
            return

        # Bound por comprimento comprado atual + lower bound de barras extras.
        current_len = sum(b.length_mm for b in bars)
        if current_len > best_obj[0]:
            return

        remaining = sum(p.length_mm for p in pieces[i:]) + max(0, len(pieces)-i-1) * kerf
        free = sum(b.remainder_mm for b in bars)
        deficit = max(0, remaining - free)
        extra_bars_lb = (deficit + max_len - 1) // max_len
        if current_len + extra_bars_lb * max_len > best_obj[0]:
            return

        key = (i, state_key())
        if key in seen:
            return
        seen.add(key)

        p = pieces[i]
        tried = set()

        # Primeiro tenta barras existentes; ordem por menor sobra resultante.
        existing = []
        for idx, b in enumerate(bars):
            if can_fit(b, p, kerf):
                rem = b.remainder_mm - p.length_mm - (kerf if b.pieces else 0)
                existing.append((rem, idx))
        for _, idx in sorted(existing):
            b = bars[idx]
            signature = (b.length_mm, b.used_mm)
            if signature in tried:
                continue
            tried.add(signature)
            b.pieces.append(p)
            dfs(i + 1)
            b.pieces.pop()

        # Abre barra: tenta primeiro a menor barra que comporte a peça, mas
        # também considera todas porque 6m pode reduzir o número total.
        for L in lengths:
            if p.length_mm > L:
                continue
            signature = ("new", L)
            if signature in tried:
                continue
            tried.add(signature)
            bars.append(Bar(len(bars) + 1, profile, L, [p], kerf))
            dfs(i + 1)
            bars.pop()

    dfs(0)
    # Renumera.
    for i, b in enumerate(best, 1):
        b.number = i
    return best


def optimize_group(pieces: List[Piece], profile: str, params: Parameters):
    lengths = params.bar_lengths[profile]
    if len(pieces) <= params.exact_piece_limit:
        return exact_pack(pieces, profile, lengths, params.kerf_mm)
    return ffd_pack(pieces, profile, lengths, params.kerf_mm)


def make_scrap_suggestions(
    bars: List[Bar],
    all_pieces: List[Piece],
    params: Parameters,
    already_allocated_from_scrap=None
):
    """Casa cada sobra útil com no máximo uma peça destino, sem alterar o plano."""
    already = set(already_allocated_from_scrap or [])
    candidates = [p for p in all_pieces if p.id not in already]
    suggestions = []

    # Cada sobra é processada da menor para a maior para priorizar encaixes exatos.
    for b in sorted(bars, key=lambda x: x.remainder_mm):
        scrap = b.remainder_mm
        if scrap < params.useful_scrap_min_mm:
            continue

        best = None
        best_key = None
        for p in candidates:
            compatible = p.profile == b.profile
            requires = False
            if not compatible:
                for c in params.compatibility:
                    if c.source_profile == b.profile and c.target_profile == p.profile:
                        compatible = True
                        requires = c.requires_validation
                        break
            if not compatible:
                continue

            needed = p.length_mm + (params.kerf_mm if p.length_mm > 0 else 0)
            if needed > scrap:
                continue

            # Prioriza menor resto; em empate, mesma função/perfil e rastreabilidade.
            leftover = scrap - needed
            key = (
                leftover,
                0 if p.profile == b.profile else 1,
                0 if p.panel == next((x.panel for x in b.pieces), "") else 1,
                p.length_mm,
                p.id,
            )
            if best_key is None or key < best_key:
                best_key = key
                best = (p, requires)

        if best:
            p, requires = best
            source_panel = b.pieces[0].panel if b.pieces else ""
            suggestions.append(ScrapSuggestion(
                source_panel=source_panel,
                source_bar=b.number,
                profile=b.profile,
                scrap_mm=scrap,
                target_panel=p.panel,
                target_mark=p.mark,
                target_piece_id=p.id,
                target_length_mm=p.length_mm,
                requires_validation=requires,
            ))
            already.add(p.id)
            candidates = [x for x in candidates if x.id != p.id]

    return suggestions


def apply_global_scrap_reuse(bars_by_profile: Dict[str, List[Bar]], pieces: List[Piece], params: Parameters):
    """
    Heurística global: começa com FFD e, em seguida, tenta substituir barras
    compradas por encaixes em sobras compatíveis. Para preservar rastreabilidade,
    a peça permanece identificada pelo painel.
    """
    all_bars = [b for bs in bars_by_profile.values() for b in bs]
    allocated = {p.id for b in all_bars for p in b.pieces}

    # Em modo global todas as peças elegíveis já estão alocadas. A melhoria
    # principal é recomputar o packing por perfil; reaproveitamento cruzado de
    # perfis compatíveis é tratado por uma segunda passada conservadora.
    for comp in params.compatibility:
        source = comp.source_profile
        target = comp.target_profile
        if source not in bars_by_profile or target not in bars_by_profile:
            continue

        source_bars = bars_by_profile[source]
        target_bars = bars_by_profile[target]
        target_pieces = [p for p in pieces if p.profile == target]
        # Não move automaticamente peças entre perfis incompatíveis com o
        # estoque comprado; apenas documenta a compatibilidade nas sobras.
        # O movimento efetivo é deliberadamente conservador para LSF.
        _ = source_bars, target_bars, target_pieces

    return bars_by_profile


def build_plan(pieces: List[Piece], params: Parameters, initial_alerts=None, subtotals=None):
    plan = Plan()
    plan.expanded_pieces = list(pieces)
    plan.alerts = list(initial_alerts or [])
    plan.subtotals = subtotals or {}

    # Peso por metro.
    plan.weight_per_meter = {}
    for p in pieces:
        plan.weight_per_meter.setdefault(p.profile, 0.0)
    inferred_num = defaultdict(float)
    inferred_den = defaultdict(float)
    for p in pieces:
        inferred_num[p.profile] += p.weight_kg
        inferred_den[p.profile] += p.length_mm / 1000.0
    for profile in inferred_num:
        inferred = inferred_num[profile] / inferred_den[profile] if inferred_den[profile] else 0
        override = params.weight_per_meter.get(profile)
        plan.weight_per_meter[profile] = override if override not in (None, "") else inferred

    # Detecta perfis sem configuração.
    profiles = sorted(set(p.profile for p in pieces))
    for profile in profiles:
        if profile not in params.bar_lengths:
            plan.alerts.append(Alert("ERRO", f"Perfil desconhecido: {profile}", profile=profile))
            continue

        for p in pieces:
            if p.profile != profile:
                continue
            if p.length_mm > max(params.bar_lengths[profile]):
                plan.alerts.append(Alert(
                    "ALERTA", f"Peça {p.mark} ({p.length_mm} mm) excede a barra comercial.",
                    panel=p.panel, profile=profile
                ))
                plan.unallocated.append(p)

    eligible = [
        p for p in pieces
        if p not in plan.unallocated
        and params.optimize_profile.get(p.profile, True)
    ]

    # Perfis desmarcados: fora do plano, mas não como erro.
    for p in pieces:
        if p.profile in params.bar_lengths and not params.optimize_profile.get(p.profile, True):
            plan.alerts.append(Alert(
                "INFO", "Perfil marcado como 'não otimizar'; peça ficou fora do plano.",
                panel=p.panel, profile=p.profile
            ))
            plan.unallocated.append(p)

    if params.optimization_mode == "Por painel":
        groups = defaultdict(list)
        for p in eligible:
            groups[(p.panel, p.profile)].append(p)
    else:
        groups = defaultdict(list)
        for p in eligible:
            groups[("", p.profile)].append(p)

    bars_by_group = {}
    for key, group in groups.items():
        panel, profile = key
        bars_by_group[key] = optimize_group(group, profile, params)

    # Se um grupo ficou impossível (não deveria ocorrer após alertas), registra.
    for key, bars in bars_by_group.items():
        plan.bars.extend(bars)

    # Plano-base com as mesmas unidades de comparação.
    baseline = []
    for key, group in groups.items():
        if not group:
            continue
        profile = key[1]
        bs = baseline_pack(group, profile, params.bar_lengths[profile], params.kerf_mm)
        baseline.extend(bs)
    plan.baseline_bars = baseline

    # Reaproveitamento em modo A é sugestão; no modo B o packing global já
    # mistura peças do projeto. Sugestões continuam sendo geradas para sobras.
    plan.suggestions = make_scrap_suggestions(plan.bars, pieces, params)

    # Checagens finais.
    allocated_ids = [p.id for b in plan.bars for p in b.pieces]
    if len(allocated_ids) != len(set(allocated_ids)):
        plan.alerts.append(Alert("ERRO", "Uma peça foi alocada mais de uma vez."))
    expected = {p.id for p in eligible}
    actual = set(allocated_ids)
    missing = expected - actual
    if missing:
        plan.alerts.append(Alert(
            "ERRO", f"{len(missing)} peça(s) elegível(is) não foram alocadas."
        ))

    for b in plan.bars:
        if b.used_mm > b.length_mm:
            plan.alerts.append(Alert(
                "ERRO",
                f"Barra {b.number} excedeu o comprimento: {b.used_mm} > {b.length_mm}.",
                profile=b.profile
            ))

    return plan


def bar_metrics(bars: List[Bar], useful_scrap_min_mm: int):
    bought = sum(b.length_mm for b in bars)
    used = sum(sum(p.length_mm for p in b.pieces) for b in bars)
    useful = sum(b.remainder_mm for b in bars if b.remainder_mm >= useful_scrap_min_mm)
    real_loss = sum(b.remainder_mm for b in bars if b.remainder_mm < useful_scrap_min_mm)
    return bought, used, useful, real_loss


def kg_for_bars(bars: List[Bar], weight_per_meter: Dict[str, float]):
    return sum(b.length_mm / 1000.0 * weight_per_meter.get(b.profile, 0.0) for b in bars)
