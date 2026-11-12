"""Turn Stock Summary lines into the rows of the Tag & Env Stock report.

One report row = one blend/flavour "family". Its envelopes on the different
machine lines (Constanta, C250, MD-20, Perfecta, MD HS, Maisa) are placed side
by side, with the matching tag next to its envelope. A tag that serves several
families (e.g. "TAG GV FLAVOUR DC C250 TBGS") is shown once, in a cell merged
across those families, exactly like the manual report.
"""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field

import pandas as pd

from . import config as C
from .readers import norm_name


def family_key(item_name: str) -> str:
    """'ENV AL FINA DC SF C250 TBGS' -> 'AL FINA'."""
    s = norm_name(item_name)
    s = re.sub(r"^(ENV|TAG)\s+", "", s)
    s = C.STRIP_PARENS.sub(" ", s)
    s = re.sub(r"-FT\b", " ", s)
    tokens = [t for t in s.split() if t not in C.STRIP_TOKENS]
    return " ".join(tokens).strip(" -")


@dataclass
class StockLine:
    item: str
    item_norm: str
    machine_line: str
    column: str            # logical column id from config.MAIN_COLUMNS
    qty: float


@dataclass
class ReportRow:
    key: str
    kind: str                                   # "env" family row or "tag" only row
    lines: dict[str, list[StockLine]] = field(default_factory=lambda: defaultdict(list))
    name: str = ""
    remark: str = ""
    remark_reason: str = ""

    def qty(self, col: str) -> float:
        return float(sum(l.qty for l in self.lines.get(col, [])))

    def has(self, col: str) -> bool:
        return bool(self.lines.get(col))

    @property
    def env_columns(self) -> list[str]:
        return [c for c in C.ENV_COLUMNS if self.has(c)]

    def all_lines(self) -> list[StockLine]:
        return [l for ls in self.lines.values() for l in ls]


@dataclass
class SharedTag:
    """A tag stock line shown once and merged across several family rows."""
    column: str
    lines: list[StockLine]
    member_keys: list[str]

    @property
    def qty(self) -> float:
        return float(sum(l.qty for l in self.lines))


@dataclass
class ExportLine:
    section: str
    item: str
    column: str            # "env" or "tag"
    qty: float


@dataclass
class Grouping:
    rows: list[ReportRow]
    shared_tags: list[SharedTag]
    export_lines: list[ExportLine]
    issues: list[dict]


def _issue(kind: str, item: str, detail: str) -> dict:
    return {"Check": kind, "Item": item, "Detail": detail}


def build_rows(stock: pd.DataFrame, bom: pd.DataFrame) -> Grouping:
    issues: list[dict] = []
    lines: list[StockLine] = []
    export_lines: list[ExportLine] = []

    for rec in stock.itertuples(index=False):
        col = C.machine_line_column(rec.machine_line)
        name_u = rec.item_norm
        looks_pm = name_u.startswith(("ENV ", "TAG "))
        if col is None:
            if looks_pm:
                issues.append(_issue("Unknown machine line", rec.item,
                                     f"Machine line '{rec.machine_line or '(blank)'}' is not mapped "
                                     f"to a report column - item left out ({rec.qty:,.2f} kg)."))
            continue
        if rec.qty == 0:
            continue
        section = C.EXPORT_EXTRA_ITEMS.get(name_u)
        if section is None:
            for title, marker in C.EXPORT_SECTIONS:
                if marker in name_u:
                    section = title
                    break
        if section is not None:
            export_lines.append(ExportLine(section, rec.item,
                                           "tag" if col in C.TAG_COLUMNS else "env", float(rec.qty)))
            continue
        lines.append(StockLine(rec.item, name_u, rec.machine_line, col, float(rec.qty)))

    # 1. Envelope families -------------------------------------------------
    rows: dict[str, ReportRow] = {}
    for ln in lines:
        if ln.column in C.ENV_COLUMNS:
            key = family_key(ln.item)
            row = rows.setdefault(key, ReportRow(key=key, kind="env"))
            row.lines[ln.column].append(ln)

    # envelope item -> family key (for BOM based tag linking)
    env_family = {ln.item_norm: family_key(ln.item) for ln in lines if ln.column in C.ENV_COLUMNS}

    # FG -> set of envelope families present in stock (from BOM)
    tag_bom = bom[bom["item_type"].isin(["TAG"]) | bom["item_norm"].str.startswith("TAG ")]
    env_bom = bom[bom["item_norm"].isin(env_family.keys())]
    fg_to_families: dict[str, set[str]] = defaultdict(set)
    for fg, itn in zip(env_bom["fg_norm"], env_bom["item_norm"]):
        fg_to_families[fg].add(env_family[itn])
    tag_to_fgs: dict[str, set[str]] = defaultdict(set)
    for fg, itn in zip(tag_bom["fg_norm"], tag_bom["item_norm"]):
        tag_to_fgs[itn].add(fg)

    # 2. Tags ---------------------------------------------------------------
    # 2a. a tag whose own name matches an envelope family joins that row
    leftover: list[StockLine] = []
    for ln in lines:
        if ln.column not in C.TAG_COLUMNS:
            continue
        target = rows.get(family_key(ln.item))
        if target is not None and target.kind == "env" and target.has(C.TAG_PAIRED_ENV[ln.column]):
            target.lines[ln.column].append(ln)
        else:
            leftover.append(ln)

    # 2b. common tags: link to families through the BOM, then hand the brand's
    #     remaining families on that machine line to the most likely tag.
    def free(f: str, col: str) -> bool:
        r = rows.get(f)
        return r is not None and r.kind == "env" and r.has(C.TAG_PAIRED_ENV[col]) and not r.has(col)

    standalone: list[StockLine] = []
    pools: dict[tuple[str, str], list[tuple[StockLine, set[str]]]] = defaultdict(list)
    for ln in leftover:
        linked = set()
        for fg in tag_to_fgs.get(ln.item_norm, ()):
            linked |= fg_to_families.get(fg, set())
        linked = {f for f in linked if f in rows and rows[f].has(C.TAG_PAIRED_ENV[ln.column])}
        linked_free = {f for f in linked if free(f, ln.column)}
        if linked and not linked_free:
            standalone.append(ln)          # its families already carry their own tag
            continue
        brand = family_key(ln.item).split(" ")[0]
        pools[(ln.column, brand)].append((ln, linked_free))

    generic = {"TEA", "-", "GMS", "COMMON", "FLAVOUR"}
    shared: dict[tuple[str, frozenset], SharedTag] = {}
    for (col, brand), tags in pools.items():
        taken = set().union(*(lf for _, lf in tags))
        assigned = {id(ln): set(lf) for ln, lf in tags}
        mates = sorted(f for f in rows if f.split(" ")[0] == brand and free(f, col) and f not in taken)
        candidates = [t for t in tags if t[1]] or tags
        if len(candidates) == 1:
            assigned[id(candidates[0][0])] |= set(mates)
        else:
            for f in mates:
                fw = set(f.split()[1:]) - generic
                best = max(candidates, key=lambda t: len(fw & (set(family_key(t[0].item).split()[1:]) - generic)))
                if fw & (set(family_key(best[0].item).split()[1:]) - generic):
                    assigned[id(best[0])].add(f)
        # a family listed under two tags stays with the bigger group
        seen: set[str] = set()
        for ln, _ in sorted(tags, key=lambda t: -len(assigned[id(t[0])])):
            members = sorted(assigned[id(ln)] - seen)
            seen |= set(members)
            if len(members) == 1:
                rows[members[0]].lines[col].append(ln)
            elif len(members) > 1:
                sk = (col, frozenset(members))
                if sk in shared:
                    shared[sk].lines.append(ln)
                else:
                    shared[sk] = SharedTag(col, [ln], members)
            else:
                standalone.append(ln)

    for ln in standalone:
        key = "TAG::" + family_key(ln.item)
        rows.setdefault(key, ReportRow(key=key, kind="tag")).lines[ln.column].append(ln)

    # 3. Row names -----------------------------------------------------------
    for row in rows.values():
        if row.kind == "env":
            for col in C.ROW_NAME_PRIORITY:
                if row.has(col):
                    # biggest stock line on the preferred machine line names the row
                    row.name = max(row.lines[col], key=lambda l: l.qty).item
                    break
        else:
            row.name = max(row.all_lines(), key=lambda l: l.qty).item
        for col, ls in row.lines.items():
            if len(ls) > 1:
                issues.append(_issue("Combined on one row", row.name,
                                     f"{C.MAIN_COLUMNS[col][1].strip()} {C.MAIN_COLUMNS[col][2]}: "
                                     + " + ".join(f"{l.item} ({l.qty:,.2f})" for l in ls)))

    ordered = order_rows(list(rows.values()), list(shared.values()))
    return Grouping(rows=ordered, shared_tags=list(shared.values()),
                    export_lines=export_lines, issues=issues)


def order_rows(rows: list[ReportRow], shared: list[SharedTag]) -> list[ReportRow]:
    """Envelope rows A-Z, tag-only rows after them (grouped by tag line).

    Shared-tag families are kept next to each other when two shared tags of the
    same column would otherwise interleave (e.g. JUMBO fruit / black / green /
    rooibos tags), so each merged cell covers a clean block.
    """
    env_rows = sorted([r for r in rows if r.kind == "env"], key=lambda r: r.name.upper())
    tag_rows = sorted([r for r in rows if r.kind == "tag"],
                      key=lambda r: (min(C.TAG_ONLY_ORDER.index(c) for c in r.lines if r.lines[c]),
                                     r.name.upper()))

    def positions(order):
        return {r.key: i for i, r in enumerate(order)}

    by_col: dict[str, list[SharedTag]] = defaultdict(list)
    for st in shared:
        by_col[st.column].append(st)

    for col, groups in by_col.items():
        pos = positions(env_rows)
        spans = [(min(pos[k] for k in g.member_keys), max(pos[k] for k in g.member_keys), g)
                 for g in groups]
        clash = set()
        for i, a in enumerate(spans):
            for b in spans[i + 1:]:
                if a[0] <= b[1] and b[0] <= a[1]:
                    clash |= {id(a[2]), id(b[2])}
            if any(env_rows[j].has(col) for j in range(a[0], a[1] + 1)):
                clash.add(id(a[2]))
        # pull each clashing group's members together at its first member
        for _, _, g in sorted(spans, key=lambda s: s[0]):
            if id(g) not in clash:
                continue
            members = [r for r in env_rows if r.key in g.member_keys]
            first = env_rows.index(members[0])
            rest = [r for r in env_rows if r.key not in g.member_keys]
            insert_at = sum(1 for r in env_rows[:first] if r.key not in g.member_keys)
            env_rows = rest[:insert_at] + members + rest[insert_at:]
    return env_rows + tag_rows


def merge_span(rows: list[ReportRow], st: SharedTag) -> tuple[int, int, bool]:
    """(first index, last index, can_merge) of a shared tag block in the row list."""
    idx = [i for i, r in enumerate(rows) if r.key in st.member_keys]
    lo, hi = min(idx), max(idx)
    clean = not any(rows[i].has(st.column) for i in range(lo, hi + 1))
    return lo, hi, clean
