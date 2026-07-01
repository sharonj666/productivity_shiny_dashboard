#!/usr/bin/env python3
"""Clean 2025 ROST productivity and resighting data.

All data manipulation lives in this script. The companion notebook reads the
generated CSV files and is limited to table display, statistics, and figures.
The script deliberately uses only the Python standard library so that it can
run before the optional notebook dependencies are installed.
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import statistics
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable


DEFAULT_YEAR = 2025
SPECIES = "ROST"
VALID_STATUS = {"A", "P", "M", "D", "L", "H", "F"}
PLACEHOLDER_PFR = {"", "UNBANDED"}
# Stakeholder-adjudicated outcomes that are not represented by Status = F or an
# eligible resight in the supplied workbooks. Keep these explicit and keyed by
# year and unique PFR so the exception remains reviewable.
REVIEW_CONFIRMED_FLEDGLINGS = {
    (2025, "XT0"),
    (2025, "YE0"),
}
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def excel_column_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference)
    if not letters:
        raise ValueError(f"Invalid Excel cell reference: {reference}")
    number = 0
    for character in letters.group(0):
        number = number * 26 + ord(character) - 64
    return number


def workbook_sheet_names(path: Path) -> list[str]:
    """Return worksheet names from an xlsx workbook."""
    with zipfile.ZipFile(path) as archive:
        workbook = ET.parse(archive.open("xl/workbook.xml")).getroot()
        return [sheet.attrib["name"] for sheet in workbook.find(XLSX_NS + "sheets")]


def read_sheet(
    path: Path, sheet_name: str | None = None
) -> tuple[str, list[dict[str, str]]]:
    """Read a selected worksheet from an xlsx file without third-party modules."""
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.parse(archive.open("xl/sharedStrings.xml")).getroot()
            for item in root.findall(XLSX_NS + "si"):
                shared_strings.append(
                    "".join(node.text or "" for node in item.iter(XLSX_NS + "t"))
                )

        workbook = ET.parse(archive.open("xl/workbook.xml")).getroot()
        relationships = ET.parse(
            archive.open("xl/_rels/workbook.xml.rels")
        ).getroot()
        relationship_map = {
            node.attrib["Id"]: node.attrib["Target"]
            for node in relationships.findall(PKG_REL_NS + "Relationship")
        }
        sheets = list(workbook.find(XLSX_NS + "sheets"))
        if sheet_name is None:
            sheet = sheets[0]
        else:
            matches = [sheet for sheet in sheets if sheet.attrib["name"] == sheet_name]
            if not matches:
                raise ValueError(f"Worksheet not found: {sheet_name}")
            sheet = matches[0]
        sheet_name = sheet.attrib["name"]
        target = relationship_map[sheet.attrib[REL_NS + "id"]]
        if not target.startswith("xl/"):
            target = "xl/" + target

        raw_rows: list[dict[int, str]] = []
        for _, element in ET.iterparse(archive.open(target), events=("end",)):
            if element.tag != XLSX_NS + "row":
                continue
            row: dict[int, str] = {}
            for cell in element.findall(XLSX_NS + "c"):
                column_number = excel_column_number(cell.attrib["r"])
                cell_type = cell.attrib.get("t")
                value_node = cell.find(XLSX_NS + "v")
                if cell_type == "inlineStr":
                    inline = cell.find(XLSX_NS + "is")
                    value = (
                        "".join(
                            node.text or "" for node in inline.iter(XLSX_NS + "t")
                        )
                        if inline is not None
                        else ""
                    )
                elif value_node is None:
                    value = ""
                elif cell_type == "s":
                    value = shared_strings[int(value_node.text)]
                else:
                    value = value_node.text or ""
                row[column_number] = value
            raw_rows.append(row)
            element.clear()

    if not raw_rows:
        return sheet_name, []
    maximum_column = max(raw_rows[0])
    headers = [text(raw_rows[0].get(i)) for i in range(1, maximum_column + 1)]
    rows: list[dict[str, str]] = []
    for source_row, raw in enumerate(raw_rows[1:], start=2):
        record = {
            header: text(raw.get(i))
            for i, header in enumerate(headers, start=1)
            if header
        }
        record["_source_row"] = str(source_row)
        rows.append(record)
    return sheet_name, rows


def read_first_sheet(path: Path) -> tuple[str, list[dict[str, str]]]:
    """Backward-compatible first-sheet reader."""
    return read_sheet(path)


def remap_columns(
    rows: list[dict[str, str]], column_map: dict[str, str] | None
) -> list[dict[str, str]]:
    if not column_map:
        return rows
    return [
        {
            **{canonical: row.get(source, "") for canonical, source in column_map.items()},
            "_source_row": row["_source_row"],
        }
        for row in rows
    ]


def parse_excel_date(value: str) -> tuple[date | None, str]:
    """Return a date and a parse flag."""
    raw = text(value)
    if not raw:
        return None, "missing"
    try:
        number = float(raw)
        return (datetime(1899, 12, 30) + timedelta(days=number)).date(), "excel_serial"
    except ValueError:
        pass

    prefix = re.match(r"^\s*(\d{1,2}/\d{1,2}/\d{4})", raw)
    if prefix:
        try:
            parsed = datetime.strptime(prefix.group(1), "%m/%d/%Y").date()
            return parsed, "repaired_date_prefix" if prefix.group(1) != raw else "text_date"
        except ValueError:
            return None, "invalid"
    for date_format in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(raw, date_format).date(), "text_date"
        except ValueError:
            continue
    return None, "invalid"


def parse_indicator(value: str) -> int | None:
    raw = text(value)
    if raw in {"0", "0.0"}:
        return 0
    if raw in {"1", "1.0"}:
        return 1
    return None


def normalize_pfr(value: str) -> str:
    normalized = re.sub(r"\s+", " ", text(value).upper())
    metal = re.fullmatch(r"METAL\s*(\d+)", normalized)
    return f"METAL {metal.group(1)}" if metal else normalized


def iso(value: date | None) -> str:
    return value.isoformat() if value else ""


def join_flags(flags: Iterable[str]) -> str:
    return ";".join(sorted({flag for flag in flags if flag}))


def observation_state(
    eggs: int | None, chicks: int | None, status: str, status_raw: str
) -> str:
    if (eggs, chicks) == (1, 1):
        status_states = {
            "A": "status_priority_alive",
            "P": "status_priority_presumed_alive",
            "M": "status_priority_missing",
            "D": "status_priority_dead",
            "L": "status_priority_lay",
            "H": "status_priority_hatch",
            "F": "status_priority_fledge",
        }
        if status in status_states:
            return status_states[status]
        return (
            "status_priority_nonstandard"
            if text(status_raw)
            else "status_priority_missing"
        )
    states = {
        (1, 0): "egg_only",
        (0, 1): "chick_only",
        (0, 0): "neither",
    }
    return states.get((eggs, chicks), "unknown")


def midpoint(start: date | None, end: date | None) -> date | None:
    if start is None or end is None or start > end:
        return None
    return start + timedelta(days=(end - start).days / 2)


def percentile(values: list[float], probability: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * probability
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def descriptive_rows(
    values: list[float], group: str, metric: str, unit: str
) -> dict[str, str]:
    n = len(values)
    mean = statistics.fmean(values) if values else None
    sd = statistics.stdev(values) if n > 1 else None
    se = sd / math.sqrt(n) if sd is not None else None
    return {
        "group": group,
        "metric": metric,
        "unit": unit,
        "n": str(n),
        "numerator": "",
        "denominator": "",
        "mean": format_number(mean),
        "se": format_number(se),
        "q02_5": format_number(percentile(values, 0.025)),
        "q25": format_number(percentile(values, 0.25)),
        "median": format_number(percentile(values, 0.50)),
        "q75": format_number(percentile(values, 0.75)),
        "q97_5": format_number(percentile(values, 0.975)),
    }


def proportion_row(
    numerator: int, denominator: int, group: str, metric: str
) -> dict[str, str]:
    proportion = numerator / denominator if denominator else None
    se = (
        math.sqrt(proportion * (1 - proportion) / denominator)
        if proportion is not None
        else None
    )
    return {
        "group": group,
        "metric": metric,
        "unit": "proportion",
        "n": str(denominator),
        "numerator": str(numerator),
        "denominator": str(denominator),
        "mean": format_number(proportion),
        "se": format_number(se),
        "q02_5": "",
        "q25": "",
        "median": "",
        "q75": "",
        "q97_5": "",
    }


def date_summary_row(
    values: list[date], group: str, event: str, method: str
) -> dict[str, str]:
    ordinals = [float(value.toordinal()) for value in values]
    numeric = descriptive_rows(ordinals, group, f"{event}_{method}", "date")

    def ordinal_date(value: float | None) -> str:
        if value is None:
            return ""
        return date.fromordinal(round(value)).isoformat()

    return {
        "group": group,
        "event": event,
        "method": method,
        "n": numeric["n"],
        "mean_date": ordinal_date(float(numeric["mean"])) if numeric["mean"] else "",
        "se_days": numeric["se"],
        "q02_5": ordinal_date(float(numeric["q02_5"])) if numeric["q02_5"] else "",
        "q25": ordinal_date(float(numeric["q25"])) if numeric["q25"] else "",
        "median": ordinal_date(float(numeric["median"])) if numeric["median"] else "",
        "q75": ordinal_date(float(numeric["q75"])) if numeric["q75"] else "",
        "q97_5": ordinal_date(float(numeric["q97_5"])) if numeric["q97_5"] else "",
        "earliest": iso(min(values) if values else None),
        "latest": iso(max(values) if values else None),
    }


def format_number(value: float | None) -> str:
    return "" if value is None else f"{value:.6f}".rstrip("0").rstrip(".")


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def clean_productivity(
    path: Path,
    sheet_name: str | None = None,
    column_map: dict[str, str] | None = None,
    year_filter: int | None = None,
    species_default: str = SPECIES,
) -> tuple[list[dict[str, Any]], str, list[dict[str, str]]]:
    sheet_name, raw_rows = read_sheet(path, sheet_name)
    raw_rows = remap_columns(raw_rows, column_map)
    cleaned: list[dict[str, Any]] = []
    qc: list[dict[str, str]] = []
    for raw in raw_rows:
        source_row = int(raw["_source_row"])
        observed_date, date_flag = parse_excel_date(raw.get("DATE", ""))
        if observed_date and year_filter is not None and observed_date.year != year_filter:
            continue
        year = observed_date.year if observed_date else (year_filter or 0)
        eggs = parse_indicator(raw.get("Eggs", ""))
        chicks = parse_indicator(raw.get("Chicks", ""))
        status_raw = text(raw.get("Status", ""))
        status = status_raw.upper()
        status_valid = status in VALID_STATUS
        species = text(raw.get("Species", "")) or text(species_default)
        species = species.upper()
        plot = text(raw.get("PLOT", ""))
        nest_id = text(raw.get("Nest#", "")).upper()
        slot_label = text(raw.get("A or B chick", "")).upper()
        pfr_raw = text(raw.get("PFR", ""))
        pfr = normalize_pfr(pfr_raw)
        pfr_correction = ""
        if species == SPECIES and year == 2025 and pfr == "XN1" and nest_id == "R1897":
            pfr = ""
            pfr_correction = "excluded_wrong_nest_assignment;correct_nest=R897"
        location = text(raw.get("LOCATION", "")).upper()
        nest_key = f"{year}|{species}|{plot}|{nest_id}"
        slot_key = f"{nest_key}|{slot_label}"
        flags: list[str] = []
        if date_flag in {"missing", "invalid", "repaired_date_prefix"}:
            flags.append(f"date_{date_flag}")
        if eggs is None:
            flags.append("invalid_eggs")
        if chicks is None:
            flags.append("invalid_chicks")
        if not status:
            flags.append("missing_status")
        elif not status_valid:
            flags.append("nonstandard_status")
        if not plot:
            flags.append("missing_plot")
        if not nest_id:
            flags.append("missing_nest_id")
        if slot_label not in {"A", "B", "C"}:
            flags.append("invalid_slot_label")
        if not pfr:
            flags.append("missing_pfr")
        elif pfr == "UNBANDED":
            flags.append("unbanded")
        if pfr_correction:
            flags.append("known_pfr_correction")
        if location in {"", "#N/A"}:
            flags.append("invalid_location")

        record = {
            "productivity_record_id": f"PROD-{source_row:05d}",
            "year": year,
            "species": species,
            "observation_date": iso(observed_date),
            "date_raw": raw.get("DATE", ""),
            "date_parse_flag": date_flag,
            "observer": text(raw.get("OBSERVER", "")),
            "plot": plot,
            "nest_id": nest_id,
            "nest_key": nest_key,
            "slot_label": slot_label,
            "slot_key": slot_key,
            "eggs": "" if eggs is None else eggs,
            "eggs_raw": raw.get("Eggs", ""),
            "chicks": "" if chicks is None else chicks,
            "chicks_raw": raw.get("Chicks", ""),
            "observation_state": observation_state(eggs, chicks, status, status_raw),
            "status": status if status_valid else "",
            "status_raw": status_raw,
            "status_valid": status_valid,
            "pfr": "" if pfr in PLACEHOLDER_PFR else pfr,
            "pfr_raw": pfr_raw,
            "pfr_correction": pfr_correction,
            "pfr_placeholder": pfr in PLACEHOLDER_PFR,
            "location": location,
            "array": location if plot == "Exp Array" and location in set("BCDEFGHI") else "",
            "box": "",
            "telemetry_flag": "",
            "notes": text(raw.get("Notes", "")),
            "source_file": path.name,
            "source_sheet": sheet_name,
            "source_row": source_row,
            "qc_flags": join_flags(flags),
            "_date": observed_date,
            "_eggs": eggs,
            "_chicks": chicks,
        }
        cleaned.append(record)
        for flag in flags:
            qc.append(
                qc_record(
                    "Productivity_Master",
                    record["productivity_record_id"],
                    flag,
                    source_row,
                    f"{slot_key}: {flag}",
                )
            )
    return cleaned, sheet_name, qc


def clean_resights(
    path: Path,
    sheet_name: str | None = None,
    column_map: dict[str, str] | None = None,
    year_filter: int | None = None,
) -> tuple[list[dict[str, Any]], str, list[dict[str, str]]]:
    sheet_name, raw_rows = read_sheet(path, sheet_name)
    raw_rows = remap_columns(raw_rows, column_map)
    cleaned: list[dict[str, Any]] = []
    qc: list[dict[str, str]] = []
    for raw in raw_rows:
        species = text(raw.get("Species", "")).upper()
        if not species:
            species = SPECIES
        observed_date, date_flag = parse_excel_date(raw.get("Favorite Date", ""))
        if observed_date is None:
            continue
        year = observed_date.year
        if year_filter is not None and year != year_filter:
            continue
        source_row = int(raw["_source_row"])
        combo_raw = text(raw.get("Combo", ""))
        combo = normalize_pfr(combo_raw)
        age = text(raw.get("Age", "")).title()
        fledged_raw = text(raw.get("Fledged?", ""))
        fledged = {"YES!": "yes", "NO": "no", "": ""}.get(fledged_raw.upper(), "invalid")
        flags: list[str] = []
        if date_flag in {"invalid", "repaired_date_prefix"}:
            flags.append(f"date_{date_flag}")
        if not combo or combo == "UNBANDED":
            flags.append("invalid_combo")
        if fledged == "invalid":
            flags.append("invalid_fledged_value")
        record = {
            "resight_record_id": f"RESIGHT-{source_row:05d}",
            "year": year,
            "species": species,
            "observation_date": iso(observed_date),
            "date_raw": raw.get("Favorite Date", ""),
            "date_parse_flag": date_flag,
            "combo": "" if combo in PLACEHOLDER_PFR else combo,
            "combo_raw": combo_raw,
            "age": age,
            "nest_id_reported": text(
                raw.get('NEST NUMBER (ROST NEST STARTS WITH "R")', "")
            ).upper(),
            "fledged": fledged,
            "fledged_raw": fledged_raw,
            "location": text(raw.get("Location", "")),
            "observer": text(raw.get("Favorite Observer (who is not Joan)", "")),
            "notes": text(raw.get("Notes - please be brief", "")),
            "source_file": path.name,
            "source_sheet": sheet_name,
            "source_row": source_row,
            "eligible_for_verification": bool(
                combo not in PLACEHOLDER_PFR and fledged in {"", "yes", "no"}
            ),
            "qc_flags": join_flags(flags),
            "_date": observed_date,
        }
        cleaned.append(record)
        for flag in flags:
            qc.append(
                qc_record(
                    "cleaned_resight_observations",
                    record["resight_record_id"],
                    flag,
                    source_row,
                    f"{combo or combo_raw}: {flag}",
                )
            )
    return cleaned, sheet_name, qc


def qc_record(
    layer: str, record_id: str, issue: str, source_row: int | str, detail: str
) -> dict[str, str]:
    return {
        "layer": layer,
        "record_id": record_id,
        "issue": issue,
        "source_row": str(source_row),
        "detail": detail,
    }


def derive_tables(
    productivity: list[dict[str, Any]],
    resights: list[dict[str, Any]],
    initial_qc: list[dict[str, str]],
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, str]],
]:
    qc = list(initial_qc)
    productivity.sort(
        key=lambda row: (
            row["slot_key"],
            row["_date"] or date.max,
            int(row["source_row"]),
        )
    )
    by_slot: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_nest: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in productivity:
        by_slot[row["slot_key"]].append(row)
        by_nest[row["nest_key"]].append(row)

    same_slot_date: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in productivity:
        same_slot_date[(row["slot_key"], row["observation_date"])].append(row)
    for (slot_key, observed), records in same_slot_date.items():
        states = {row["observation_state"] for row in records}
        if len(records) > 1:
            issue = "duplicate_slot_date_conflict" if len(states) > 1 else "duplicate_slot_date"
            qc.append(
                qc_record(
                    "Productivity_Master",
                    slot_key,
                    issue,
                    ",".join(str(row["source_row"]) for row in records),
                    f"{observed}: {sorted(states)}",
                )
            )

    resights_by_pfr: dict[tuple[int, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in resights:
        if row["eligible_for_verification"] and row["combo"]:
            resights_by_pfr[(int(row["year"]), row["species"], row["combo"])].append(row)

    pfr_to_slots: dict[tuple[int, str, str], set[str]] = defaultdict(set)
    for row in productivity:
        if row["pfr"]:
            pfr_to_slots[(int(row["year"]), row["species"], row["pfr"])].add(row["slot_key"])
    for (year, species, pfr), slots in pfr_to_slots.items():
        if len(slots) > 1:
            qc.append(
                qc_record(
                    "chick_summary",
                    f"{year}|{species}|{pfr}",
                    "pfr_multiple_slots",
                    "",
                    ";".join(sorted(slots)),
                )
            )

    chick_rows: list[dict[str, Any]] = []
    chronology_rows: list[dict[str, Any]] = []
    for slot_key, records in by_slot.items():
        valid_dates = [row["_date"] for row in records if row["_date"]]
        plot = records[0]["plot"]
        year = int(records[0]["year"])
        species = records[0]["species"]
        nest_key = records[0]["nest_key"]
        nest_id = records[0]["nest_id"]
        slot_label = records[0]["slot_label"]
        pfr_values = sorted({row["pfr"] for row in records if row["pfr"]})
        pfr = pfr_values[-1] if len(pfr_values) == 1 else ""
        if len(pfr_values) > 1:
            qc.append(
                qc_record(
                    "chick_summary",
                    slot_key,
                    "slot_multiple_pfr",
                    "",
                    ";".join(pfr_values),
                )
            )

        lay_dates = [row["_date"] for row in records if row["status"] == "L" and row["_date"]]
        first_lay_observed = min(lay_dates) if lay_dates else None
        pre_lay_dates = [
            row["_date"]
            for row in records
            if row["_date"] and first_lay_observed and row["_date"] < first_lay_observed
        ]
        last_pre_lay = max(pre_lay_dates) if pre_lay_dates else None
        lay_midpoint = midpoint(last_pre_lay, first_lay_observed)
        valid_h_dates = [
            row["_date"] for row in records if row["status"] == "H" and row["_date"]
        ]
        chick_dates = [
            row["_date"] for row in records if row["_chicks"] == 1 and row["_date"]
        ]
        first_chick_date = min(chick_dates) if chick_dates else None
        first_hatch_observed = min(valid_h_dates + chick_dates) if (valid_h_dates or chick_dates) else None
        pre_hatch_dates = [
            row["_date"]
            for row in records
            if row["_date"]
            and first_hatch_observed
            and row["_date"] < first_hatch_observed
            and row["_eggs"] == 1
            and row["_chicks"] == 0
        ]
        last_pre_hatch = max(pre_hatch_dates) if pre_hatch_dates else None
        hatch_midpoint = midpoint(last_pre_hatch, first_hatch_observed)

        fledge_dates = [
            row["_date"] for row in records if row["status"] == "F" and row["_date"]
        ]
        first_f_status = min(fledge_dates) if fledge_dates else None
        pre_fledge_dates = [
            row["_date"]
            for row in records
            if row["_date"]
            and first_f_status
            and row["_date"] < first_f_status
            and row["status"] != "F"
        ]
        last_pre_fledge = max(pre_fledge_dates) if pre_fledge_dates else None

        pfr_ambiguous = bool(
            pfr and len(pfr_to_slots.get((year, species, pfr), set())) > 1
        )
        eligible_resights = (
            []
            if pfr_ambiguous
            else list(resights_by_pfr.get((year, species, pfr), []))
        )
        resight_yes = any(row["fledged"] == "yes" for row in eligible_resights)
        resight_no = any(row["fledged"] == "no" for row in eligible_resights)
        productivity_f = bool(first_f_status)
        reviewer_confirmed = (year, pfr) in REVIEW_CONFIRMED_FLEDGLINGS
        yes_resight_dates = [
            row["_date"]
            for row in eligible_resights
            if row["fledged"] == "yes" and row["_date"]
        ]
        first_resight_yes = min(yes_resight_dates) if yes_resight_dates else None
        verified_fledged = bool(productivity_f or resight_yes or reviewer_confirmed)
        verified_fledge_dates = [
            value for value in (first_f_status, first_resight_yes) if value
        ]
        verified_fledge_date = (
            min(verified_fledge_dates) if verified_fledge_dates else None
        )
        productivity_pre_verified = [
            row["_date"]
            for row in records
            if row["_date"]
            and verified_fledge_date
            and row["_date"] < verified_fledge_date
            and row["status"] != "F"
        ]
        resight_pre_verified = [
            row["_date"]
            for row in eligible_resights
            if row["_date"]
            and verified_fledge_date
            and row["_date"] < verified_fledge_date
            and row["fledged"] == "no"
        ]
        fledge_interval_start = max(
            productivity_pre_verified + resight_pre_verified, default=None
        )
        if productivity_f and resight_yes:
            verification_basis = "both"
        elif productivity_f:
            verification_basis = "productivity_status_f"
        elif resight_yes:
            verification_basis = "resight_yes"
        elif reviewer_confirmed:
            verification_basis = "reviewer_confirmation"
        else:
            verification_basis = ""
        known_dead = any(row["status"] == "D" for row in records) and not verified_fledged
        hatched = bool(valid_h_dates or chick_dates or verified_fledged)
        has_egg = any(row["_eggs"] == 1 for row in records)
        if verified_fledged:
            outcome = "verified_fledged"
        elif known_dead:
            outcome = "known_dead"
        elif hatched:
            outcome = "unresolved"
        else:
            outcome = "not_hatched"

        chick_flags: list[str] = []
        chick_sequence = [row["_chicks"] for row in records if row["_chicks"] is not None]
        if 1 in chick_sequence:
            first_one = chick_sequence.index(1)
            if 0 in chick_sequence[first_one + 1 :]:
                chick_flags.append("chick_to_zero_reversal")
        if first_hatch_observed and not valid_h_dates:
            chick_flags.append("state_hatch_without_status_h")
        if valid_h_dates and not chick_dates:
            chick_flags.append("status_h_without_chick_state")
        if pfr and (year, species, pfr) not in resights_by_pfr:
            chick_flags.append("pfr_without_resight")
        if reviewer_confirmed:
            chick_flags.append("reviewer_confirmed_fledge")
        if pfr_ambiguous:
            chick_flags.append("ambiguous_pfr_multiple_slots")
        if not pfr:
            chick_flags.append("no_unique_pfr")
        if resight_no and resight_yes:
            no_dates = [
                row["_date"] for row in eligible_resights if row["fledged"] == "no"
            ]
            if yes_resight_dates and no_dates and max(no_dates) > min(yes_resight_dates):
                chick_flags.append("resight_no_after_yes")
        if first_f_status:
            death_dates = [
                row["_date"] for row in records if row["status"] == "D" and row["_date"]
            ]
            if death_dates and min(death_dates) < first_f_status:
                chick_flags.append("death_before_fledge")

        for flag in chick_flags:
            qc.append(qc_record("chick_summary", slot_key, flag, "", flag))

        chick_row = {
            "chick_id": f"{year}|{species}|{pfr}" if pfr else slot_key,
            "year": year,
            "species": species,
            "plot": plot,
            "nest_id": nest_id,
            "nest_key": nest_key,
            "slot_label": slot_label,
            "slot_key": slot_key,
            "pfr": pfr,
            "pfr_values": ";".join(pfr_values),
            "first_observation_date": iso(min(valid_dates) if valid_dates else None),
            "last_observation_date": iso(max(valid_dates) if valid_dates else None),
            "has_egg_evidence": has_egg,
            "hatched": hatched,
            "first_chick_date": iso(first_chick_date),
            "status_h_date": iso(min(valid_h_dates) if valid_h_dates else None),
            "productivity_fledge_evidence": productivity_f,
            "productivity_fledge_date": iso(first_f_status),
            "resight_yes_evidence": resight_yes,
            "resight_yes_date": iso(first_resight_yes),
            "resight_no_evidence": resight_no,
            "verified_fledged": verified_fledged,
            "verified_fledge_date": iso(verified_fledge_date),
            "fledge_verification_basis": verification_basis,
            "known_dead": known_dead,
            "outcome": outcome,
            "qc_flags": join_flags(chick_flags),
        }
        chick_rows.append(chick_row)

        chronology_rows.append(
            {
                "year": year,
                "species": species,
                "plot": plot,
                "nest_id": nest_id,
                "nest_key": nest_key,
                "slot_label": slot_label,
                "slot_key": slot_key,
                "pfr": pfr,
                "lay_first_observed": iso(first_lay_observed),
                "lay_interval_start": iso(last_pre_lay),
                "lay_midpoint": iso(lay_midpoint),
                "hatch_first_observed": iso(first_hatch_observed),
                "hatch_interval_start": iso(last_pre_hatch),
                "hatch_midpoint": iso(hatch_midpoint),
                "fledge_first_observed": iso(verified_fledge_date),
                "fledge_interval_start": iso(fledge_interval_start),
                "fledge_midpoint": iso(
                    midpoint(fledge_interval_start, verified_fledge_date)
                ),
                "verified_fledged": verified_fledged,
                "fledge_verification_basis": verification_basis,
            }
        )

    nest_rows: list[dict[str, Any]] = []
    for nest_key, records in by_nest.items():
        plots = sorted({row["plot"] for row in records})
        locations = sorted(
            {row["location"] for row in records if row["location"] not in {"", "#N/A"}}
        )
        arrays = sorted({row["array"] for row in records if row["array"]})
        nest_slots = [
            row for row in chick_rows if row["nest_key"] == nest_key
        ]
        flags: list[str] = []
        if len(plots) > 1:
            flags.append("multiple_plots")
        if len(locations) > 1:
            flags.append("multiple_locations")
        if len(arrays) > 1:
            flags.append("multiple_arrays")
        for flag in flags:
            qc.append(
                qc_record(
                    "Nest_Lookup",
                    nest_key,
                    flag,
                    "",
                    ";".join(locations if "location" in flag else arrays),
                )
            )
        clutch_size = sum(bool(row["has_egg_evidence"]) for row in nest_slots)
        hatched_count = sum(bool(row["hatched"]) for row in nest_slots)
        verified_count = sum(bool(row["verified_fledged"]) for row in nest_slots)
        dead_count = sum(bool(row["known_dead"]) for row in nest_slots)
        nest_rows.append(
            {
                "year": int(records[0]["year"]),
                "species": records[0]["species"],
                "plot": plots[0] if len(plots) == 1 else ";".join(plots),
                "nest_id": records[0]["nest_id"],
                "nest_key": nest_key,
                "location": locations[0] if len(locations) == 1 else "",
                "location_values": ";".join(locations),
                "array": arrays[0] if len(arrays) == 1 else "",
                "box": "",
                "treatment_flag": "",
                "telemetry_flag": "",
                "clutch_size": clutch_size,
                "hatched_chicks": hatched_count,
                "verified_fledglings": verified_count,
                "known_dead_chicks": dead_count,
                "notes": "",
                "qc_flags": join_flags(flags),
            }
        )

    summary_rows: list[dict[str, Any]] = []
    groups: list[tuple[int, str, str, list[dict[str, Any]], list[dict[str, Any]]]] = []
    year_species = sorted({(int(row["year"]), row["species"]) for row in nest_rows if int(row["year"])})
    for year, species in year_species:
        year_nests = [row for row in nest_rows if int(row["year"]) == year and row["species"] == species]
        year_chicks = [row for row in chick_rows if int(row["year"]) == year and row["species"] == species]
        groups.append((year, species, "Overall", year_nests, year_chicks))
        for plot in sorted({row["plot"] for row in year_nests}):
            groups.append(
                (
                    year, species,
                    plot,
                    [row for row in year_nests if row["plot"] == plot],
                    [row for row in year_chicks if row["plot"] == plot],
                )
            )
    for year, species, group, group_nests, group_chicks in groups:
        summary_rows.append(
            {"year": year, "species": species, **descriptive_rows(
                [float(row["clutch_size"]) for row in group_nests],
                group,
                "clutch_size",
                "eggs_per_nest",
            )}
        )
        hatched = [row for row in group_chicks if row["hatched"]]
        verified = sum(bool(row["verified_fledged"]) for row in hatched)
        dead = sum(bool(row["known_dead"]) for row in hatched)
        summary_rows.append(
            {"year": year, "species": species, **proportion_row(verified, len(hatched), group, "apparent_verified_fledge_rate")}
        )
        summary_rows.append(
            {"year": year, "species": species, **proportion_row(
                verified,
                verified + dead,
                group,
                "resolved_outcome_fledge_rate",
            )}
        )
        summary_rows.append(
            {
                "year": year, "species": species, **descriptive_rows(
                    [float(row["hatched_chicks"]) for row in group_nests],
                    group,
                    "hatched_chicks_per_nest",
                    "chicks_per_nest",
                )
            }
        )
        summary_rows.append(
            {
                "year": year, "species": species, **descriptive_rows(
                    [float(row["verified_fledglings"]) for row in group_nests],
                    group,
                    "verified_fledglings_per_nest",
                    "fledglings_per_nest",
                )
            }
        )

    chronology_summary_rows: list[dict[str, Any]] = []
    chronology_groups: list[tuple[int, str, str, list[dict[str, Any]]]] = []
    chronology_year_species = sorted({(int(row["year"]), row["species"]) for row in chronology_rows if int(row["year"])})
    for year, species in chronology_year_species:
        year_rows = [row for row in chronology_rows if int(row["year"]) == year and row["species"] == species]
        chronology_groups.append((year, species, "Overall", year_rows))
        for plot in sorted({row["plot"] for row in year_rows}):
            chronology_groups.append(
                (year, species, plot, [row for row in year_rows if row["plot"] == plot])
            )
    date_fields = [
        ("lay", "first_observed", "lay_first_observed"),
        ("lay", "midpoint", "lay_midpoint"),
        ("hatch", "first_observed", "hatch_first_observed"),
        ("hatch", "midpoint", "hatch_midpoint"),
        ("fledge", "first_observed", "fledge_first_observed"),
        ("fledge", "midpoint", "fledge_midpoint"),
    ]
    for year, species, group, group_rows in chronology_groups:
        for event, method, field in date_fields:
            values = [
                datetime.strptime(row[field], "%Y-%m-%d").date()
                for row in group_rows
                if row[field]
            ]
            chronology_summary_rows.append(
                {"year": year, "species": species, **date_summary_row(values, group, event, method)}
            )

    unmatched_resight_pfr = sorted(
        {
            (int(row["year"]), row["species"], row["combo"])
            for row in resights
            if row["eligible_for_verification"]
            and row["combo"]
            and row["age"] in {"", "Chick"}
            and (int(row["year"]), row["species"], row["combo"]) not in pfr_to_slots
        }
    )
    for year, species, pfr in unmatched_resight_pfr:
        qc.append(
            qc_record(
                "cleaned_resight_observations",
                f"{year}|{species}|{pfr}",
                "resight_without_productivity_pfr",
                "",
                pfr,
            )
        )

    return (
        nest_rows,
        chick_rows,
        chronology_rows,
        summary_rows,
        chronology_summary_rows,
        qc,
    )


PRODUCTIVITY_FIELDS = [
    "productivity_record_id",
    "year",
    "species",
    "observation_date",
    "date_raw",
    "date_parse_flag",
    "observer",
    "plot",
    "nest_id",
    "nest_key",
    "slot_label",
    "slot_key",
    "eggs",
    "eggs_raw",
    "chicks",
    "chicks_raw",
    "observation_state",
    "status",
    "status_raw",
    "status_valid",
    "pfr",
    "pfr_raw",
    "pfr_correction",
    "pfr_placeholder",
    "location",
    "array",
    "box",
    "telemetry_flag",
    "notes",
    "source_file",
    "source_sheet",
    "source_row",
    "qc_flags",
]

RESIGHT_FIELDS = [
    "resight_record_id",
    "year",
    "species",
    "observation_date",
    "date_raw",
    "date_parse_flag",
    "combo",
    "combo_raw",
    "age",
    "nest_id_reported",
    "fledged",
    "fledged_raw",
    "location",
    "observer",
    "notes",
    "source_file",
    "source_sheet",
    "source_row",
    "eligible_for_verification",
    "qc_flags",
]

NEST_FIELDS = [
    "year",
    "species",
    "plot",
    "nest_id",
    "nest_key",
    "location",
    "location_values",
    "array",
    "box",
    "treatment_flag",
    "telemetry_flag",
    "clutch_size",
    "hatched_chicks",
    "verified_fledglings",
    "known_dead_chicks",
    "notes",
    "qc_flags",
]

CHICK_FIELDS = [
    "chick_id",
    "year",
    "species",
    "plot",
    "nest_id",
    "nest_key",
    "slot_label",
    "slot_key",
    "pfr",
    "pfr_values",
    "first_observation_date",
    "last_observation_date",
    "has_egg_evidence",
    "hatched",
    "first_chick_date",
    "status_h_date",
    "productivity_fledge_evidence",
    "productivity_fledge_date",
    "resight_yes_evidence",
    "resight_yes_date",
    "resight_no_evidence",
    "verified_fledged",
    "verified_fledge_date",
    "fledge_verification_basis",
    "known_dead",
    "outcome",
    "qc_flags",
]

CHRONOLOGY_FIELDS = [
    "year",
    "species",
    "plot",
    "nest_id",
    "nest_key",
    "slot_label",
    "slot_key",
    "pfr",
    "lay_first_observed",
    "lay_interval_start",
    "lay_midpoint",
    "hatch_first_observed",
    "hatch_interval_start",
    "hatch_midpoint",
    "fledge_first_observed",
    "fledge_interval_start",
    "fledge_midpoint",
    "verified_fledged",
    "fledge_verification_basis",
]

SUMMARY_FIELDS = [
    "year",
    "species",
    "group",
    "metric",
    "unit",
    "n",
    "numerator",
    "denominator",
    "mean",
    "se",
    "q02_5",
    "q25",
    "median",
    "q75",
    "q97_5",
]

CHRONOLOGY_SUMMARY_FIELDS = [
    "year",
    "species",
    "group",
    "event",
    "method",
    "n",
    "mean_date",
    "se_days",
    "q02_5",
    "q25",
    "median",
    "q75",
    "q97_5",
    "earliest",
    "latest",
]

QC_FIELDS = ["year", "species", "layer", "record_id", "issue", "source_row", "detail"]


def assign_qc_years(
    qc: list[dict[str, Any]],
    productivity: list[dict[str, Any]],
    resights: list[dict[str, Any]],
) -> None:
    productivity_years = {
        row["productivity_record_id"]: int(row["year"]) for row in productivity
    }
    resight_years = {row["resight_record_id"]: int(row["year"]) for row in resights}
    productivity_species = {
        row["productivity_record_id"]: row["species"] for row in productivity
    }
    resight_species = {
        row["resight_record_id"]: row["species"] for row in resights
    }
    for row in qc:
        year_match = re.search(
            r"(20\d{2})\|", row["record_id"] + " " + row["detail"]
        )
        row["year"] = (
            productivity_years.get(row["record_id"])
            or resight_years.get(row["record_id"])
            or (int(year_match.group(1)) if year_match else "")
        )
        key_match = re.search(r"20\d{2}\|([^|]+)\|", row["record_id"] + " " + row["detail"])
        row["species"] = (
            productivity_species.get(row["record_id"])
            or resight_species.get(row["record_id"])
            or (key_match.group(1) if key_match else "")
        )


def analyze_workbooks(
    productivity_path: Path,
    productivity_sheet: str | None = None,
    productivity_map: dict[str, str] | None = None,
    resight_path: Path | None = None,
    resight_sheet: str | None = None,
    resight_map: dict[str, str] | None = None,
    species_default: str = SPECIES,
) -> dict[str, list[dict[str, Any]]]:
    """Analyze uploaded workbooks without writing outputs to disk."""
    productivity, _, productivity_qc = clean_productivity(
        productivity_path, productivity_sheet, productivity_map, species_default=species_default
    )
    invalid_date_qc = [
        qc_record(
            "Productivity_Master",
            row["productivity_record_id"],
            "excluded_missing_analysis_year",
            row["source_row"],
            "A valid observation date is required to assign an analysis year",
        )
        for row in productivity
        if not int(row["year"])
    ]
    productivity = [row for row in productivity if int(row["year"])]
    if not productivity:
        raise ValueError("No productivity rows contain a valid analysis year")

    if resight_path is None:
        resights: list[dict[str, Any]] = []
        resight_qc: list[dict[str, str]] = []
    else:
        resights, _, resight_qc = clean_resights(
            resight_path, resight_sheet, resight_map
        )

    nests, chicks, chronology, summaries, chronology_summaries, qc = derive_tables(
        productivity,
        resights,
        productivity_qc + invalid_date_qc + resight_qc,
    )
    assign_qc_years(qc, productivity, resights)
    return {
        "productivity": productivity,
        "resights": resights,
        "nests": nests,
        "chicks": chicks,
        "chronology": chronology,
        "summary": summaries,
        "chronology_summary": chronology_summaries,
        "qc": qc,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--productivity",
        type=Path,
        default=Path("PRODUCTIVITY 2025.xlsx"),
    )
    parser.add_argument(
        "--resights",
        type=Path,
        default=Path("RESIGHTS 2024 AND 2025.xlsx"),
    )
    parser.add_argument("--year", type=int, default=DEFAULT_YEAR)
    parser.add_argument("--output-dir", type=Path, default=Path("outputs"))
    args = parser.parse_args()

    productivity, _, productivity_qc = clean_productivity(
        args.productivity, year_filter=args.year
    )
    resights, _, resight_qc = clean_resights(args.resights, year_filter=args.year)
    nests, chicks, chronology, summaries, chronology_summaries, qc = derive_tables(
        productivity, resights, productivity_qc + resight_qc
    )
    assign_qc_years(qc, productivity, resights)

    output = args.output_dir
    analysis_output = output / "Analysis_Output"
    write_csv(output / "Productivity_Master.csv", productivity, PRODUCTIVITY_FIELDS)
    write_csv(
        output / "cleaned_resight_observations.csv", resights, RESIGHT_FIELDS
    )
    write_csv(output / "Nest_Lookup.csv", nests, NEST_FIELDS)
    write_csv(output / "chick_summary.csv", chicks, CHICK_FIELDS)
    write_csv(output / "breeding_chronology.csv", chronology, CHRONOLOGY_FIELDS)
    write_csv(
        analysis_output / "productivity_summary.csv", summaries, SUMMARY_FIELDS
    )
    write_csv(
        analysis_output / "chronology_summary.csv",
        chronology_summaries,
        CHRONOLOGY_SUMMARY_FIELDS,
    )
    write_csv(output / "quality_control_exceptions.csv", qc, QC_FIELDS)

    print(f"Wrote {len(productivity):,} productivity observations")
    print(f"Wrote {len(resights):,} {args.year} ROST resight observations")
    print(f"Wrote {len(nests):,} nests and {len(chicks):,} nest slots")
    print(f"Wrote {len(qc):,} QC exceptions")
    print(f"Outputs: {output.resolve()}")


if __name__ == "__main__":
    main()
