"""Load and validate the analysis outputs used by the Shiny app."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
OUTPUTS = ROOT / "outputs"


@dataclass(frozen=True)
class AppData:
    summary: pd.DataFrame
    chronology_summary: pd.DataFrame
    nests: pd.DataFrame
    chicks: pd.DataFrame
    chronology: pd.DataFrame
    qc: pd.DataFrame

    @property
    def years(self) -> list[int]:
        return sorted(int(year) for year in self.chicks["year"].dropna().unique())

    def for_year(self, year: int) -> "AppData":
        def subset(frame: pd.DataFrame) -> pd.DataFrame:
            if "year" not in frame.columns:
                return frame.copy()
            years = pd.to_numeric(frame["year"], errors="coerce")
            if frame is self.qc:
                return frame.loc[(years == int(year)) | years.isna()].copy()
            return frame.loc[years == int(year)].copy()

        return AppData(
            summary=subset(self.summary),
            chronology_summary=subset(self.chronology_summary),
            nests=subset(self.nests),
            chicks=subset(self.chicks),
            chronology=subset(self.chronology),
            qc=self.qc.copy(),
        )

    @property
    def banded_chicks(self) -> pd.DataFrame:
        banded = self.chicks.loc[self.chicks["pfr"].fillna("").str.strip().ne("")].copy()
        banded["fledged"] = banded["verified_fledged"].map({True: "Yes", False: "No"})
        banded["outcome"] = banded["outcome"].str.replace("_", " ").str.title()
        banded["verification_basis"] = (
            banded["fledge_verification_basis"]
            .fillna("")
            .replace({"both": "Both", "productivity_status_f": "Status = F"})
        )
        return banded.rename(
            columns={
                "pfr": "band_id",
                "nest_id": "nest",
                "slot_label": "slot",
                "verified_fledge_date": "verified_date",
            }
        )[
            [
                "band_id",
                "plot",
                "nest",
                "slot",
                "fledged",
                "outcome",
                "verified_date",
                "verification_basis",
            ]
        ].sort_values(["plot", "nest", "slot"])


FILES = {
    "summary": OUTPUTS / "Analysis_Output" / "productivity_summary.csv",
    "chronology_summary": OUTPUTS / "Analysis_Output" / "chronology_summary.csv",
    "nests": OUTPUTS / "Nest_Lookup.csv",
    "chicks": OUTPUTS / "chick_summary.csv",
    "chronology": OUTPUTS / "breeding_chronology.csv",
    "qc": OUTPUTS / "quality_control_exceptions.csv",
}

REQUIRED_COLUMNS = {
    "summary": {"group", "metric", "mean", "numerator", "denominator"},
    "chronology_summary": {"group", "event", "method", "median", "n"},
    "nests": {"plot", "nest_id", "clutch_size", "hatched_chicks", "verified_fledglings"},
    "chicks": {
        "plot",
        "nest_id",
        "slot_label",
        "pfr",
        "hatched",
        "verified_fledged",
        "verified_fledge_date",
        "fledge_verification_basis",
        "outcome",
    },
    "chronology": {"plot", "lay_first_observed", "hatch_midpoint", "fledge_midpoint"},
    "qc": {"layer", "record_id", "issue", "detail"},
}


def load_data() -> AppData:
    missing_files = [str(path) for path in FILES.values() if not path.exists()]
    if missing_files:
        raise FileNotFoundError("Missing required analysis files: " + ", ".join(missing_files))

    frames = {name: pd.read_csv(path) for name, path in FILES.items()}
    data = prepare_data(frames)
    validate_acceptance_totals(data)
    return data


def from_analysis(analysis: dict[str, list[dict[str, Any]]]) -> AppData:
    frames = {
        name: pd.DataFrame(analysis[name])
        for name in (
            "summary",
            "chronology_summary",
            "nests",
            "chicks",
            "chronology",
            "qc",
        )
    }
    return prepare_data(frames)


def prepare_data(frames: dict[str, pd.DataFrame]) -> AppData:
    for name, required in REQUIRED_COLUMNS.items():
        missing = required - set(frames[name].columns)
        if missing:
            raise ValueError(f"{name} data is missing columns: {sorted(missing)}")

    for column in ("hatched", "verified_fledged", "known_dead"):
        if column in frames["chicks"]:
            if frames["chicks"][column].dtype != bool:
                frames["chicks"][column] = frames["chicks"][column].astype(str).eq("True")

    return AppData(**frames)


def validate_acceptance_totals(data: AppData) -> None:
    expected = {
        "nests": (len(data.nests), 114),
        "slots": (len(data.chicks), 204),
        "hatched": (int(data.chicks["hatched"].sum()), 156),
        "verified": (int(data.chicks["verified_fledged"].sum()), 104),
        "banded": (len(data.banded_chicks), 116),
        "banded verified": ((data.banded_chicks["fledged"] == "Yes").sum(), 103),
    }
    failures = [f"{name}: got {actual}, expected {wanted}" for name, (actual, wanted) in expected.items() if actual != wanted]
    if failures:
        raise ValueError("Analysis acceptance totals failed: " + "; ".join(failures))
