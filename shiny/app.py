"""Session-based, multi-year dashboard for ROST productivity workbooks."""

from __future__ import annotations

import os
from pathlib import Path
import sys
from io import BytesIO
import zipfile

ROOT = Path(__file__).resolve().parents[1]
os.environ.setdefault("MPLCONFIGDIR", str(ROOT / "outputs" / ".matplotlib"))

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
from shiny import App, reactive, render, req, ui

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "scripts"))
from clean_productivity import (  # noqa: E402
    analyze_workbooks,
    parse_excel_date,
    read_sheet,
    workbook_sheet_names,
)
from data import AppData, from_analysis  # noqa: E402


MAX_UPLOAD_BYTES = 50 * 1024 * 1024
APP_CSS = Path(__file__).with_name("www").joinpath("styles.css").read_text()

PRODUCTIVITY_FIELDS = {
    "date": ("DATE", True),
    "species": ("Species", False),
    "plot": ("PLOT", True),
    "nest": ("Nest#", True),
    "slot": ("A or B chick", True),
    "eggs": ("Eggs", True),
    "chicks": ("Chicks", True),
    "status": ("Status", True),
    "pfr": ("PFR", False),
    "location": ("LOCATION", False),
    "observer": ("OBSERVER", False),
    "notes": ("Notes", False),
}
RESIGHT_FIELDS = {
    "date": ("Favorite Date", True),
    "species": ("Species", True),
    "combo": ("Combo", True),
    "fledged": ("Fledged?", True),
    "age": ("Age", False),
    "nest": ('NEST NUMBER (ROST NEST STARTS WITH "R")', False),
    "location": ("Location", False),
    "observer": ("Favorite Observer (who is not Joan)", False),
    "notes": ("Notes - please be brief", False),
}


def upload_paths(value) -> list[Path]:
    if not value:
        return []
    paths = []
    for item in value:
        if int(item.get("size", 0)) > MAX_UPLOAD_BYTES:
            raise ValueError("Each workbook must be 50 MB or smaller")
        name = str(item.get("name", ""))
        if not name.lower().endswith(".xlsx"):
            raise ValueError(f"{name or 'Upload'} must be an .xlsx workbook")
        paths.append(Path(item["datapath"]))
    return paths


def upload_path(value) -> Path | None:
    paths = upload_paths(value)
    return paths[0] if paths else None


def mapping_ui(prefix: str, fields, headers: list[str]):
    choices = {"": "— Not mapped —", **{header: header for header in headers}}
    controls = []
    for key, (canonical, required) in fields.items():
        controls.append(
            ui.input_select(
                f"{prefix}_map_{key}",
                f"{canonical}{' *' if required else ''}",
                choices,
                selected=canonical if canonical in headers else "",
            )
        )
    return ui.layout_columns(*controls, col_widths=(4, 4, 4))


def plot_choices(data: AppData) -> list[str]:
    return ["Overall", *sorted(data.chicks["plot"].dropna().unique())]


def safe_name(value: object) -> str:
    return "".join(c if c.isalnum() else "_" for c in str(value)).strip("_")


def empty_figure(message: str, figsize=(6, 4)):
    fig, ax = plt.subplots(figsize=figsize)
    ax.text(0.5, 0.5, message, ha="center", va="center", transform=ax.transAxes)
    ax.set_axis_off()
    fig.tight_layout()
    return fig


def outcome_figure(chicks: pd.DataFrame):
    order = ["verified_fledged", "known_dead", "unresolved", "not_hatched"]
    labels = ["Verified fledged", "Known dead", "Unresolved", "Not hatched"]
    counts = chicks["outcome"].value_counts().reindex(order, fill_value=0)
    fig, ax = plt.subplots(figsize=(7, 3.8))
    bars = ax.bar(labels, counts, color=["#3e8e41", "#b54137", "#da9e2c", "#78909c"])
    ax.bar_label(bars)
    ax.set_ylabel("Chick slots")
    ax.tick_params(axis="x", rotation=20)
    fig.tight_layout()
    return fig


def nest_box_figure(nests: pd.DataFrame):
    if nests.empty:
        return empty_figure("No nest data")
    fig, ax = plt.subplots(figsize=(7, 3.8))
    columns = ["clutch_size", "hatched_chicks", "verified_fledglings"]
    ax.boxplot(
        [nests[column].dropna() for column in columns],
        tick_labels=["Clutch size", "Hatched", "Fledged"],
    )
    ax.set_ylabel("Count per nest")
    fig.tight_layout()
    return fig


def clutch_figure(nests: pd.DataFrame, style: str):
    values = nests["clutch_size"].dropna()
    if values.empty:
        return empty_figure("No clutch-size data", (5.5, 4))
    fig, ax = plt.subplots(figsize=(5.5, 4))
    if style == "Box":
        ax.boxplot(values, tick_labels=["Clutch size"])
        ax.set_ylabel("Eggs per nest")
    else:
        counts = values.astype(int).value_counts().sort_index()
        bars = ax.bar(counts.index.astype(str), counts, color="#2670a0")
        ax.bar_label(bars)
        ax.set(xlabel="Clutch size", ylabel="Nests")
    fig.tight_layout()
    return fig


def nest_productivity_figure(nests: pd.DataFrame, style: str):
    columns = ["hatched_chicks", "verified_fledglings"]
    labels = ["Hatched chicks", "Verified fledglings"]
    if nests.empty:
        return empty_figure("No productivity data", (5.5, 4))
    fig, ax = plt.subplots(figsize=(5.5, 4))
    if style == "Box":
        ax.boxplot([nests[column].dropna() for column in columns], tick_labels=labels)
        ax.set_ylabel("Count per nest")
    else:
        means = [nests[column].mean() for column in columns]
        bars = ax.bar(labels, means, color=["#2670a0", "#3e8e41"])
        ax.bar_label(bars, fmt="%.2f")
        ax.set_ylabel("Mean per nest")
    fig.tight_layout()
    return fig


def chronology_figure(frame: pd.DataFrame, method: str, style: str):
    columns = {
        "Lay": f"lay_{method}",
        "Hatch": f"hatch_{method}",
        "Fledge": f"fledge_{method}",
    }
    series = {
        label: pd.to_datetime(frame[column], errors="coerce").dropna()
        for label, column in columns.items()
    }
    if not any(not values.empty for values in series.values()):
        return empty_figure("No chronology data", (9, 4.5))
    fig, ax = plt.subplots(figsize=(9, 4.5))
    if style == "Cumulative":
        for label, dates in series.items():
            dates = dates.sort_values()
            if not dates.empty:
                ax.step(dates, range(1, len(dates) + 1), where="post", label=label)
        ax.set(xlabel="Date", ylabel="Cumulative events")
        ax.legend()
    elif style == "Box":
        present = [(label, dates.map(mdates.date2num)) for label, dates in series.items() if not dates.empty]
        ax.boxplot([values for _, values in present], tick_labels=[label for label, _ in present])
        ax.yaxis_date()
        ax.set_ylabel("Event date")
    else:
        medians = [(label, dates.median()) for label, dates in series.items() if not dates.empty]
        bars = ax.bar(
            [label for label, _ in medians],
            [mdates.date2num(value) for _, value in medians],
            color=["#da9e2c", "#2670a0", "#3e8e41"][: len(medians)],
        )
        ax.bar_label(bars, labels=[value.strftime("%b %d") for _, value in medians])
        ax.yaxis_date()
        ax.set_ylabel("Median event date")
    fig.autofmt_xdate()
    fig.tight_layout()
    return fig


def comparison_figure(frame: pd.DataFrame, style: str, label: str, is_date: bool = False):
    if frame.empty or frame["group"].nunique() < 2:
        return empty_figure("Select at least two groups with data", (9, 4.8))
    groups = list(dict.fromkeys(frame["group"]))
    values = [frame.loc[frame["group"] == group, "value"].dropna() for group in groups]
    fig, ax = plt.subplots(figsize=(9, 4.8))
    if style == "Box":
        ax.boxplot(values, tick_labels=groups)
    else:
        means = [value.mean() for value in values]
        bars = ax.bar(groups, means, color="#2670a0")
        ax.bar_label(bars, fmt="%.2f")
    ax.set_ylabel(label)
    if is_date:
        ticks = ax.get_yticks()
        ax.set_yticks(ticks, [pd.Timestamp(2024, 1, 1).__add__(pd.Timedelta(days=max(0, tick - 1))).strftime("%b %d") for tick in ticks])
    fig.tight_layout()
    return fig


app_ui = ui.page_navbar(
    ui.nav_panel(
        "Upload data",
        ui.layout_columns(
            ui.card(
                ui.card_header("1. Select workbooks"),
                ui.input_radio_buttons(
                    "upload_comparison_mode",
                    "Analysis setup",
                    {
                        "single": "Analyze one species in one year",
                        "years": "Compare one species across years",
                        "species": "Compare species within one year",
                    },
                    selected="single",
                ),
                ui.output_ui("upload_requirements"),
                ui.input_file(
                    "productivity_file",
                    "Productivity workbook(s) *",
                    accept=[".xlsx"],
                    multiple=True,
                ),
                ui.output_ui("productivity_sheet_ui"),
                ui.input_file(
                    "resight_file",
                    "Resight workbook (optional)",
                    accept=[".xlsx"],
                    multiple=False,
                ),
                ui.output_ui("resight_sheet_ui"),
                ui.p("Uploads are processed only for this browser session.", class_="callout"),
            ),
            ui.card(
                ui.card_header("2. Map productivity columns"),
                ui.output_ui("productivity_mapping_ui"),
                ui.output_ui("species_fallback_ui"),
                ui.output_ui("detected_years_preview"),
            ),
            col_widths=(4, 8),
        ),
        ui.card(
            ui.card_header("3. Map optional resight columns"),
            ui.output_ui("resight_mapping_ui"),
        ),
        ui.div(
            ui.input_action_button(
                "generate_analysis", "Generate analysis", class_="btn-success btn-lg"
            ),
            ui.output_ui("upload_status"),
            class_="generate-row",
        ),
    ),
    ui.nav_panel(
        "Overview",
        ui.output_ui("overview_cards"),
        ui.layout_columns(
            ui.card(ui.card_header("Chick outcomes"), ui.output_plot("outcome_plot")),
            ui.card(ui.card_header("Nest distributions"), ui.output_plot("overview_box_plot")),
            ui.card(ui.card_header("Fledging rates"), ui.output_ui("rate_summary")),
            col_widths=(4, 4, 4),
        ),
    ),
    ui.nav_panel(
        "Productivity",
        ui.layout_sidebar(
            ui.sidebar(
                ui.output_ui("productivity_plot_control"),
                ui.input_select("productivity_chart_style", "Chart style", ["Bar", "Box"]),
            ),
            ui.layout_columns(
                ui.card(ui.card_header("Clutch size"), ui.output_plot("clutch_plot")),
                ui.card(
                    ui.card_header("Hatched chicks and fledglings per nest"),
                    ui.output_plot("nest_productivity_plot"),
                ),
                col_widths=(6, 6),
            ),
        ),
    ),
    ui.nav_panel(
        "Chronology",
        ui.layout_sidebar(
            ui.sidebar(
                ui.output_ui("chronology_plot_control"),
                ui.input_select(
                    "chronology_method",
                    "Date method",
                    {"midpoint": "Interval midpoint", "first_observed": "First observed"},
                ),
                ui.input_select(
                    "chronology_chart_style",
                    "Chart style",
                    ["Cumulative", "Bar", "Box"],
                ),
            ),
            ui.card(
                ui.card_header(ui.output_text("chronology_chart_title")),
                ui.output_plot("chronology_plot"),
            ),
            ui.output_ui("chronology_medians"),
        ),
    ),
    ui.nav_panel(
        "Compare",
        ui.layout_sidebar(
            ui.sidebar(
                ui.output_ui("comparison_mode_summary"),
                ui.output_ui("comparison_selectors"),
                ui.output_ui("comparison_plot_filter_ui"),
                ui.input_select(
                    "comparison_metric",
                    "Metric",
                    {
                        "clutch_size": "Clutch size",
                        "hatched_chicks": "Hatched chicks per nest",
                        "verified_fledglings": "Verified fledglings per nest",
                        "fledge_rate": "Verified fledging rate",
                        "lay_midpoint": "Lay date",
                        "hatch_midpoint": "Hatch date",
                        "fledge_midpoint": "Fledge date",
                    },
                ),
                ui.input_select("comparison_chart_style", "Chart style", ["Bar", "Box"]),
            ),
            ui.output_ui("comparison_status"),
            ui.card(ui.card_header("Comparison"), ui.output_plot("comparison_plot")),
        ),
    ),
    ui.nav_panel(
        "Downloads",
        ui.card(
            ui.card_header("Download current analysis"),
            ui.output_ui("download_selector_ui"),
            ui.download_button("download_selected", "Download selected"),
            ui.download_button("download_all", "Download all charts and CSVs"),
            ui.p("The ZIP includes all dashboard figures, analysis tables, and a comparison export when the current comparison is valid.", class_="callout"),
        ),
    ),
    ui.nav_panel(
        "Banded chicks",
        ui.layout_sidebar(
            ui.sidebar(
                ui.output_ui("band_plot_control"),
                ui.input_select("band_fledged", "Fledged", ["All", "Yes", "No"]),
                ui.input_select(
                    "band_outcome",
                    "Outcome",
                    ["All", "Verified Fledged", "Known Dead", "Unresolved"],
                ),
                ui.download_button("download_banded", "Download filtered CSV"),
            ),
            ui.p(
                "“No” means fledging was not verified. Outcome distinguishes "
                "known-dead from unresolved chicks.",
                class_="callout",
            ),
            ui.output_data_frame("banded_table"),
        ),
    ),
    ui.nav_panel(
        "Quality control",
        ui.layout_columns(
            ui.card(ui.card_header("Exceptions by issue"), ui.output_plot("qc_plot")),
            ui.card(
                ui.card_header("Exception records"),
                ui.download_button("download_qc", "Download QC CSV"),
                ui.output_data_frame("qc_table"),
            ),
            col_widths=(5, 7),
        ),
    ),
    ui.nav_panel(
        "Methods",
        ui.card(
            ui.h3("Data and classifications"),
            ui.tags.ul(
                ui.tags.li("The shared Python cleaning workflow is the source of truth."),
                ui.tags.li("Each year is analyzed independently."),
                ui.tags.li("Status = F or a same-year matched Yes! resight verifies fledging."),
                ui.tags.li("Apparent rate: verified fledglings / all hatched chicks."),
                ui.tags.li("Resolved rate: verified / (verified + known dead)."),
                ui.tags.li("Uploaded files are not written into project outputs."),
            ),
        ),
    ),
    ui.nav_spacer(),
    ui.nav_control(ui.output_ui("global_controls")),
    title="ROST Productivity",
    id="main_nav",
    header=ui.tags.style(APP_CSS),
    fillable=True,
)


def server(input, output, session):
    analysis_data = reactive.value(None)
    error_message = reactive.value("")
    success_message = reactive.value("")
    analysis_mode = reactive.value("single")
    productivity_headers = reactive.value([])
    resight_headers = reactive.value([])

    @render.ui
    def upload_requirements():
        if input.upload_comparison_mode() == "single":
            return ui.p(
                "Upload one workbook containing one year and one species.",
                class_="callout",
            )
        if input.upload_comparison_mode() == "species":
            return ui.p(
                "Upload one workbook containing one year and at least two species.",
                class_="callout",
            )
        return ui.p(
            "Upload at least two workbooks. Each workbook must contain one distinct year and the same species.",
            class_="callout",
        )

    @reactive.effect
    @reactive.event(input.productivity_file)
    def _load_productivity_sheets():
        try:
            path = upload_path(input.productivity_file())
            req(path)
            sheets = workbook_sheet_names(path)
            ui.update_select("productivity_sheet", choices=sheets, selected=sheets[0])
            _, rows = read_sheet(path, sheets[0])
            productivity_headers.set(
                [column for column in rows[0] if column != "_source_row"] if rows else []
            )
            error_message.set("")
        except Exception as exc:
            productivity_headers.set([])
            error_message.set(str(exc))

    @reactive.effect
    @reactive.event(input.productivity_sheet)
    def _load_productivity_headers():
        try:
            path = upload_path(input.productivity_file())
            req(path, input.productivity_sheet())
            _, rows = read_sheet(path, input.productivity_sheet())
            productivity_headers.set(
                [column for column in rows[0] if column != "_source_row"] if rows else []
            )
        except Exception as exc:
            error_message.set(str(exc))

    @reactive.effect
    @reactive.event(input.resight_file)
    def _load_resight_sheets():
        try:
            path = upload_path(input.resight_file())
            if path is None:
                resight_headers.set([])
                return
            sheets = workbook_sheet_names(path)
            ui.update_select("resight_sheet", choices=sheets, selected=sheets[0])
            _, rows = read_sheet(path, sheets[0])
            resight_headers.set(
                [column for column in rows[0] if column != "_source_row"] if rows else []
            )
        except Exception as exc:
            resight_headers.set([])
            error_message.set(str(exc))

    @reactive.effect
    @reactive.event(input.resight_sheet)
    def _load_resight_headers():
        try:
            path = upload_path(input.resight_file())
            if path is None:
                return
            _, rows = read_sheet(path, input.resight_sheet())
            resight_headers.set(
                [column for column in rows[0] if column != "_source_row"] if rows else []
            )
        except Exception as exc:
            error_message.set(str(exc))

    @render.ui
    def productivity_sheet_ui():
        return ui.input_select("productivity_sheet", "Productivity worksheet", [])

    @render.ui
    def resight_sheet_ui():
        if not input.resight_file():
            return ui.p("No resight workbook selected; Status = F will verify fledging.")
        return ui.input_select("resight_sheet", "Resight worksheet", [])

    @render.ui
    def productivity_mapping_ui():
        if not productivity_headers():
            return ui.p("Upload a productivity workbook to map its columns.")
        return mapping_ui("prod", PRODUCTIVITY_FIELDS, productivity_headers())

    @render.ui
    def species_fallback_ui():
        if not productivity_headers() or input.prod_map_species():
            return None
        return ui.input_text(
            "productivity_species_default",
            "Species for all productivity rows *",
            value="ROST",
            placeholder="For example: ROST",
        )

    @render.ui
    def detected_years_preview():
        if not productivity_headers():
            return None
        source_column = input.prod_map_date()
        if not source_column:
            return ui.p("Map the date column to preview analysis years.", class_="text-warning")
        try:
            paths = upload_paths(input.productivity_file())
            req(paths, input.productivity_sheet())
            details = []
            upload_items = input.productivity_file()
            for index, path in enumerate(paths):
                _, rows = read_sheet(path, input.productivity_sheet())
                parsed = [parse_excel_date(row.get(source_column, ""))[0] for row in rows]
                years = sorted({value.year for value in parsed if value})
                invalid = sum(value is None for value in parsed)
                filename = str(upload_items[index].get("name", path.name))
                label = filename + ": " + (", ".join(map(str, years)) if years else "no valid year")
                if invalid:
                    label += f" ({invalid:,} unreadable date row(s))"
                details.append(ui.tags.li(label))
            return ui.div(ui.strong("Detected years by workbook"), ui.tags.ul(*details), class_="callout")
        except Exception as exc:
            return ui.p(str(exc), class_="text-danger")

    @render.ui
    def resight_mapping_ui():
        if not input.resight_file():
            return ui.p("Optional. Upload a resight workbook to enable resight verification.")
        if not resight_headers():
            return ui.p("Select a readable resight worksheet.")
        return mapping_ui("res", RESIGHT_FIELDS, resight_headers())

    def collect_map(prefix: str, fields) -> dict[str, str]:
        result = {}
        selected = []
        for key, (canonical, required) in fields.items():
            accessor = getattr(input, f"{prefix}_map_{key}")
            source = accessor()
            if required and not source:
                raise ValueError(f"Map the required {canonical} column")
            if source:
                result[canonical] = source
                selected.append(source)
        duplicates = sorted({value for value in selected if selected.count(value) > 1})
        if duplicates:
            raise ValueError("Each source column may be mapped once: " + ", ".join(duplicates))
        return result

    @reactive.effect
    @reactive.event(input.generate_analysis)
    def _generate():
        error_message.set("")
        success_message.set("")
        try:
            productivity_paths = upload_paths(input.productivity_file())
            if not productivity_paths:
                raise ValueError("Upload at least one productivity workbook")
            mode = input.upload_comparison_mode()
            if mode == "years" and len(productivity_paths) < 2:
                raise ValueError("Cross-year comparison requires at least two productivity workbooks")
            if mode == "species" and len(productivity_paths) != 1:
                raise ValueError("Cross-species comparison requires exactly one productivity workbook")
            if mode == "single" and len(productivity_paths) != 1:
                raise ValueError("Single-year analysis requires exactly one productivity workbook")
            productivity_map = collect_map("prod", PRODUCTIVITY_FIELDS)
            species_default = ""
            if "Species" not in productivity_map:
                species_default = input.productivity_species_default().strip()
                if not species_default:
                    raise ValueError("Enter a species label or map a Species column")
            resight_path = upload_path(input.resight_file())
            resight_map = (
                collect_map("res", RESIGHT_FIELDS) if resight_path is not None else None
            )
            results = []
            workbook_years = []
            upload_items = input.productivity_file()
            for index, productivity_path in enumerate(productivity_paths):
                filename = str(upload_items[index].get("name", productivity_path.name))
                result = analyze_workbooks(
                    productivity_path=productivity_path,
                    productivity_sheet=input.productivity_sheet(),
                    productivity_map=productivity_map,
                    resight_path=resight_path,
                    resight_sheet=input.resight_sheet() if resight_path else None,
                    resight_map=resight_map,
                    species_default=species_default,
                )
                years = sorted({int(row["year"]) for row in result["chicks"]})
                if len(years) != 1:
                    raise ValueError(
                        f"{filename} must contain exactly one analysis year; found "
                        + (", ".join(map(str, years)) if years else "none")
                    )
                workbook_years.append(years[0])
                results.append(result)
            if mode == "years" and len(set(workbook_years)) != len(workbook_years):
                raise ValueError("Each cross-year workbook must represent a distinct year")
            combined = {
                key: [row for result in results for row in result[key]]
                for key in ("summary", "chronology_summary", "nests", "chicks", "chronology", "qc")
            }
            data = from_analysis(combined)
            if mode == "years" and len(data.species) != 1:
                raise ValueError("Cross-year workbooks must all contain the same single species")
            if mode == "species" and len(data.species) < 2:
                raise ValueError("Cross-species comparison requires at least two species in the workbook")
            if mode == "single" and (len(data.years) != 1 or len(data.species) != 1):
                raise ValueError("Single-year analysis requires exactly one year and one species")
            if not data.years:
                raise ValueError("No usable analysis years were detected")
            analysis_data.set(data)
            analysis_mode.set(mode)
            latest_year = data.years[-1]
            latest_species = sorted(
                data.chicks.loc[data.chicks["year"].eq(latest_year), "species"].unique()
            )
            ui.update_select(
                "selected_year",
                choices=[str(year) for year in data.years],
                selected=str(latest_year),
            )
            ui.update_select(
                "selected_species",
                choices=latest_species,
                selected=latest_species[0],
            )
            success_message.set(
                f"Analysis ready for {len(data.years)} year(s): "
                + ", ".join(map(str, data.years))
            )
            ui.update_navs("main_nav", selected="Overview")
        except Exception as exc:
            analysis_data.set(None)
            error_message.set(str(exc))

    @reactive.effect
    @reactive.event(input.reset_data)
    def _reset():
        analysis_data.set(None)
        error_message.set("")
        success_message.set("")
        ui.update_navs("main_nav", selected="Upload data")

    @render.ui
    def upload_status():
        if error_message():
            return ui.div(error_message(), class_="alert alert-danger")
        if success_message():
            return ui.div(success_message(), class_="alert alert-success")
        return ui.p("Required mappings are marked with *.")

    @render.ui
    def global_controls():
        if analysis_data() is None:
            return ui.span("Upload data to begin", class_="navbar-text")
        return ui.div(
            ui.input_select(
                "selected_year",
                "Year",
                [str(year) for year in analysis_data().years],
                selected=str(analysis_data().years[-1]),
            ),
            ui.input_select(
                "selected_species",
                "Species",
                sorted(
                    analysis_data().chicks.loc[
                        analysis_data().chicks["year"].eq(analysis_data().years[-1]),
                        "species",
                    ].unique()
                ),
            ),
            ui.input_action_button("reset_data", "Upload different data", class_="btn-outline-light"),
            class_="global-controls",
        )

    @reactive.effect
    @reactive.event(input.selected_year)
    def _update_species_for_year():
        data = analysis_data()
        req(data, input.selected_year())
        year = int(input.selected_year())
        choices = sorted(data.chicks.loc[data.chicks["year"].eq(year), "species"].unique())
        selected = input.selected_species() if input.selected_species() in choices else choices[0]
        ui.update_select("selected_species", choices=choices, selected=selected)

    @reactive.calc
    def current_data() -> AppData:
        data = analysis_data()
        req(data, input.selected_year(), input.selected_species())
        return data.for_year(int(input.selected_year()), input.selected_species())

    @render.ui
    def productivity_plot_control():
        req(analysis_data())
        return ui.input_select("productivity_plot", "Plot", plot_choices(current_data()))

    @render.ui
    def chronology_plot_control():
        req(analysis_data())
        return ui.input_select("chronology_plot_filter", "Plot", plot_choices(current_data()))

    @render.ui
    def band_plot_control():
        req(analysis_data())
        return ui.input_select(
            "band_plot", "Plot", ["All", *sorted(current_data().chicks["plot"].dropna().unique())]
        )

    @reactive.calc
    def overview_chicks():
        return current_data().chicks

    @render.ui
    def overview_cards():
        data = current_data()
        chicks = data.chicks
        counts = [
            ("Nests", len(data.nests), "primary"),
            ("Hatched", int(chicks["hatched"].sum()), "info"),
            ("Verified fledged", int(chicks["verified_fledged"].sum()), "success"),
            ("Known dead", int((chicks["outcome"] == "known_dead").sum()), "danger"),
            ("Unresolved", int((chicks["outcome"] == "unresolved").sum()), "warning"),
        ]
        return ui.layout_columns(
            *[ui.value_box(label, str(value), theme=theme) for label, value, theme in counts],
            col_widths=(2, 2, 3, 2, 3),
        )

    @render.plot
    def outcome_plot():
        return outcome_figure(overview_chicks())

    @render.plot
    def overview_box_plot():
        return nest_box_figure(current_data().nests)

    def metric_value(group: str, metric: str) -> pd.Series:
        rows = current_data().summary.loc[
            (current_data().summary["group"] == group)
            & (current_data().summary["metric"] == metric)
        ]
        req(not rows.empty)
        return rows.iloc[0]

    @render.ui
    def rate_summary():
        apparent = metric_value("Overall", "apparent_verified_fledge_rate")
        resolved = metric_value("Overall", "resolved_outcome_fledge_rate")
        return ui.div(
            ui.value_box(
                "Apparent verified rate",
                f"{float(apparent['mean']):.1%}",
                f"{int(apparent['numerator'])}/{int(apparent['denominator'])}",
                theme="primary",
            ),
            ui.value_box(
                "Resolved-outcome rate",
                f"{float(resolved['mean']):.1%}",
                f"{int(resolved['numerator'])}/{int(resolved['denominator'])}",
                theme="success",
            ),
        )

    @reactive.calc
    def productivity_nests():
        selected = input.productivity_plot()
        data = current_data().nests
        return data if selected == "Overall" else data.loc[data["plot"] == selected]

    @render.plot
    def clutch_plot():
        return clutch_figure(productivity_nests(), input.productivity_chart_style())

    @render.plot
    def nest_productivity_plot():
        return nest_productivity_figure(productivity_nests(), input.productivity_chart_style())

    @render.plot
    def chronology_plot():
        selected = input.chronology_plot_filter()
        method = input.chronology_method()
        frame = current_data().chronology
        if selected != "Overall":
            frame = frame.loc[frame["plot"] == selected]
        return chronology_figure(frame, method, input.chronology_chart_style())

    @render.text
    def chronology_chart_title():
        return f"{input.chronology_chart_style()} breeding chronology"

    @render.ui
    def chronology_medians():
        group = input.chronology_plot_filter()
        method = input.chronology_method()
        rows = current_data().chronology_summary.loc[
            (current_data().chronology_summary["group"] == group)
            & (current_data().chronology_summary["method"] == method)
        ]
        return ui.layout_columns(
            *[
                ui.value_box(
                    f"Median {row.event.lower()}",
                    str(row.median) if pd.notna(row.median) else "—",
                    f"n = {int(row.n)}",
                    theme={"lay": "warning", "hatch": "info", "fledge": "success"}[row.event],
                )
                for row in rows.itertuples()
            ],
            col_widths=(4, 4, 4),
        )

    def selected_comparison_mode() -> str:
        return analysis_mode()

    @render.ui
    def comparison_mode_summary():
        labels = {
            "single": "Single species, single year — comparison is not enabled",
            "years": "Same species across years",
            "species": "Same year across species",
        }
        label = labels[selected_comparison_mode()]
        return ui.div(ui.strong("Comparison type"), ui.p(label), class_="callout")

    @render.ui
    def comparison_selectors():
        data = analysis_data()
        req(data)
        if selected_comparison_mode() == "single":
            return ui.p(
                "Upload data using a comparison option to enable this tab.",
                class_="text-muted",
            )
        if selected_comparison_mode() == "species":
            return ui.TagList(
                ui.input_select(
                    "comparison_year",
                    "Year",
                    [str(year) for year in data.years],
                    selected=str(data.years[-1]),
                ),
                ui.input_select(
                    "comparison_species",
                    "Species (choose at least two)",
                    data.species,
                    selected=data.species,
                    multiple=True,
                ),
            )
        return ui.TagList(
            ui.input_select(
                "comparison_single_species",
                "Species",
                data.species,
                selected=data.species[0],
            ),
            ui.input_select(
                "comparison_years",
                "Years (choose at least two)",
                [str(year) for year in data.years],
                selected=[str(year) for year in data.years],
                multiple=True,
            ),
        )

    @render.ui
    def comparison_plot_filter_ui():
        data = analysis_data()
        req(data)
        plots = sorted(data.nests["plot"].dropna().astype(str).unique())
        locations = sorted(data.nests["location"].dropna().astype(str).unique())
        locations = [value for value in locations if value]
        choices = {"Overall": "Overall"}
        choices.update({f"plot:{value}": f"Plot: {value}" for value in plots})
        choices.update({f"location:{value}": f"Location: {value}" for value in locations})
        return ui.input_select("comparison_filter", "Plot/location", choices)

    def comparison_groups() -> tuple[str, list[str]]:
        if selected_comparison_mode() == "single":
            return "group", []
        if selected_comparison_mode() == "species":
            selected = list(input.comparison_species() or [])
            return "species", selected
        selected = [str(value) for value in (input.comparison_years() or [])]
        return "year", selected

    @reactive.calc
    def comparison_data() -> pd.DataFrame:
        data = analysis_data()
        req(data)
        if selected_comparison_mode() == "single":
            return pd.DataFrame(columns=["group", "value"])
        group_column, selected = comparison_groups()
        metric = input.comparison_metric()
        if group_column == "species":
            year = int(input.comparison_year())
            nests = data.nests.loc[pd.to_numeric(data.nests["year"], errors="coerce").eq(year) & data.nests["species"].isin(selected)].copy()
            chicks = data.chicks.loc[pd.to_numeric(data.chicks["year"], errors="coerce").eq(year) & data.chicks["species"].isin(selected)].copy()
            chronology = data.chronology.loc[pd.to_numeric(data.chronology["year"], errors="coerce").eq(year) & data.chronology["species"].isin(selected)].copy()
        else:
            years = [int(value) for value in selected]
            species = input.comparison_single_species()
            nests = data.nests.loc[pd.to_numeric(data.nests["year"], errors="coerce").isin(years) & data.nests["species"].eq(species)].copy()
            chicks = data.chicks.loc[pd.to_numeric(data.chicks["year"], errors="coerce").isin(years) & data.chicks["species"].eq(species)].copy()
            chronology = data.chronology.loc[pd.to_numeric(data.chronology["year"], errors="coerce").isin(years) & data.chronology["species"].eq(species)].copy()
        filter_value = input.comparison_filter()
        if filter_value and filter_value != "Overall":
            column, value = filter_value.split(":", 1)
            nests = nests.loc[nests[column].eq(value)]
            chicks = chicks.loc[chicks["plot"].eq(value)] if column == "plot" else chicks.loc[chicks["nest_key"].isin(nests["nest_key"])]
            chronology = chronology.loc[chronology["plot"].eq(value)] if column == "plot" else chronology.loc[chronology["nest_key"].isin(nests["nest_key"])]
        source = nests
        if metric == "fledge_rate":
            source = chicks.loc[chicks["hatched"]].copy()
            source["value"] = source["verified_fledged"].astype(float)
        elif metric in {"lay_midpoint", "hatch_midpoint", "fledge_midpoint"}:
            source = chronology.copy()
            source["value"] = pd.to_datetime(source[metric], errors="coerce").dt.dayofyear
        else:
            source = nests.copy()
            source["value"] = pd.to_numeric(source[metric], errors="coerce")
        source["group"] = source[group_column].astype(str)
        return source[["group", "value"]].dropna()

    @render.ui
    def comparison_status():
        if selected_comparison_mode() == "single":
            return ui.div(
                "This upload is configured for single-year analysis; no comparison is generated.",
                class_="alert alert-info",
            )
        group_column, selected = comparison_groups()
        available = comparison_data()["group"].nunique()
        if len(selected) < 2:
            return ui.div("Choose at least two years or species.", class_="alert alert-warning")
        if available < 2:
            return ui.div("Fewer than two selected groups contain data for this metric and filter.", class_="alert alert-warning")
        return ui.div(f"Comparing {available} {group_column} groups.", class_="alert alert-success")

    @render.plot
    def comparison_plot():
        metric = input.comparison_metric()
        labels = {
            "clutch_size": "Clutch size",
            "hatched_chicks": "Hatched chicks per nest",
            "verified_fledglings": "Verified fledglings per nest",
            "fledge_rate": "Verified fledging rate",
            "lay_midpoint": "Lay date",
            "hatch_midpoint": "Hatch date",
            "fledge_midpoint": "Fledge date",
        }
        return comparison_figure(
            comparison_data(),
            input.comparison_chart_style(),
            labels[metric],
            metric.endswith("_midpoint"),
        )

    @reactive.calc
    def filtered_banded():
        frame = current_data().banded_chicks
        if input.band_plot() != "All":
            frame = frame.loc[frame["plot"] == input.band_plot()]
        if input.band_fledged() != "All":
            frame = frame.loc[frame["fledged"] == input.band_fledged()]
        if input.band_outcome() != "All":
            frame = frame.loc[frame["outcome"] == input.band_outcome()]
        return frame

    @render.data_frame
    def banded_table():
        return render.DataGrid(filtered_banded(), filters=True, width="100%", height="620px")

    @render.download(
        filename=lambda: f"banded_ROST_chicks_{input.selected_year()}_filtered.csv"
    )
    def download_banded():
        yield filtered_banded().to_csv(index=False)

    @render.plot
    def qc_plot():
        counts = current_data().qc["issue"].value_counts().head(12).sort_values()
        fig, ax = plt.subplots(figsize=(6, 5))
        bars = ax.barh(counts.index.str.replace("_", " "), counts, color="#2670a0")
        ax.bar_label(bars)
        ax.set_xlabel("Flagged records")
        fig.tight_layout()
        return fig

    @render.data_frame
    def qc_table():
        return render.DataGrid(current_data().qc, filters=True, width="100%", height="600px")

    def filtered_chronology() -> pd.DataFrame:
        frame = current_data().chronology
        selected = input.chronology_plot_filter()
        return frame if selected == "Overall" else frame.loc[frame["plot"] == selected]

    def figure_bytes(fig) -> bytes:
        buffer = BytesIO()
        fig.savefig(buffer, format="png", dpi=180, bbox_inches="tight")
        plt.close(fig)
        return buffer.getvalue()

    def artifact_context() -> str:
        return "_".join(
            [
                safe_name(input.selected_species()),
                safe_name(input.selected_year()),
            ]
        )

    def artifact_payload(key: str) -> tuple[str, bytes]:
        context = artifact_context()
        style = input.productivity_chart_style()
        if key == "outcome_png":
            return f"chick_outcomes_{context}.png", figure_bytes(outcome_figure(current_data().chicks))
        if key == "overview_box_png":
            return f"nest_distributions_{context}.png", figure_bytes(nest_box_figure(current_data().nests))
        if key == "clutch_png":
            name = f"clutch_size_{safe_name(style)}_{context}_{safe_name(input.productivity_plot())}.png"
            return name, figure_bytes(clutch_figure(productivity_nests(), style))
        if key == "productivity_png":
            name = f"nest_productivity_{safe_name(style)}_{context}_{safe_name(input.productivity_plot())}.png"
            return name, figure_bytes(nest_productivity_figure(productivity_nests(), style))
        if key == "chronology_png":
            style = input.chronology_chart_style()
            name = f"chronology_{safe_name(style)}_{context}_{safe_name(input.chronology_plot_filter())}_{safe_name(input.chronology_method())}.png"
            return name, figure_bytes(chronology_figure(filtered_chronology(), input.chronology_method(), style))
        if key == "comparison_png":
            metric = input.comparison_metric()
            labels = {
                "clutch_size": "Clutch size",
                "hatched_chicks": "Hatched chicks per nest",
                "verified_fledglings": "Verified fledglings per nest",
                "fledge_rate": "Verified fledging rate",
                "lay_midpoint": "Lay date",
                "hatch_midpoint": "Hatch date",
                "fledge_midpoint": "Fledge date",
            }
            name = f"comparison_{safe_name(selected_comparison_mode())}_{safe_name(metric)}_{safe_name(input.comparison_chart_style())}_{safe_name(input.comparison_filter())}.png"
            fig = comparison_figure(comparison_data(), input.comparison_chart_style(), labels[metric], metric.endswith("_midpoint"))
            return name, figure_bytes(fig)
        csv_artifacts = {
            "nests_csv": ("nests", current_data().nests),
            "chicks_csv": ("chicks", current_data().chicks),
            "summary_csv": ("summary", current_data().summary),
            "chronology_csv": ("chronology", current_data().chronology),
            "banded_csv": ("banded_filtered", filtered_banded()),
            "qc_csv": ("quality_control", current_data().qc),
            "comparison_csv": (
                f"comparison_{safe_name(selected_comparison_mode())}_{safe_name(input.comparison_metric())}_{safe_name(input.comparison_filter())}",
                comparison_data(),
            ),
        }
        label, frame = csv_artifacts[key]
        return f"{label}_{context}.csv", frame.to_csv(index=False).encode("utf-8")

    @render.ui
    def download_selector_ui():
        req(analysis_data())
        return ui.input_select(
            "download_artifact",
            "Figure or table",
            {
                "outcome_png": "Chick outcomes (PNG)",
                "overview_box_png": "Nest distributions (PNG)",
                "clutch_png": "Clutch size, current view (PNG)",
                "productivity_png": "Nest productivity, current view (PNG)",
                "chronology_png": "Chronology, current view (PNG)",
                "comparison_png": "Comparison, current view (PNG)",
                "nests_csv": "Nests (CSV)",
                "chicks_csv": "Chicks (CSV)",
                "summary_csv": "Summary (CSV)",
                "chronology_csv": "Chronology (CSV)",
                "banded_csv": "Filtered banded chicks (CSV)",
                "qc_csv": "Quality control (CSV)",
                "comparison_csv": "Comparison values (CSV)",
            },
        )

    @render.download(filename=lambda: artifact_payload(input.download_artifact())[0])
    def download_selected():
        yield artifact_payload(input.download_artifact())[1]

    @render.download(filename=lambda: f"ROST_dashboard_exports_{artifact_context()}.zip")
    def download_all():
        buffer = BytesIO()
        keys = [
            "outcome_png", "overview_box_png", "nests_csv", "chicks_csv", "summary_csv",
            "chronology_csv", "banded_csv", "qc_csv",
        ]
        if comparison_data()["group"].nunique() >= 2:
            keys.append("comparison_csv")
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for key in keys:
                name, payload = artifact_payload(key)
                archive.writestr(name, payload)
            context = artifact_context()
            productivity_context = f"{context}_{safe_name(input.productivity_plot())}"
            chronology_context = f"{context}_{safe_name(input.chronology_plot_filter())}_{safe_name(input.chronology_method())}"
            for style in ("Bar", "Box"):
                archive.writestr(
                    f"clutch_size_{style}_{productivity_context}.png",
                    figure_bytes(clutch_figure(productivity_nests(), style)),
                )
                archive.writestr(
                    f"nest_productivity_{style}_{productivity_context}.png",
                    figure_bytes(nest_productivity_figure(productivity_nests(), style)),
                )
            for style in ("Cumulative", "Bar", "Box"):
                archive.writestr(
                    f"chronology_{style}_{chronology_context}.png",
                    figure_bytes(
                        chronology_figure(
                            filtered_chronology(), input.chronology_method(), style
                        )
                    ),
                )
            if comparison_data()["group"].nunique() >= 2:
                metric = input.comparison_metric()
                label = {
                    "clutch_size": "Clutch size",
                    "hatched_chicks": "Hatched chicks per nest",
                    "verified_fledglings": "Verified fledglings per nest",
                    "fledge_rate": "Verified fledging rate",
                    "lay_midpoint": "Lay date",
                    "hatch_midpoint": "Hatch date",
                    "fledge_midpoint": "Fledge date",
                }[metric]
                for style in ("Bar", "Box"):
                    archive.writestr(
                        f"comparison_{safe_name(selected_comparison_mode())}_{safe_name(metric)}_{style}_{safe_name(input.comparison_filter())}.png",
                        figure_bytes(
                            comparison_figure(
                                comparison_data(),
                                style,
                                label,
                                metric.endswith("_midpoint"),
                            )
                        ),
                    )
        yield buffer.getvalue()

    @render.download(
        filename=lambda: f"quality_control_exceptions_{input.selected_year()}.csv"
    )
    def download_qc():
        yield current_data().qc.to_csv(index=False)


app = App(app_ui, server)
