"""Session-based, multi-year dashboard for ROST productivity workbooks."""

from __future__ import annotations

import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
os.environ.setdefault("MPLCONFIGDIR", str(ROOT / "outputs" / ".matplotlib"))

import matplotlib.pyplot as plt
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


def upload_path(value) -> Path | None:
    if not value:
        return None
    item = value[0]
    if int(item.get("size", 0)) > MAX_UPLOAD_BYTES:
        raise ValueError("Each workbook must be 50 MB or smaller")
    name = str(item.get("name", ""))
    if not name.lower().endswith(".xlsx"):
        raise ValueError(f"{name or 'Upload'} must be an .xlsx workbook")
    return Path(item["datapath"])


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


app_ui = ui.page_navbar(
    ui.nav_panel(
        "Upload data",
        ui.layout_columns(
            ui.card(
                ui.card_header("1. Select workbooks"),
                ui.input_file(
                    "productivity_file",
                    "Productivity workbook *",
                    accept=[".xlsx"],
                    multiple=False,
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
            ui.card(ui.card_header("Fledging rates"), ui.output_ui("rate_summary")),
            col_widths=(7, 5),
        ),
    ),
    ui.nav_panel(
        "Productivity",
        ui.layout_sidebar(
            ui.sidebar(ui.output_ui("productivity_plot_control")),
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
            ),
            ui.card(
                ui.card_header("Cumulative breeding chronology"),
                ui.output_plot("chronology_plot"),
            ),
            ui.output_ui("chronology_medians"),
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
    productivity_headers = reactive.value([])
    resight_headers = reactive.value([])

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
    def detected_years_preview():
        if not productivity_headers():
            return None
        source_column = input.prod_map_date()
        if not source_column:
            return ui.p("Map the date column to preview analysis years.", class_="text-warning")
        try:
            path = upload_path(input.productivity_file())
            req(path, input.productivity_sheet())
            _, rows = read_sheet(path, input.productivity_sheet())
            parsed = [parse_excel_date(row.get(source_column, ""))[0] for row in rows]
            years = sorted({value.year for value in parsed if value})
            invalid = sum(value is None for value in parsed)
            if not years:
                return ui.p("No valid years detected in the mapped date column.", class_="text-danger")
            suffix = f"; {invalid:,} row(s) have unreadable dates" if invalid else ""
            return ui.p(
                "Detected analysis years: " + ", ".join(map(str, years)) + suffix,
                class_="callout",
            )
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
            productivity_path = upload_path(input.productivity_file())
            if productivity_path is None:
                raise ValueError("Upload a productivity workbook")
            productivity_map = collect_map("prod", PRODUCTIVITY_FIELDS)
            resight_path = upload_path(input.resight_file())
            resight_map = (
                collect_map("res", RESIGHT_FIELDS) if resight_path is not None else None
            )
            result = analyze_workbooks(
                productivity_path=productivity_path,
                productivity_sheet=input.productivity_sheet(),
                productivity_map=productivity_map,
                resight_path=resight_path,
                resight_sheet=input.resight_sheet() if resight_path else None,
                resight_map=resight_map,
            )
            data = from_analysis(result)
            if not data.years:
                raise ValueError("No usable analysis years were detected")
            analysis_data.set(data)
            ui.update_select(
                "selected_year",
                choices=[str(year) for year in data.years],
                selected=str(data.years[-1]),
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
            ui.input_action_button("reset_data", "Upload different data", class_="btn-outline-light"),
            class_="global-controls",
        )

    @reactive.calc
    def current_data() -> AppData:
        data = analysis_data()
        req(data, input.selected_year())
        return data.for_year(int(input.selected_year()))

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
        order = ["verified_fledged", "known_dead", "unresolved", "not_hatched"]
        labels = ["Verified fledged", "Known dead", "Unresolved", "Not hatched"]
        counts = overview_chicks()["outcome"].value_counts().reindex(order, fill_value=0)
        fig, ax = plt.subplots(figsize=(7, 3.8))
        bars = ax.bar(labels, counts, color=["#3e8e41", "#b54137", "#da9e2c", "#78909c"])
        ax.bar_label(bars)
        ax.set_ylabel("Chick slots")
        ax.tick_params(axis="x", rotation=20)
        fig.tight_layout()
        return fig

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
        counts = productivity_nests()["clutch_size"].dropna().astype(int).value_counts().sort_index()
        fig, ax = plt.subplots(figsize=(5.5, 4))
        bars = ax.bar(counts.index.astype(str), counts, color="#2670a0")
        ax.bar_label(bars)
        ax.set(xlabel="Egg-bearing slots", ylabel="Nests")
        fig.tight_layout()
        return fig

    @render.plot
    def nest_productivity_plot():
        nests = productivity_nests()
        means = [nests["hatched_chicks"].mean(), nests["verified_fledglings"].mean()]
        fig, ax = plt.subplots(figsize=(5.5, 4))
        bars = ax.bar(["Hatched chicks", "Verified fledglings"], means, color=["#2670a0", "#3e8e41"])
        ax.bar_label(bars, fmt="%.2f")
        ax.set_ylabel("Mean per nest")
        fig.tight_layout()
        return fig

    @render.plot
    def chronology_plot():
        selected = input.chronology_plot_filter()
        method = input.chronology_method()
        frame = current_data().chronology
        if selected != "Overall":
            frame = frame.loc[frame["plot"] == selected]
        columns = {"Lay": f"lay_{method}", "Hatch": f"hatch_{method}", "Fledge": f"fledge_{method}"}
        fig, ax = plt.subplots(figsize=(9, 4.5))
        for label, column in columns.items():
            dates = pd.to_datetime(frame[column], errors="coerce").dropna().sort_values()
            if not dates.empty:
                ax.step(dates, range(1, len(dates) + 1), where="post", label=label)
        ax.set(xlabel="Date", ylabel="Cumulative events")
        ax.legend()
        fig.autofmt_xdate()
        fig.tight_layout()
        return fig

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

    @render.download(
        filename=lambda: f"quality_control_exceptions_{input.selected_year()}.csv"
    )
    def download_qc():
        yield current_data().qc.to_csv(index=False)


app = App(app_ui, server)
