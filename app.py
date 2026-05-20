from __future__ import annotations

import re
import math
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import folium
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
import streamlit.components.v1 as components
from streamlit_folium import st_folium


st.set_page_config(
    page_title="Apple Cup Dashboard",
    page_icon="🏈",
    layout="wide",
)


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = Path(r"C:\Users\jesse\6361 Streamlit\Apple_Cup_History_AP_Rankings.xlsx")
TWBX_FILE = BASE_DIR / "Apple Cup Games History Stats.twbx"
LOGO_FILE = BASE_DIR / "apple-cup.jpg"

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

RESULT_COLORS = {
    "UW": "#B07AA1",
    "WSU": "#E15759",
    "Tie": "#4E79A7",
}
RANKED_COLORS = {
    "Yes": "#59A14F",
    "No": "#BAB0AC",
    "Pre-AP Poll": "#F28E2B",
}
HOME_FIELD_POSITIONS = {
    "Seattle": {"x": [0.11, 0.31], "y": [0.54, 0.76]},
    "Spokane": {"x": [0.60, 0.80], "y": [0.57, 0.79]},
    "Pullman": {"x": [0.67, 0.87], "y": [0.22, 0.44]},
}
CITY_COORDINATES = {
    "Seattle": {"lat": 47.6062, "lon": -122.3321},
    "Spokane": {"lat": 47.6588, "lon": -117.4260},
    "Pullman": {"lat": 46.7298, "lon": -117.1817},
}


def empty_figure(title: str, height: int) -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(
        text="No data for the current filters",
        x=0.5,
        y=0.5,
        showarrow=False,
        font=dict(size=16, color="#8A8A8A"),
    )
    fig.update_layout(title=dict(text=title, x=0.5, xanchor="center"))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return configure_chart(fig, height)


def rgba(hex_color: str, alpha: float) -> str:
    hex_color = hex_color.lstrip("#")
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    return f"rgba({red}, {green}, {blue}, {alpha})"


def polar_to_cartesian(cx: float, cy: float, radius: float, angle_deg: float) -> tuple[float, float]:
    angle_rad = math.radians(angle_deg - 90)
    return cx + radius * math.cos(angle_rad), cy + radius * math.sin(angle_rad)


def svg_arc_path(cx: float, cy: float, radius: float, start_angle: float, end_angle: float) -> str:
    start_x, start_y = polar_to_cartesian(cx, cy, radius, start_angle)
    end_x, end_y = polar_to_cartesian(cx, cy, radius, end_angle)
    large_arc = 1 if end_angle - start_angle > 180 else 0
    return (
        f"M {cx} {cy} "
        f"L {start_x:.2f} {start_y:.2f} "
        f"A {radius} {radius} 0 {large_arc} 1 {end_x:.2f} {end_y:.2f} Z"
    )


def build_pie_marker_svg(city: str, uw: int, wsu: int, tie: int) -> str:
    total = max(uw + wsu + tie, 1)
    radius = 30 + min(total, 20) * 0.7
    center = radius + 8
    size = int(center * 2 + 52)
    pie_bottom = center * 2 + 10
    current_angle = 0.0
    slices = [("UW", uw), ("WSU", wsu), ("Tie", tie)]
    paths: list[str] = []
    for label, value in slices:
        if value <= 0:
            continue
        sweep = 360 * (value / total)
        path = svg_arc_path(center, center, radius, current_angle, current_angle + sweep)
        paths.append(
            f"<path d='{path}' fill='{RESULT_COLORS[label]}' stroke='white' stroke-width='2' />"
        )
        current_angle += sweep

    legend_text = f"""
    <text x="{center}" y="{center + 7}" text-anchor="middle"
          font-family="Arial, sans-serif" font-size="22" font-weight="700" fill="#404040">{total}</text>
    <text x="{center}" y="{pie_bottom + 18}" text-anchor="middle"
          font-family="Arial, sans-serif" font-size="14" font-weight="700" fill="#404040">{city}</text>
    """
    return f"""
    <svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">
      <defs>
        <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="2" stdDeviation="2" flood-color="#000000" flood-opacity="0.16"/>
        </filter>
      </defs>
      <g filter="url(#shadow)">
        <circle cx="{center}" cy="{center}" r="{radius + 3}" fill="white" opacity="0.95"/>
        {''.join(paths)}
        <circle cx="{center}" cy="{center}" r="{radius * 0.44:.1f}" fill="white" />
      </g>
      {legend_text}
    </svg>
    """


def excel_column_index(cell_reference: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_reference).group(1)
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def workbook_rows(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("a:si", NS):
                shared_strings.append("".join(node.text or "" for node in item.iterfind(".//a:t", NS)))

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in rel_root
        }

        first_sheet = workbook_root.find("a:sheets", NS)[0]
        relationship_id = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        worksheet_path = "xl/" + rel_map[relationship_id]
        worksheet_root = ET.fromstring(archive.read(worksheet_path))

        rows: list[list[str]] = []
        for row in worksheet_root.findall(".//a:sheetData/a:row", NS):
            current: list[str] = []
            current_index = 0
            for cell in row.findall("a:c", NS):
                cell_index = excel_column_index(cell.attrib["r"])
                while current_index < cell_index:
                    current.append("")
                    current_index += 1

                value_node = cell.find("a:v", NS)
                if value_node is None:
                    current.append("")
                elif cell.attrib.get("t") == "s":
                    current.append(shared_strings[int(value_node.text)])
                else:
                    current.append(value_node.text)
                current_index += 1
            rows.append(current)
        return rows


@st.cache_data(show_spinner=False)
def load_data() -> pd.DataFrame:
    rows = workbook_rows(DATA_FILE)
    header = rows[0]
    normalized_rows = [row + [""] * (len(header) - len(row)) for row in rows[1:]]
    df = pd.DataFrame(normalized_rows, columns=header)

    numeric_columns = [
        "Deficit- UW",
        "No.",
        "Score -L",
        "Score-W",
        "UW SOV",
        "WIn Margin",
        "WSU SOV",
    ]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df["Game Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Year"] = df["Game Date"].dt.year
    df["Abs Win Margin"] = df["WIn Margin"].abs()
    df["SOV Calc"] = df.apply(
        lambda row: row["UW SOV"] if row["Result"] == "UW" else row["WSU SOV"] if row["Result"] == "WSU" else pd.NA,
        axis=1,
    )
    return df.sort_values("Game Date").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_logo() -> bytes | None:
    if LOGO_FILE.exists():
        return LOGO_FILE.read_bytes()
    if TWBX_FILE.exists():
        with zipfile.ZipFile(TWBX_FILE) as archive:
            if "Image/apple-cup.jpg" in archive.namelist():
                return archive.read("Image/apple-cup.jpg")
    return None


def configure_chart(fig: go.Figure, height: int) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=56, b=36),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color="#303030"),
        showlegend=False,
    )
    return fig


def apply_filters(
    df: pd.DataFrame,
    year_range: tuple[int, int],
    results: list[str],
    home_fields: list[str],
    ranked_statuses: list[str],
    ot_values: list[str],
) -> pd.DataFrame:
    mask = (
        df["Year"].between(year_range[0], year_range[1])
        & df["Result"].isin(results)
        & df["Home Field"].isin(home_fields)
        & df["Both Teams Ranked"].isin(ranked_statuses)
        & df["OT"].isin(ot_values)
    )
    return df.loc[mask].copy()


def build_total_wins_donut(df: pd.DataFrame, title: str, center_label: str, height: int = 320) -> go.Figure:
    if df.empty:
        return empty_figure(title, height)

    counts = df["Result"].value_counts().reindex(["UW", "WSU", "Tie"], fill_value=0)
    text_labels = [f"{label}<br>{value}" if value > 0 else label for label, value in counts.items()]

    fig = go.Figure(
        go.Pie(
            labels=counts.index,
            values=counts.values,
            hole=0.68,
            sort=False,
            direction="clockwise",
            marker=dict(colors=[RESULT_COLORS[key] for key in counts.index], line=dict(color="white", width=1)),
            text=text_labels,
            textinfo="text",
            textposition="outside",
            hovertemplate="%{label}: %{value} games<extra></extra>",
        )
    )
    fig.add_annotation(
        text=center_label,
        x=0.5,
        y=0.5,
        showarrow=False,
        font=dict(size=18, color="#666666"),
        align="center",
    )
    fig.update_layout(title=dict(text=title, x=0.5, xanchor="center"))
    return configure_chart(fig, height)


def build_probability_chart(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return empty_figure("Winning Probability", 320)

    counts = df["Result"].value_counts().reindex(["UW", "WSU", "Tie"], fill_value=0)
    probability = (counts / counts.sum()).fillna(0).sort_values(ascending=True)
    text_positions = ["outside" if value < 0.12 else "inside" for value in probability.values]

    fig = go.Figure(
        go.Bar(
            x=probability.values,
            y=probability.index,
            orientation="h",
            marker=dict(color=[RESULT_COLORS[key] for key in probability.index]),
            text=[f"{value:.1%}" for value in probability.values],
            textposition=text_positions,
            cliponaxis=False,
            hovertemplate="%{y}: %{x:.1%}<extra></extra>",
        )
    )
    fig.update_layout(title=dict(text="Winning Probability", x=0.5, xanchor="center"))
    fig.update_xaxes(range=[0, 1], tickformat=".0%", title_text="Probability", gridcolor="#EFEFEF", zeroline=False)
    fig.update_yaxes(title_text="")
    return configure_chart(fig, 320)


def build_largest_win_margin_chart(df: pd.DataFrame, reference_value: int) -> go.Figure:
    chart_df = (
        df[df["Result"].isin(["UW", "WSU"])]
        .sort_values(["Abs Win Margin", "Game Date"], ascending=[False, True])
        .drop_duplicates("Result")
        .sort_values("Abs Win Margin")
    )
    if chart_df.empty:
        return empty_figure("Largest Win Margin", 320)

    fig = go.Figure(
        go.Bar(
            x=chart_df["Abs Win Margin"],
            y=chart_df["Result"],
            orientation="h",
            marker=dict(color=[RESULT_COLORS[key] for key in chart_df["Result"]]),
            text=chart_df["Abs Win Margin"].astype(int).astype(str),
            textposition="outside",
            customdata=chart_df[["Year", "Score"]],
            hovertemplate="%{y}<br>Margin: %{x}<br>Year: %{customdata[0]}<br>Score: %{customdata[1]}<extra></extra>",
        )
    )
    fig.add_vline(
        x=reference_value,
        line_dash="dot",
        line_color="#9E9E9E",
        opacity=0.7,
    )
    fig.update_layout(title=dict(text="Largest Win Margin", x=0.5, xanchor="center"))
    fig.update_xaxes(title_text="Point Margin", gridcolor="#EFEFEF", zeroline=False)
    fig.update_yaxes(title_text="")
    return configure_chart(fig, 320)


def build_sov_chart(df: pd.DataFrame) -> go.Figure:
    chart_df = (
        df[df["Result"].isin(["UW", "WSU"])]
        .groupby("Result", as_index=False)["SOV Calc"]
        .mean()
        .sort_values("SOV Calc")
    )
    if chart_df.empty:
        return empty_figure("SOV Calc", 320)

    fig = go.Figure(
        go.Bar(
            x=chart_df["SOV Calc"],
            y=chart_df["Result"],
            orientation="h",
            marker=dict(color=[RESULT_COLORS[key] for key in chart_df["Result"]]),
            text=[f"{value:.1f}" for value in chart_df["SOV Calc"]],
            textposition="outside",
            hovertemplate="%{y}<br>Average SOV: %{x:.2f}<extra></extra>",
        )
    )
    fig.update_layout(title=dict(text="SOV Calc", x=0.5, xanchor="center"))
    fig.update_xaxes(title_text="Average Strength of Victory", gridcolor="#EFEFEF", zeroline=False)
    fig.update_yaxes(title_text="")
    return configure_chart(fig, 320)


def build_scores_timeline(
    df: pd.DataFrame,
    year_range: tuple[int, int],
    reference_value: int,
    reveal_count: int | None = None,
) -> go.Figure:
    chart_df = df[
        df["Result"].isin(["UW", "WSU"])
        & df["Year"].between(year_range[0], year_range[1])
    ].copy()
    if chart_df.empty:
        return empty_figure("Scores Timeline", 480)

    chart_df = chart_df.sort_values("Game Date").reset_index(drop=True)
    if reveal_count is not None:
        chart_df = chart_df.iloc[:reveal_count].copy()

    chart_df["Label"] = chart_df["Abs Win Margin"].where(chart_df["Abs Win Margin"] >= 24, "").astype(str)
    chart_df.loc[chart_df["Label"] == "nan", "Label"] = ""
    tick_values = sorted(df[df["Year"].between(year_range[0], year_range[1])]["Year"].dropna().astype(int).unique().tolist())

    fig = go.Figure(
        go.Bar(
            x=chart_df["Year"],
            y=chart_df["WIn Margin"],
            marker=dict(color=[RESULT_COLORS[key] for key in chart_df["Result"]]),
            text=chart_df["Label"],
            textposition="outside",
            customdata=chart_df[["Result", "Score", "Winner", "Home Field"]].to_numpy(),
            hovertemplate=(
                "Year: %{x}<br>"
                "Margin: %{y}<br>"
                "Result: %{customdata[0]}<br>"
                "Score: %{customdata[1]}<br>"
                "Winner: %{customdata[2]}<br>"
                "Home Field: %{customdata[3]}<extra></extra>"
            ),
        )
    )
    fig.add_hline(y=0, line_color="#BEBEBE", line_width=1)
    fig.add_hline(y=reference_value, line_dash="dot", line_color="#A4A4A4", opacity=0.6)
    fig.add_hline(y=-reference_value, line_dash="dot", line_color="#A4A4A4", opacity=0.6)
    fig.update_layout(title=dict(text="Scores Timeline", x=0.5, xanchor="center"))
    fig.update_xaxes(title_text="Year of Game Date", tickmode="array", tickvals=tick_values, gridcolor="#F4F4F4")
    fig.update_yaxes(title_text="Win Margin", gridcolor="#EFEFEF", zeroline=False)
    return configure_chart(fig, 480)


def build_scores_timeline_animated(df: pd.DataFrame, year_range: tuple[int, int], reference_value: int) -> go.Figure:
    chart_df = df[
        df["Result"].isin(["UW", "WSU"])
        & df["Year"].between(year_range[0], year_range[1])
    ].copy()
    if chart_df.empty:
        return empty_figure("Scores Timeline", 560)

    chart_df = chart_df.sort_values("Game Date").reset_index(drop=True)
    chart_df["Label"] = chart_df["Abs Win Margin"].where(chart_df["Abs Win Margin"] >= 24, "").astype(str)
    chart_df.loc[chart_df["Label"] == "nan", "Label"] = ""
    tick_values = sorted(chart_df["Year"].dropna().astype(int).unique().tolist())

    def timeline_trace(frame_df: pd.DataFrame) -> go.Bar:
        return go.Bar(
            x=frame_df["Year"],
            y=frame_df["WIn Margin"],
            marker=dict(color=[RESULT_COLORS[key] for key in frame_df["Result"]]),
            text=frame_df["Label"],
            textposition="outside",
            customdata=frame_df[["Result", "Score", "Winner", "Home Field"]].to_numpy(),
            hovertemplate=(
                "Year: %{x}<br>"
                "Margin: %{y}<br>"
                "Result: %{customdata[0]}<br>"
                "Score: %{customdata[1]}<br>"
                "Winner: %{customdata[2]}<br>"
                "Home Field: %{customdata[3]}<extra></extra>"
            ),
        )

    frames = []
    for idx in range(len(chart_df)):
        frame_df = chart_df.iloc[: idx + 1]
        frame_label = str(frame_df.iloc[-1]["Date"])
        frames.append(go.Frame(name=frame_label, data=[timeline_trace(frame_df)]))

    fig = go.Figure(data=[timeline_trace(chart_df.iloc[:1])], frames=frames)
    fig.add_hline(y=0, line_color="#BEBEBE", line_width=1)
    fig.add_hline(y=reference_value, line_dash="dot", line_color="#A4A4A4", opacity=0.6)
    fig.add_hline(y=-reference_value, line_dash="dot", line_color="#A4A4A4", opacity=0.6)
    fig.update_layout(
        title=dict(text="Scores Timeline", x=0.5, xanchor="center"),
        updatemenus=[
            {
                "type": "buttons",
                "showactive": False,
                "x": 0.01,
                "y": 1.16,
                "xanchor": "left",
                "yanchor": "top",
                "direction": "left",
                "buttons": [
                    {
                        "label": "Play",
                        "method": "animate",
                        "args": [
                            None,
                            {
                                "frame": {"duration": 120, "redraw": True},
                                "transition": {"duration": 0},
                                "fromcurrent": True,
                                "mode": "immediate",
                            },
                        ],
                    },
                    {
                        "label": "Pause",
                        "method": "animate",
                        "args": [
                            [None],
                            {
                                "frame": {"duration": 0, "redraw": False},
                                "transition": {"duration": 0},
                                "mode": "immediate",
                            },
                        ],
                    },
                ],
            }
        ],
        sliders=[
            {
                "active": 0,
                "x": 0.12,
                "len": 0.84,
                "y": 1.12,
                "xanchor": "left",
                "yanchor": "top",
                "currentvalue": {"prefix": "Game: ", "font": {"size": 14, "color": "#444444"}},
                "steps": [
                    {
                        "label": frame.name,
                        "method": "animate",
                        "args": [
                            [frame.name],
                            {
                                "frame": {"duration": 0, "redraw": True},
                                "transition": {"duration": 0},
                                "mode": "immediate",
                            },
                        ],
                    }
                    for frame in frames
                ],
            }
        ],
    )
    fig.update_xaxes(title_text="Year of Game Date", tickmode="array", tickvals=tick_values, gridcolor="#F4F4F4")
    fig.update_yaxes(title_text="Win Margin", gridcolor="#EFEFEF", zeroline=False)
    return configure_chart(fig, 560)


def render_scores_timeline_animation(df: pd.DataFrame, year_range: tuple[int, int], reference_value: int) -> None:
    fig = build_scores_timeline_animated(df, year_range, reference_value)
    html = pio.to_html(fig, include_plotlyjs=True, full_html=False, config={"displayModeBar": False})
    components.html(html, height=700)


def build_home_field_map(df: pd.DataFrame):
    if df.empty:
        return None

    counts = (
        df.groupby(["Home Field", "Result"])
        .size()
        .unstack(fill_value=0)
        .reindex(index=["Seattle", "Spokane", "Pullman"], fill_value=0)
        .reindex(columns=["UW", "WSU", "Tie"], fill_value=0)
        .reset_index()
    )
    counts["total_games"] = counts[["UW", "WSU", "Tie"]].sum(axis=1).astype(int)
    counts["lat"] = counts["Home Field"].map(lambda city: CITY_COORDINATES[city]["lat"])
    counts["lon"] = counts["Home Field"].map(lambda city: CITY_COORDINATES[city]["lon"])
    fmap = folium.Map(
        location=[47.35, -119.65],
        zoom_start=6,
        tiles="CartoDB positron",
        control_scale=False,
        zoom_control=False,
    )

    for row in counts.to_dict("records"):
        svg = build_pie_marker_svg(
            row["Home Field"],
            int(row["UW"]),
            int(row["WSU"]),
            int(row["Tie"]),
        )
        icon_size = [120, 120]
        tooltip_html = (
            f"<b>{row['Home Field']}</b><br>"
            f"UW: {int(row['UW'])}<br>"
            f"WSU: {int(row['WSU'])}<br>"
            f"Tie: {int(row['Tie'])}<br>"
            f"Total: {int(row['total_games'])}"
        )
        folium.Marker(
            location=[row["lat"], row["lon"]],
            icon=folium.DivIcon(
                html=svg,
                icon_size=icon_size,
                icon_anchor=[icon_size[0] // 2, 52],
            ),
            tooltip=tooltip_html,
        ).add_to(fmap)

    return fmap


def build_both_ranked_square(df: pd.DataFrame) -> go.Figure:
    counts = (
        df[df["Both Teams Ranked"] == "Yes"]
        .groupby("Result")
        .size()
        .reindex(["UW", "WSU"], fill_value=0)
    )
    counts = counts[counts > 0]
    if counts.empty:
        return empty_figure("Both Ranked", 210)

    fig = go.Figure(
        go.Scatter(
            x=counts.index,
            y=[""] * len(counts),
            mode="markers+text",
            marker=dict(
                symbol="square",
                size=[140] * len(counts),
                color=[RESULT_COLORS[key] for key in counts.index],
                line=dict(color="white", width=2),
            ),
            text=counts.astype(int).astype(str),
            textfont=dict(size=36, color="white"),
            hovertemplate="%{x}: %{text} games<extra></extra>",
        )
    )
    fig.update_layout(title=dict(text="Both Ranked", x=0.5, xanchor="center"))
    fig.update_xaxes(title_text="", showgrid=False)
    fig.update_yaxes(title_text="", showgrid=False, visible=False)
    return configure_chart(fig, 210)


def build_sankey(df: pd.DataFrame) -> go.Figure:
    chart_df = df[
        df["Both Teams Ranked"].isin(["Yes", "No"])
        & df["Result"].isin(["UW", "WSU"])
    ].copy()
    if chart_df.empty:
        return empty_figure("Sankey", 540)

    years = sorted(chart_df["Year"].dropna().astype(int).unique().tolist())
    ranked_statuses = ["Yes", "No"]
    results = ["UW", "WSU"]
    node_labels = [str(year) for year in years] + ranked_statuses + results
    node_colors = (
        [rgba("#A0CBE8", 0.95) for _ in years]
        + [rgba(RANKED_COLORS[status], 0.95) for status in ranked_statuses]
        + [rgba(RESULT_COLORS[result], 0.95) for result in results]
    )
    node_index = {label: idx for idx, label in enumerate(node_labels)}

    links_source: list[int] = []
    links_target: list[int] = []
    links_value: list[int] = []
    links_color: list[str] = []

    year_to_ranked = chart_df.groupby(["Year", "Both Teams Ranked"]).size().reset_index(name="games")
    for record in year_to_ranked.itertuples(index=False):
        links_source.append(node_index[str(int(record.Year))])
        links_target.append(node_index[record[1]])
        links_value.append(int(record.games))
        links_color.append(rgba(RANKED_COLORS[record[1]], 0.35))

    ranked_to_result = chart_df.groupby(["Both Teams Ranked", "Result"]).size().reset_index(name="games")
    for record in ranked_to_result.itertuples(index=False):
        links_source.append(node_index[record[0]])
        links_target.append(node_index[record[1]])
        links_value.append(int(record.games))
        links_color.append(rgba(RESULT_COLORS[record[1]], 0.35))

    fig = go.Figure(
        go.Sankey(
            arrangement="snap",
            node=dict(
                label=node_labels,
                color=node_colors,
                pad=10,
                thickness=16,
                line=dict(color="white", width=1),
            ),
            link=dict(
                source=links_source,
                target=links_target,
                value=links_value,
                color=links_color,
            ),
        )
    )
    fig.update_layout(title=dict(text="Sankey", x=0.5, xanchor="center"))
    return configure_chart(fig, 540)


def render_result_legend() -> None:
    legend = """
    <div style="padding-top:1rem;">
      <div style="font-size:0.95rem;color:#666;margin-bottom:0.5rem;">Result</div>
      <div style="display:flex;gap:0.5rem;align-items:center;margin-bottom:0.35rem;">
        <span style="width:12px;height:12px;border-radius:50%;background:#4E79A7;display:inline-block;"></span>
        <span style="font-size:0.95rem;">Tie</span>
      </div>
      <div style="display:flex;gap:0.5rem;align-items:center;margin-bottom:0.35rem;">
        <span style="width:12px;height:12px;border-radius:50%;background:#B07AA1;display:inline-block;"></span>
        <span style="font-size:0.95rem;">UW</span>
      </div>
      <div style="display:flex;gap:0.5rem;align-items:center;">
        <span style="width:12px;height:12px;border-radius:50%;background:#E15759;display:inline-block;"></span>
        <span style="font-size:0.95rem;">WSU</span>
      </div>
    </div>
    """
    st.markdown(legend, unsafe_allow_html=True)


def render_summary_metrics(df: pd.DataFrame) -> None:
    total_games = len(df)
    uw_wins = int((df["Result"] == "UW").sum())
    wsu_wins = int((df["Result"] == "WSU").sum())
    ties = int((df["Result"] == "Tie").sum())
    metrics = st.columns(4)
    metrics[0].metric("Games", total_games)
    metrics[1].metric("UW Wins", uw_wins)
    metrics[2].metric("WSU Wins", wsu_wins)
    metrics[3].metric("Ties", ties)


def main() -> None:
    st.markdown(
        """
        <style>
          .block-container {padding-top: 1.25rem; padding-bottom: 2rem;}
          [data-testid="stMetric"] {background: transparent;}
          div[role="radiogroup"] > label {
            background: #f5f5f5;
            border: 1px solid #d9d9d9;
            border-radius: 10px;
            padding: 0.5rem 0.9rem;
            margin-right: 0.5rem;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    df = load_data()
    logo_bytes = load_logo()
    min_year = int(df["Year"].min())
    max_year = int(df["Year"].max())
    result_options = ["UW", "WSU", "Tie"]
    home_field_options = sorted(df["Home Field"].dropna().unique().tolist())
    ranked_options = ["Yes", "No", "Pre-AP Poll"]
    ot_options = sorted(df["OT"].dropna().unique().tolist())

    with st.sidebar:
        st.header("Dashboard Filters")
        global_year_range = st.slider(
            "Year Range",
            min_value=min_year,
            max_value=max_year,
            value=(min_year, max_year),
        )
        selected_results = st.multiselect("Result", result_options, default=result_options)
        selected_home_fields = st.multiselect("Home Field", home_field_options, default=home_field_options)
        selected_ranked = st.multiselect("Both Teams Ranked", ranked_options, default=ranked_options)
        selected_ot = st.multiselect("Overtime", ot_options, default=ot_options)
        st.caption("All visuals update from these filters.")

    if not selected_results:
        selected_results = result_options
    if not selected_home_fields:
        selected_home_fields = home_field_options
    if not selected_ranked:
        selected_ranked = ranked_options
    if not selected_ot:
        selected_ot = ot_options

    filtered_df = apply_filters(
        df,
        global_year_range,
        selected_results,
        selected_home_fields,
        selected_ranked,
        selected_ot,
    )

    dashboard_view = st.radio(
        "Dashboard View",
        [
            "Apple Cup: The Big Picture",
            "Apple Cup: Games That Defined the Rivalry",
        ],
        horizontal=True,
    )

    if dashboard_view == "Apple Cup: The Big Picture":
        render_summary_metrics(filtered_df)
        top_left, top_middle, top_right = st.columns([4, 2, 3])
        with top_left:
            st.markdown("**Live Filters Applied**")
            st.caption(
                f"{global_year_range[0]} to {global_year_range[1]} | "
                f"{', '.join(selected_results)} | "
                f"{', '.join(selected_home_fields)}"
            )
            render_result_legend()
        with top_middle:
            st.write("")
        with top_right:
            if logo_bytes is not None:
                st.image(logo_bytes, use_container_width=True)

        chart_cols = st.columns([5, 4, 5, 6])
        reference_value = int(st.session_state.get("sov_parameter", 48))
        with chart_cols[0]:
            st.plotly_chart(
                build_total_wins_donut(filtered_df, "Total Wins / Tie", "Total Games<br>Won"),
                use_container_width=True,
                config={"displayModeBar": False},
            )
        with chart_cols[1]:
            st.plotly_chart(
                build_probability_chart(filtered_df),
                use_container_width=True,
                config={"displayModeBar": False},
            )

        with chart_cols[2]:
            st.plotly_chart(
                build_largest_win_margin_chart(filtered_df, reference_value),
                use_container_width=True,
                config={"displayModeBar": False},
            )
        with chart_cols[3]:
            st.plotly_chart(
                build_sov_chart(filtered_df),
                use_container_width=True,
                config={"displayModeBar": False},
            )

        reference_value = st.slider("SOV Parameter", min_value=-27, max_value=48, value=48, key="sov_parameter")
        timeline_df = filtered_df[
            filtered_df["Result"].isin(["UW", "WSU"])
            & filtered_df["Year"].between(global_year_range[0], global_year_range[1])
        ].sort_values("Game Date").reset_index(drop=True)

        if timeline_df.empty:
            render_scores_timeline_animation(filtered_df, global_year_range, reference_value)
        else:
            st.caption("Use Play to reveal one game at a time from 1901 through 2025 directly in the browser.")
            render_scores_timeline_animation(filtered_df, global_year_range, reference_value)
        with st.expander("Filtered Game Log", expanded=False):
            display_columns = [
                "Date",
                "Home Field",
                "Result",
                "Winner",
                "Score",
                "WIn Margin",
                "UW SOV",
                "WSU SOV",
                "Both Teams Ranked",
                "OT",
            ]
            st.dataframe(
                filtered_df.sort_values("Game Date", ascending=False)[display_columns],
                use_container_width=True,
                hide_index=True,
            )

    else:
        rivalry_threshold = st.slider(
            "Rivalry Threshold (absolute win margin)",
            min_value=0,
            max_value=48,
            value=0,
            help="Default 24 matches the Tableau story-point view.",
        )
        rivalry_df = filtered_df[filtered_df["Abs Win Margin"] >= rivalry_threshold].copy()
        render_summary_metrics(rivalry_df)
        st.caption("This tab updates from the sidebar filters and the rivalry threshold.")
        top_left, top_right = st.columns([1, 1])
        with top_left:
            st.plotly_chart(
                build_total_wins_donut(rivalry_df, "Total Games", "Total Games<br>Won", height=340),
                use_container_width=True,
                config={"displayModeBar": False},
            )
        with top_right:
            st.markdown("**Map**")
            st.caption("Each city now uses a true pie marker: UW in purple, WSU in red, and ties in blue.")
            st.markdown(
                """
                <div style="display:flex;gap:1rem;align-items:center;margin-bottom:0.5rem;font-size:0.92rem;color:#555;">
                  <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#B07AA1;margin-right:0.3rem;"></span>UW</span>
                  <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#E15759;margin-right:0.3rem;"></span>WSU</span>
                  <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#4E79A7;margin-right:0.3rem;"></span>Tie</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            rivalry_map = build_home_field_map(rivalry_df)
            if rivalry_map is None:
                st.info("No map data for the current filters.")
            else:
                st_folium(rivalry_map, use_container_width=True, height=430, returned_objects=[])

        st.plotly_chart(
            build_both_ranked_square(rivalry_df),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        st.plotly_chart(
            build_sankey(rivalry_df),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        with st.expander("Rivalry Games Table", expanded=False):
            st.dataframe(
                rivalry_df.sort_values(["Abs Win Margin", "Game Date"], ascending=[False, False])[
                    [
                        "Date",
                        "Year",
                        "Home Field",
                        "Result",
                        "Score",
                        "WIn Margin",
                        "Both Teams Ranked",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )


if __name__ == "__main__":
    main()
