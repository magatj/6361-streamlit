from __future__ import annotations

import json
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
from plotly.utils import PlotlyJSONEncoder
from streamlit_folium import st_folium


st.set_page_config(
    page_title="Apple Cup Dashboard",
    page_icon="🏈",
    layout="wide",
)


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = Path(__file__).parent / "Apple_Cup_History_AP_Rankings.xlsx"
WEATHER_FILE = BASE_DIR / "apple_cup_daily_weather.csv"
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
    weather_df = load_weather_data()
    if weather_df.empty:
        df["Mean Temperature (F)"] = pd.NA
        df["Precipitation (in)"] = pd.NA
        df["Snowfall (in)"] = pd.NA
    else:
        df = df.merge(weather_df, on=["Game Date", "Home Field"], how="left")
    df["Weather Available"] = df["Mean Temperature (F)"].notna()
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


@st.cache_data(show_spinner=False)
def load_weather_data() -> pd.DataFrame:
    if not WEATHER_FILE.exists():
        return pd.DataFrame(columns=["Game Date", "Home Field", "Mean Temperature (F)", "Precipitation (in)", "Snowfall (in)"])

    weather_df = pd.read_csv(WEATHER_FILE)
    if weather_df.empty:
        return pd.DataFrame(columns=["Game Date", "Home Field", "Mean Temperature (F)", "Precipitation (in)", "Snowfall (in)"])

    weather_df["Game Date"] = pd.to_datetime(weather_df["date"], errors="coerce")
    return weather_df.rename(
        columns={
            "home_field": "Home Field",
            "mean_temperature_f": "Mean Temperature (F)",
            "precipitation_in": "Precipitation (in)",
            "snowfall_in": "Snowfall (in)",
        }
    )[["Game Date", "Home Field", "Mean Temperature (F)", "Precipitation (in)", "Snowfall (in)"]]


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


def build_probability_chart(df: pd.DataFrame, title: str = "Winning Probability", height: int = 320) -> go.Figure:
    if df.empty:
        return empty_figure(title, height)

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
    fig.update_layout(title=dict(text=title, x=0.5, xanchor="center"))
    fig.update_xaxes(range=[0, 1], tickformat=".0%", title_text="Probability", gridcolor="#EFEFEF", zeroline=False)
    fig.update_yaxes(title_text="")
    return configure_chart(fig, height)


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
        .groupby("Result")["SOV Calc"]
        .agg(["mean", "count"])
        .reindex(["UW", "WSU"])
        .reset_index()
    )
    if chart_df.empty:
        return empty_figure("SOV Calc", 320)

    chart_df["band_60"] = chart_df["mean"] * 0.60
    chart_df["band_80"] = chart_df["mean"] * 0.80

    fig = go.Figure(
        go.Bar(
            x=chart_df["Result"],
            y=chart_df["mean"],
            marker=dict(color=[RESULT_COLORS[key] for key in chart_df["Result"]]),
            width=0.58,
            text=[f"{value:.1f}" for value in chart_df["mean"]],
            textposition="outside",
            customdata=chart_df[["band_60", "band_80", "count"]].to_numpy(),
            hovertemplate=(
                "%{x}<br>"
                "Average SOV: %{y:.2f}<br>"
                "60% of Average: %{customdata[0]:.2f}<br>"
                "80% of Average: %{customdata[1]:.2f}<br>"
                "Games: %{customdata[2]:.0f}<extra></extra>"
            ),
        )
    )

    for idx, row in chart_df.iterrows():
        fig.add_shape(
            type="rect",
            xref="x",
            yref="y",
            x0=idx - 0.33,
            x1=idx + 0.33,
            y0=float(row["band_60"]),
            y1=float(row["band_80"]),
            fillcolor="rgba(70, 70, 70, 0.35)",
            line=dict(color="rgba(40, 40, 40, 0.95)", width=1),
            layer="above",
        )
        fig.add_shape(
            type="line",
            xref="x",
            yref="y",
            x0=idx - 0.33,
            x1=idx + 0.33,
            y0=float(row["band_80"]),
            y1=float(row["band_80"]),
            line=dict(color="rgba(20, 20, 20, 0.95)", width=2),
            layer="above",
        )
        fig.add_shape(
            type="line",
            xref="x",
            yref="y",
            x0=idx - 0.33,
            x1=idx + 0.33,
            y0=float(row["band_60"]),
            y1=float(row["band_60"]),
            line=dict(color="rgba(20, 20, 20, 0.95)", width=2),
            layer="above",
        )
        fig.add_annotation(
            x=row["Result"],
            y=float(row["band_80"]),
            text="80% of Average",
            yshift=8,
            showarrow=False,
            font=dict(size=11, color="#7A7A7A"),
            xshift=-36,
        )
        fig.add_annotation(
            x=row["Result"],
            y=float(row["band_60"]),
            text="60% of Average",
            yshift=-8,
            showarrow=False,
            font=dict(size=11, color="#5A5A5A"),
            xshift=-36,
        )

    y_max = chart_df["mean"].max() * 1.18
    fig.update_layout(
        title=dict(
            text="Average Points Win Margin<br><sup>Based on Strength of Victory (SOV)</sup>",
            x=0.5,
            xanchor="center",
        ),
    )
    fig.update_xaxes(title_text="")
    fig.update_yaxes(title_text="SOV Calc", range=[0, y_max], gridcolor="#EFEFEF", zeroline=False)
    return configure_chart(fig, 360)


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

    chart_df["Era"] = (
        (chart_df["Year"].astype(int) // 10) * 10
    ).astype(int).astype(str) + "s"
    eras = sorted(chart_df["Era"].dropna().unique().tolist())
    ranked_statuses = ["Yes", "No"]
    results = ["UW", "WSU"]

    era_totals = chart_df.groupby("Era").size().reindex(eras, fill_value=0)
    ranked_totals = chart_df.groupby("Both Teams Ranked").size().reindex(ranked_statuses, fill_value=0)
    result_totals = chart_df.groupby("Result").size().reindex(results, fill_value=0)
    node_labels = (
        [f"{era}\n{int(era_totals[era])}" for era in eras]
        + [f"{status}\n{int(ranked_totals[status])}" for status in ranked_statuses]
        + [f"{result}\n{int(result_totals[result])}" for result in results]
    )
    node_colors = (
        [rgba("#9FB7C9", 0.95) for _ in eras]
        + [rgba(RANKED_COLORS[status], 0.95) for status in ranked_statuses]
        + [rgba(RESULT_COLORS[result], 0.95) for result in results]
    )
    node_keys = eras + ranked_statuses + results
    node_index = {label: idx for idx, label in enumerate(node_keys)}

    links_source: list[int] = []
    links_target: list[int] = []
    links_value: list[int] = []
    links_color: list[str] = []

    era_to_ranked = chart_df.groupby(["Era", "Both Teams Ranked"]).size().reset_index(name="games")
    for record in era_to_ranked.itertuples(index=False):
        links_source.append(node_index[record[0]])
        links_target.append(node_index[record[1]])
        links_value.append(int(record.games))
        links_color.append(rgba(RANKED_COLORS[record[1]], 0.52 if record[1] == "Yes" else 0.30))

    ranked_to_result = chart_df.groupby(["Both Teams Ranked", "Result"]).size().reset_index(name="games")
    for record in ranked_to_result.itertuples(index=False):
        links_source.append(node_index[record[0]])
        links_target.append(node_index[record[1]])
        links_value.append(int(record.games))
        links_color.append(rgba(RESULT_COLORS[record[1]], 0.42))

    def evenly_spaced_positions(count: int, top: float = 0.04, bottom: float = 0.96) -> list[float]:
        if count == 1:
            return [0.5]
        step = (bottom - top) / (count - 1)
        return [top + idx * step for idx in range(count)]

    node_x = [0.03] * len(eras) + [0.50] * len(ranked_statuses) + [0.90] * len(results)
    node_y = evenly_spaced_positions(len(eras), 0.05, 0.95) + [0.28, 0.72] + [0.32, 0.68]

    fig = go.Figure(
        go.Sankey(
            arrangement="fixed",
            node=dict(
                label=node_labels,
                color=node_colors,
                pad=11,
                thickness=20,
                line=dict(color="rgba(255,255,255,0.92)", width=1.5),
                x=node_x,
                y=node_y,
                hovertemplate="%{label}<extra></extra>",
            ),
            link=dict(
                source=links_source,
                target=links_target,
                value=links_value,
                color=links_color,
                hovertemplate="%{value} games<extra></extra>",
            ),
        )
    )

    fig.update_layout(
        title=dict(
            text="Rivalry Flow<br><sup>Era to ranking status to winner</sup>",
            x=0.5,
            xanchor="center",
        ),
        margin=dict(l=8, r=8, t=124, b=36),
        annotations=[
            dict(x=0.03, y=1.06, xref="paper", yref="paper", text="Era", showarrow=False, font=dict(size=13, color="#606060")),
            dict(x=0.50, y=1.06, xref="paper", yref="paper", text="Both Ranked", showarrow=False, font=dict(size=13, color="#606060")),
            dict(x=0.90, y=1.06, xref="paper", yref="paper", text="Winner", showarrow=False, font=dict(size=13, color="#606060")),
        ],
    )
    fig = configure_chart(fig, 620)
    fig.update_layout(margin=dict(l=8, r=8, t=124, b=36))
    return fig


def render_interactive_sankey(df: pd.DataFrame) -> None:
    fig = build_sankey(df)
    fig_dict = json.loads(json.dumps(fig, cls=PlotlyJSONEncoder))
    fig_json = json.dumps(fig_dict)
    html = f"""
    <div style="font-family: Arial, sans-serif; margin-bottom: 0.4rem; color: #666; font-size: 0.92rem;">
      Click a flow or node to trace its connected path from left to right. Double-click the chart to reset.
    </div>
    <div id="rivalry-sankey" style="width:100%;height:640px;"></div>
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <script>
      const figure = {fig_json};
      const gd = document.getElementById('rivalry-sankey');
      Plotly.newPlot(gd, figure.data, figure.layout, {{displayModeBar:false, responsive:true}}).then(() => {{
        const nodeX = gd.data[0].node.x.slice();
        const nodeColorsOriginal = gd.data[0].node.color.slice();
        const linkColorsOriginal = gd.data[0].link.color.slice();
        const sources = gd.data[0].link.source.slice();
        const targets = gd.data[0].link.target.slice();
        const nodeDim = 'rgba(215,215,215,0.45)';
        const linkDim = 'rgba(220,220,220,0.10)';

        function resetColors() {{
          Plotly.restyle(gd, {{
            'node.color': [nodeColorsOriginal],
            'link.color': [linkColorsOriginal]
          }}, [0]);
        }}

        function highlightFromNode(nodeIndex) {{
          const highlightNodes = new Set([nodeIndex]);
          const highlightLinks = new Set();
          const xPos = nodeX[nodeIndex];

          if (xPos < 0.2) {{
            sources.forEach((src, i) => {{
              if (src === nodeIndex) {{
                highlightLinks.add(i);
                highlightNodes.add(targets[i]);
                sources.forEach((src2, j) => {{
                  if (src2 === targets[i]) {{
                    highlightLinks.add(j);
                    highlightNodes.add(targets[j]);
                  }}
                }});
              }}
            }});
          }} else if (xPos < 0.8) {{
            sources.forEach((src, i) => {{
              if (src === nodeIndex) {{
                highlightLinks.add(i);
                highlightNodes.add(targets[i]);
              }}
            }});
            targets.forEach((tgt, i) => {{
              if (tgt === nodeIndex) {{
                highlightLinks.add(i);
                highlightNodes.add(sources[i]);
              }}
            }});
          }} else {{
            targets.forEach((tgt, i) => {{
              if (tgt === nodeIndex) {{
                highlightLinks.add(i);
                highlightNodes.add(sources[i]);
                targets.forEach((tgt2, j) => {{
                  if (tgt2 === sources[i]) {{
                    highlightLinks.add(j);
                    highlightNodes.add(sources[j]);
                  }}
                }});
              }}
            }});
          }}

          const nodeColors = nodeColorsOriginal.map((color, i) => highlightNodes.has(i) ? color : nodeDim);
          const linkColors = linkColorsOriginal.map((color, i) => highlightLinks.has(i) ? color : linkDim);
          Plotly.restyle(gd, {{
            'node.color': [nodeColors],
            'link.color': [linkColors]
          }}, [0]);
        }}

        function highlightFromLink(linkIndex) {{
          const highlightNodes = new Set();
          const highlightLinks = new Set([linkIndex]);
          const src = sources[linkIndex];
          const tgt = targets[linkIndex];
          highlightNodes.add(src);
          highlightNodes.add(tgt);

          if (nodeX[src] < 0.2) {{
            sources.forEach((s, i) => {{
              if (s === tgt) {{
                highlightLinks.add(i);
                highlightNodes.add(targets[i]);
              }}
            }});
          }} else {{
            targets.forEach((t, i) => {{
              if (t === src) {{
                highlightLinks.add(i);
                highlightNodes.add(sources[i]);
              }}
            }});
          }}

          const nodeColors = nodeColorsOriginal.map((color, i) => highlightNodes.has(i) ? color : nodeDim);
          const linkColors = linkColorsOriginal.map((color, i) => highlightLinks.has(i) ? color : linkDim);
          Plotly.restyle(gd, {{
            'node.color': [nodeColors],
            'link.color': [linkColors]
          }}, [0]);
        }}

        gd.on('plotly_click', function(eventData) {{
          const pt = eventData.points[0];
          if (typeof pt.source !== 'undefined' && typeof pt.target !== 'undefined') {{
            highlightFromLink(pt.pointNumber);
          }} else {{
            highlightFromNode(pt.pointNumber);
          }}
        }});

        gd.on('plotly_doubleclick', function() {{
          resetColors();
        }});
      }});
    </script>
    """
    components.html(html, height=700)


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


def weather_temperature_bounds(df: pd.DataFrame) -> tuple[int, int] | None:
    weather_df = df[df["Weather Available"]].copy()
    if weather_df.empty:
        return None

    min_temp = int(math.floor(weather_df["Mean Temperature (F)"].min()))
    max_temp = int(math.ceil(weather_df["Mean Temperature (F)"].max()))
    if min_temp == max_temp:
        max_temp += 1
    return min_temp, max_temp


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
        weather_bounds = weather_temperature_bounds(filtered_df)
        if weather_bounds is not None:
            weather_temp_range = st.slider(
                "Weather Temperature Range (°F)",
                min_value=weather_bounds[0],
                max_value=weather_bounds[1],
                value=weather_bounds,
                key="weather_temp_range",
            )
            weather_probability_df = filtered_df[
                filtered_df["Weather Available"]
                & filtered_df["Mean Temperature (F)"].between(weather_temp_range[0], weather_temp_range[1])
            ].copy()
            rain_games = int((weather_probability_df["Precipitation (in)"].fillna(0) > 0).sum())
            snow_games = int((weather_probability_df["Snowfall (in)"].fillna(0) > 0).sum())
            avg_temp = weather_probability_df["Mean Temperature (F)"].mean()
            st.caption(
                "Historical weather sidecar data is available for Apple Cup sites from 1940 onward. "
                "Use the temperature range slider to update the weather-based winning probability view."
            )
            weather_cols = st.columns([5, 2, 2, 2])
            with weather_cols[0]:
                st.plotly_chart(
                    build_probability_chart(
                        weather_probability_df,
                        title=f"Winning Probability in {weather_temp_range[0]}°F to {weather_temp_range[1]}°F Games",
                    ),
                    use_container_width=True,
                    config={"displayModeBar": False},
                )
            with weather_cols[1]:
                st.metric("Weather-Matched Games", len(weather_probability_df))
            with weather_cols[2]:
                st.metric("Avg Temp", "n/a" if pd.isna(avg_temp) else f"{avg_temp:.1f}°F")
            with weather_cols[3]:
                st.metric("Rain / Snow Games", f"{rain_games} / {snow_games}")
        else:
            st.info("No weather-matched games are available for the current filters.")

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
                "Mean Temperature (F)",
                "Precipitation (in)",
                "Snowfall (in)",
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
        render_interactive_sankey(rivalry_df)
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
