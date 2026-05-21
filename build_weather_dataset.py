from __future__ import annotations

import csv
import re
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
from meteostat import Daily, Point


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "Apple_Cup_History_AP_Rankings.xlsx"
OUTPUT_FILE = BASE_DIR / "apple_cup_daily_weather.csv"
CITY_COORDINATES = {
    "Seattle": {"lat": 47.6062, "lon": -122.3321},
    "Spokane": {"lat": 47.6588, "lon": -117.4260},
    "Pullman": {"lat": 46.7298, "lon": -117.1817},
}
NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def celsius_to_fahrenheit(value: float | None) -> float | None:
    if value is None:
        return None
    return round((value * 9 / 5) + 32, 2)


def millimeters_to_inches(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value / 25.4, 3)


def millimeters_to_inches(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value / 25.4, 3)


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
        rel_map = {relation.attrib["Id"]: relation.attrib["Target"] for relation in rel_root}

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


def apple_cup_games() -> list[dict[str, datetime]]:
    rows = workbook_rows(DATA_FILE)
    header = rows[0]
    games: list[dict[str, datetime]] = []
    seen: set[tuple[str, str]] = set()
    for row in rows[1:]:
        normalized = row + [""] * (len(header) - len(row))
        record = dict(zip(header, normalized))
        game_date = pd.to_datetime(record["Date"], errors="coerce")
        home_field = record["Home Field"]
        if pd.isna(game_date) or home_field not in CITY_COORDINATES:
            continue
        dedupe_key = (game_date.strftime("%Y-%m-%d"), home_field)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        games.append({"game_date": game_date.to_pydatetime(), "home_field": home_field})
    return sorted(games, key=lambda game: (game["game_date"], game["home_field"]))


def fetch_game_weather(game_date: datetime, city: str) -> dict[str, str | float | None]:
    coordinates = CITY_COORDINATES[city]
    weather = Daily(Point(coordinates["lat"], coordinates["lon"]), game_date, game_date).fetch()

    if weather.empty:
        return {
            "date": game_date.strftime("%Y-%m-%d"),
            "home_field": city,
            "mean_temperature_f": None,
            "precipitation_in": None,
            "snowfall_in": None,
        }

    weather_row = weather.iloc[0]
    avg_temp_c = weather_row.get("tavg")
    if pd.isna(avg_temp_c) and not pd.isna(weather_row.get("tmin")) and not pd.isna(weather_row.get("tmax")):
        avg_temp_c = (weather_row["tmin"] + weather_row["tmax"]) / 2

    precipitation_mm = weather_row.get("prcp")
    snowfall_mm = weather_row.get("snow")
    return {
        "date": game_date.strftime("%Y-%m-%d"),
        "home_field": city,
        "mean_temperature_f": None if pd.isna(avg_temp_c) else celsius_to_fahrenheit(float(avg_temp_c)),
        "precipitation_in": None if pd.isna(precipitation_mm) else millimeters_to_inches(float(precipitation_mm)),
        "snowfall_in": None if pd.isna(snowfall_mm) else millimeters_to_inches(float(snowfall_mm)),
    }


def main() -> None:
    weather_rows = [fetch_game_weather(game["game_date"], game["home_field"]) for game in apple_cup_games()]

    weather_rows.sort(key=lambda row: (str(row["date"]), str(row["home_field"])))
    with OUTPUT_FILE.open("w", newline="", encoding="utf-8") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=["date", "home_field", "mean_temperature_f", "precipitation_in", "snowfall_in"],
        )
        writer.writeheader()
        writer.writerows(weather_rows)

    print(f"Saved {len(weather_rows)} weather rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
