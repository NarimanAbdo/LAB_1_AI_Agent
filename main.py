"""
Lab Assignment 1: Working with a Public REST API Using Python

Fetches near-Earth asteroid data from NASA's NeoWs REST API
(https://api.nasa.gov), parses the JSON response, stores each asteroid
in a small class, prints a formatted summary, and exports the data
to an Excel file (asteroids.xlsx).

The API key is read from the NASA_API_KEY environment variable, which
is loaded from the .env file.
"""

import os
import sys
from dataclasses import dataclass
from datetime import date, timedelta

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

API_URL = "https://api.nasa.gov/neo/rest/v1/feed"
EXCEL_FILE = "asteroids.xlsx"


@dataclass
class Asteroid:
    """Stores the useful details of one near-Earth asteroid."""

    name: str
    approach_date: str
    diameter_min_m: float
    diameter_max_m: float
    velocity_kmh: float
    miss_distance_km: float
    is_hazardous: bool

    @classmethod
    def from_json(cls, data: dict) -> "Asteroid":
        """Build an Asteroid from one entry of the API's JSON response."""
        approach = data["close_approach_data"][0]
        diameter = data["estimated_diameter"]["meters"]
        name = data["name"]
        if name.startswith("(") and name.endswith(")"):
            name = name[1:-1]  # "(2015 VH65)" -> "2015 VH65"
        return cls(
            name=name,
            approach_date=approach["close_approach_date"],
            diameter_min_m=diameter["estimated_diameter_min"],
            diameter_max_m=diameter["estimated_diameter_max"],
            velocity_kmh=float(approach["relative_velocity"]["kilometers_per_hour"]),
            miss_distance_km=float(approach["miss_distance"]["kilometers"]),
            is_hazardous=data["is_potentially_hazardous_asteroid"],
        )

    @property
    def average_diameter_m(self) -> float:
        return (self.diameter_min_m + self.diameter_max_m) / 2


def get_api_key() -> str:
    """Load the API key from .env / environment variables."""
    load_dotenv()
    api_key = os.getenv("NASA_API_KEY")
    if not api_key:
        sys.exit("Error: NASA_API_KEY is not set. Add it to your .env file.")
    return api_key


def fetch_asteroids(api_key: str, start: date, end: date) -> list[Asteroid]:
    """Call the API and turn the JSON response into Asteroid objects."""
    params = {
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "api_key": api_key,
    }
    try:
        response = requests.get(API_URL, params=params, timeout=15)
        response.raise_for_status()
    except requests.RequestException as error:
        sys.exit(f"Error: request to the API failed: {error}")

    data = response.json()  # parse the JSON body into Python dicts/lists
    asteroids = [
        Asteroid.from_json(item)
        for day in data["near_earth_objects"].values()
        for item in day
    ]
    return sorted(asteroids, key=lambda a: a.miss_distance_km)


def print_summary(asteroids: list[Asteroid], start: date, end: date) -> None:
    """Print a formatted summary of the retrieved asteroids."""
    width = 98
    line = "=" * width
    print(line)
    print(f"NEAR-EARTH ASTEROIDS  |  {start} to {end}".center(width))
    print(line)

    if not asteroids:
        print("No asteroids found for this period.")
        return

    print(
        f"{'Name':<26}{'Date':<12}{'Diameter (m)':>14}"
        f"{'Speed (km/h)':>16}{'Miss dist. (km)':>18}{'Hazardous':>12}"
    )
    print("-" * width)
    for a in asteroids:
        print(
            f"{a.name:<26}{a.approach_date:<12}{a.average_diameter_m:>14,.1f}"
            f"{a.velocity_kmh:>16,.0f}{a.miss_distance_km:>18,.0f}"
            f"{'YES' if a.is_hazardous else 'no':>12}"
        )

    hazardous = [a for a in asteroids if a.is_hazardous]
    closest = asteroids[0]
    largest = max(asteroids, key=lambda a: a.average_diameter_m)
    fastest = max(asteroids, key=lambda a: a.velocity_kmh)

    print(line)
    print("SUMMARY")
    print(f"  Total asteroids       : {len(asteroids)}")
    print(f"  Potentially hazardous : {len(hazardous)}")
    print(f"  Closest approach      : {closest.name} ({closest.miss_distance_km:,.0f} km)")
    print(f"  Largest               : {largest.name} (~{largest.average_diameter_m:,.0f} m)")
    print(f"  Fastest               : {fastest.name} ({fastest.velocity_kmh:,.0f} km/h)")
    print(line)


def export_to_excel(asteroids: list[Asteroid], filename: str) -> None:
    """Save the asteroids to an Excel file: one 'Asteroids' sheet and one 'Summary' sheet."""
    workbook = Workbook()

    # ---- Sheet 1: one row per asteroid ----
    sheet = workbook.active
    sheet.title = "Asteroids"
    headers = [
        "Name", "Approach Date", "Min Diameter (m)", "Max Diameter (m)",
        "Avg Diameter (m)", "Speed (km/h)", "Miss Distance (km)", "Hazardous",
    ]
    sheet.append(headers)
    for a in asteroids:
        sheet.append([
            a.name, a.approach_date, a.diameter_min_m, a.diameter_max_m,
            a.average_diameter_m, a.velocity_kmh, a.miss_distance_km,
            "YES" if a.is_hazardous else "no",
        ])

    # Bold white text on a blue background for the header row
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    # Number formats: 1 decimal for diameters, thousands separator for speed/distance
    for row in sheet.iter_rows(min_row=2):
        for cell in row[2:5]:
            cell.number_format = "#,##0.0"
        for cell in row[5:7]:
            cell.number_format = "#,##0"

    # Make the columns wide enough to read, and keep the header visible when scrolling
    for column, width in zip("ABCDEFGH", [24, 15, 18, 18, 18, 15, 20, 12]):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"

    # ---- Sheet 2: summary figures ----
    summary = workbook.create_sheet("Summary")
    closest = min(asteroids, key=lambda a: a.miss_distance_km)
    largest = max(asteroids, key=lambda a: a.average_diameter_m)
    fastest = max(asteroids, key=lambda a: a.velocity_kmh)
    summary.append(["Metric", "Value"])
    summary.append(["Total asteroids", len(asteroids)])
    summary.append(["Potentially hazardous", sum(a.is_hazardous for a in asteroids)])
    summary.append(["Closest approach", f"{closest.name} ({closest.miss_distance_km:,.0f} km)"])
    summary.append(["Largest", f"{largest.name} (~{largest.average_diameter_m:,.0f} m)"])
    summary.append(["Fastest", f"{fastest.name} ({fastest.velocity_kmh:,.0f} km/h)"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 40

    try:
        workbook.save(filename)
    except PermissionError:
        sys.exit(f"Error: cannot write {filename}. Close it in Excel and run again.")
    print(f"Data exported to {filename}")


def main() -> None:
    api_key = get_api_key()
    start = date.today()
    end = start + timedelta(days=2)
    asteroids = fetch_asteroids(api_key, start, end)
    print_summary(asteroids, start, end)
    if asteroids:
        export_to_excel(asteroids, EXCEL_FILE)


if __name__ == "__main__":
    main()
