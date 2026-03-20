"""Utility for geocoding addresses in Excel files using Nominatim via geopy."""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Callable, Iterable, Optional

import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
from geopy.exc import GeocoderServiceError, GeocoderTimedOut
from geopy.extra.rate_limiter import RateLimiter
from geopy.geocoders import Nominatim

DEFAULT_USER_AGENT = "georef-app"
DEFAULT_SUFFIX = "_geocoded"
MIN_DELAY_SECONDS = 1.1

ProgressCallback = Callable[[int, int, str], None]
StopCheck = Callable[[], bool]


class AddressColumnError(ValueError):
    """Raised when an address column is missing or cannot be determined."""


class GeocodingCancelledError(Exception):
    """Raised when geocoding is cancelled by the user."""


def get_excel_columns(excel_path: Path) -> list[str]:
    """Return the column names from the first sheet of an Excel file."""
    path = Path(excel_path).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")
    frame = pd.read_excel(path, nrows=0)
    return [str(col) for col in frame.columns]


def infer_address_column(columns: Iterable[str]) -> Optional[str]:
    """Infer the address column by looking for the word 'address' (case-insensitive)."""
    for column in columns:
        if "address" in str(column).strip().lower():
            return column
    return None


def geocode_excel_file(
    excel_path: Path,
    address_column: Optional[str] = None,
    user_agent: str = DEFAULT_USER_AGENT,
    min_delay_seconds: float = MIN_DELAY_SECONDS,
    progress_callback: Optional[ProgressCallback] = None,
    stop_check: Optional[StopCheck] = None,
) -> Path:
    """Geocode addresses in an Excel file and write a new Excel file with coordinates."""
    path = Path(excel_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    frame = pd.read_excel(path)
    if frame.empty:
        raise ValueError("The selected Excel file does not contain any data rows.")

    column_name = address_column or infer_address_column(frame.columns)
    if column_name is None or column_name not in frame.columns:
        raise AddressColumnError(
            "No address column specified and unable to infer one. "
            "Please choose the correct column explicitly."
        )

    geolocator = Nominatim(user_agent=user_agent, timeout=10)
    geocode = RateLimiter(
        geolocator.geocode,
        min_delay_seconds=max(min_delay_seconds, 1.0),
        error_wait_seconds=2.0,
        swallow_exceptions=True,
    )

    total_rows = len(frame)
    latitudes: list[Optional[float]] = []
    longitudes: list[Optional[float]] = []
    matched_addresses: list[Optional[str]] = []

    for index, value in enumerate(frame[column_name]):
        # Check if stop was requested
        if stop_check and stop_check():
            raise GeocodingCancelledError("Geocoding was cancelled by the user.")

        address_text = "" if pd.isna(value) else str(value).strip()
        if progress_callback:
            progress_callback(index + 1, total_rows, address_text)

        if not address_text:
            latitudes.append(None)
            longitudes.append(None)
            matched_addresses.append(None)
            continue

        location = geocode(address_text, addressdetails=True)
        if location is None:
            latitudes.append(None)
            longitudes.append(None)
            matched_addresses.append(None)
            continue

        latitudes.append(location.latitude)
        longitudes.append(location.longitude)
        matched_addresses.append(location.address)

    frame["latitude"] = latitudes
    frame["longitude"] = longitudes
    frame["geocoded_address"] = matched_addresses

    output_path = path.with_name(f"{path.stem}{DEFAULT_SUFFIX}{path.suffix}")
    frame.to_excel(output_path, index=False)
    return output_path


def export_to_shapefile(
    geocoded_excel_path: Path,
    lat_column: str = "latitude",
    lon_column: str = "longitude",
) -> Path:
    """Export geocoded Excel data to a shapefile in the same directory.
    
    Args:
        geocoded_excel_path: Path to the geocoded Excel file.
        lat_column: Name of the latitude column.
        lon_column: Name of the longitude column.
    
    Returns:
        Path to the generated shapefile (.shp).
    
    Raises:
        FileNotFoundError: If the Excel file doesn't exist.
        ValueError: If required columns are missing or no valid coordinates found.
    """
    path = Path(geocoded_excel_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    frame = pd.read_excel(path)
    if frame.empty:
        raise ValueError("The Excel file does not contain any data rows.")

    if lat_column not in frame.columns or lon_column not in frame.columns:
        raise ValueError(
            f"Required columns '{lat_column}' or '{lon_column}' not found in the Excel file."
        )

    # Filter out rows with missing coordinates
    valid_rows = frame.dropna(subset=[lat_column, lon_column])
    if valid_rows.empty:
        raise ValueError("No rows with valid latitude/longitude coordinates found.")

    # Create geometry column from coordinates (lon, lat order for GIS standard)
    geometry = [Point(xy) for xy in zip(valid_rows[lon_column], valid_rows[lat_column])]

    # Create GeoDataFrame with EPSG:4326 (WGS84) coordinate system
    gdf = gpd.GeoDataFrame(valid_rows, geometry=geometry, crs="EPSG:4326")

    # Determine output path (same directory as Excel file, with .shp extension)
    output_shp = path.with_name(f"{path.stem}.shp")

    # Write shapefile
    gdf.to_file(output_shp, driver="ESRI Shapefile")
    return output_shp


def reveal_in_file_manager(target: Path) -> None:
    """Open the folder containing the target path in the system file manager."""
    folder = target if target.is_dir() else target.parent
    try:
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            import subprocess

            subprocess.run(["open", str(folder)], check=False)
        else:
            import subprocess

            subprocess.run(["xdg-open", str(folder)], check=False)
    except Exception as exc:  # pragma: no cover
        print(f"Warning: unable to open folder {folder}: {exc}")


def _cli_progress(current: int, total: int, address: str) -> None:
    """Console progress callback for CLI usage."""
    shortened = (address[:60] + "...") if len(address) > 60 else address
    print(f"[{current}/{total}] {shortened}")


def run_cli(arguments: Optional[list[str]] = None) -> int:
    """Command-line interface entry point."""
    parser = argparse.ArgumentParser(description="Geocode addresses in an Excel file via Nominatim.")
    parser.add_argument("excel_path", help="Path to the Excel workbook containing addresses.")
    parser.add_argument(
        "--address-column",
        "-c",
        dest="address_column",
        help="Name of the Excel column that holds the address strings.",
    )
    parser.add_argument(
        "--user-agent",
        "-u",
        default=DEFAULT_USER_AGENT,
        help="Custom user-agent string to comply with Nominatim usage policy.",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Skip opening the output folder after processing.",
    )

    args = parser.parse_args(arguments)

    try:
        output_path = geocode_excel_file(
            excel_path=Path(args.excel_path),
            address_column=args.address_column,
            user_agent=args.user_agent,
            progress_callback=_cli_progress,
        )
    except (FileNotFoundError, AddressColumnError, ValueError) as err:
        print(f"Error: {err}", file=sys.stderr)
        return 1
    except (GeocoderTimedOut, GeocoderServiceError) as err:
        print(f"Geocoding service error: {err}", file=sys.stderr)
        return 1

    print(f"Geocoded data saved to: {output_path}")
    if not args.no_open:
        reveal_in_file_manager(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
