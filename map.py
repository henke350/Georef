"""Create an interactive HTML map from a geocoded Excel workbook."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Sequence

import folium
import pandas as pd
from folium.plugins import MarkerCluster

DEFAULT_LAT_COLUMN = "latitude"
DEFAULT_LON_COLUMN = "longitude"
DEFAULT_HTML_SUFFIX = "_map.html"
DEFAULT_POPUP_CANDIDATES: Sequence[str] = ("geocoded_address", "address")


def _select_popup_columns(frame: pd.DataFrame, requested: Iterable[str] | None) -> list[str]:
    if requested:
        columns: list[str] = []
        for name in requested:
            if name in frame.columns:
                columns.append(name)
        if columns:
            return columns
    for candidate in DEFAULT_POPUP_CANDIDATES:
        if candidate in frame.columns:
            return [candidate]
    return []


def _build_popup(row: pd.Series, columns: Sequence[str]) -> str | None:
    if not columns:
        return None
    lines: list[str] = []
    for column in columns:
        value = row.get(column, "")
        if pd.isna(value) or value == "":
            continue
        lines.append(f"{column}: {value}")
    if not lines:
        return None
    return "<br>".join(lines)


def generate_map(
    excel_path: Path | str,
    *,
    lat_column: str = DEFAULT_LAT_COLUMN,
    lon_column: str = DEFAULT_LON_COLUMN,
    popup_columns: Iterable[str] | None = None,
    output_html: Path | str | None = None,
    enable_clustering: bool = True,
) -> Path:
    path = Path(excel_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    frame = pd.read_excel(path)
    if lat_column not in frame.columns or lon_column not in frame.columns:
        raise ValueError(
            "The Excel file must contain latitude and longitude columns. "
            f"Missing: {lat_column!r} or {lon_column!r}."
        )

    usable = frame.dropna(subset=[lat_column, lon_column])
    if usable.empty:
        raise ValueError("No rows contain both latitude and longitude values.")

    center_lat = float(usable[lat_column].mean())
    center_lon = float(usable[lon_column].mean())

    fmap = folium.Map(location=[center_lat, center_lon], zoom_start=12, control_scale=True)

    marker_parent: folium.Map | MarkerCluster
    if enable_clustering and len(usable) > 1:
        marker_parent = MarkerCluster().add_to(fmap)
    else:
        marker_parent = fmap

    popup_fields = _select_popup_columns(frame, popup_columns)

    for _, row in usable.iterrows():
        lat = float(row[lat_column])
        lon = float(row[lon_column])
        popup_html = _build_popup(row, popup_fields)
        marker = folium.Marker(location=[lat, lon])
        if popup_html:
            marker.add_child(folium.Popup(popup_html, max_width=300))
        marker.add_to(marker_parent)

    if output_html is None:
        output_html = path.with_name(f"{path.stem}{DEFAULT_HTML_SUFFIX}")
    else:
        output_html = Path(output_html)
        if output_html.is_dir():
            output_html = output_html / f"{path.stem}{DEFAULT_HTML_SUFFIX}"

    fmap.save(str(output_html))
    return Path(output_html)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Plot geocoded Excel data onto an interactive map.")
    parser.add_argument("excel_path", help="Path to the geocoded Excel workbook.")
    parser.add_argument(
        "--lat-column",
        default=DEFAULT_LAT_COLUMN,
        help="Name of the latitude column (default: latitude).",
    )
    parser.add_argument(
        "--lon-column",
        default=DEFAULT_LON_COLUMN,
        help="Name of the longitude column (default: longitude).",
    )
    parser.add_argument(
        "--popup-column",
        action="append",
        dest="popup_columns",
        help="Column to include in marker popups. Use multiple times for more columns.",
    )
    parser.add_argument(
        "--output-html",
        help="Optional path for the output HTML file. Default uses the Excel name with _map.html.",
    )
    parser.add_argument(
        "--no-cluster",
        action="store_true",
        help="Disable marker clustering even when multiple points exist.",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    output = generate_map(
        args.excel_path,
        lat_column=args.lat_column,
        lon_column=args.lon_column,
        popup_columns=args.popup_columns,
        output_html=args.output_html,
        enable_clustering=not args.no_cluster,
    )
    print(f"Map saved to: {output}")


if __name__ == "__main__":
    main()
