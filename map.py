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

# Color palette for categorical coloring (up to 20 distinct colors)
MARKER_COLORS: Sequence[str] = (
    "blue", "red", "green", "purple", "orange", "darkred", "lightred", "beige",
    "darkblue", "darkgreen", "cadetblue", "darkpurple", "pink", "lightblue",
    "lightgreen", "gray", "black", "lightgray", "white", "yellow"
)


def _build_color_map(series: pd.Series) -> dict:
    """Create a mapping from unique values to colors."""
    unique_values = series.dropna().unique()
    color_map = {}
    for i, value in enumerate(unique_values):
        color_map[value] = MARKER_COLORS[i % len(MARKER_COLORS)]
    return color_map


def _build_legend_html(color_map: dict, column_name: str) -> str:
    """Build HTML for a legend showing color mappings."""
    legend_items = []
    for value, color in color_map.items():
        # Use display-friendly color names and escape HTML
        display_value = str(value) if pd.notna(value) else "N/A"
        legend_items.append(
            f'<li><span style="background:{color};width:12px;height:12px;'
            f'display:inline-block;margin-right:6px;border-radius:50%;"></span>{display_value}</li>'
        )
    
    legend_html = f'''
    <div style="
        position: fixed;
        bottom: 50px;
        left: 50px;
        z-index: 1000;
        background-color: white;
        padding: 10px;
        border: 2px solid grey;
        border-radius: 5px;
        font-size: 12px;
        max-height: 300px;
        overflow-y: auto;
    ">
        <strong>{column_name}</strong>
        <ul style="list-style: none; padding: 0; margin: 5px 0 0 0;">
            {"".join(legend_items)}
        </ul>
    </div>
    '''
    return legend_html


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
    color_column: str | None = None,
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

    # Build color map if color column is specified
    color_map: dict | None = None
    if color_column and color_column in frame.columns:
        color_map = _build_color_map(usable[color_column])

    # When using color coding, disable clustering to show individual marker colors
    use_clustering = enable_clustering and len(usable) > 1 and color_map is None
    
    marker_parent: folium.Map | MarkerCluster
    if use_clustering:
        marker_parent = MarkerCluster().add_to(fmap)
    else:
        marker_parent = fmap

    popup_fields = _select_popup_columns(frame, popup_columns)

    for _, row in usable.iterrows():
        lat = float(row[lat_column])
        lon = float(row[lon_column])
        popup_html = _build_popup(row, popup_fields)
        
        # Determine marker color
        marker_color = "blue"  # default
        if color_map and color_column:
            value = row.get(color_column)
            if pd.notna(value) and value in color_map:
                marker_color = color_map[value]
            else:
                marker_color = "gray"  # for NaN or unknown values
        
        marker = folium.Marker(
            location=[lat, lon],
            icon=folium.Icon(color=marker_color)
        )
        if popup_html:
            marker.add_child(folium.Popup(popup_html, max_width=300))
        marker.add_to(marker_parent)

    # Add legend if color coding is used
    if color_map and color_column:
        legend_html = _build_legend_html(color_map, color_column)
        fmap.get_root().html.add_child(folium.Element(legend_html))

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
    parser.add_argument(
        "--color-column",
        help="Column to use for coloring markers. Each unique value gets a different color.",
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
        color_column=args.color_column,
    )
    print(f"Map saved to: {output}")


if __name__ == "__main__":
    main()
