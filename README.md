# Excel Address Geocoder

A desktop application for geocoding addresses from Excel files and visualizing them on interactive maps. Features a modern dark-themed GUI built with CustomTkinter.

---

## Features

- 📍 **Batch Geocoding** – Process entire Excel spreadsheets with address data
- 🗺️ **Interactive Maps** – Generate HTML maps with clustered markers using Folium
- 🎨 **Color-Coded Markers** – Optionally color markers by any column value with auto-generated legend
- 📋 **Customizable Popups** – Select which columns to display when clicking markers
- 📦 **Shapefile Export** – Export geocoded data to GIS-ready shapefiles (.shp)
- ⏹️ **Cancellable Operations** – Stop long-running geocoding at any time
- 🔍 **Smart Column Detection** – Automatically detects address columns

---

## System Requirements

- Windows 10 or later
- Python 3.11+ (only required if running from source)
- Internet connection (uses OpenStreetMap's Nominatim geocoding service)

---

## Installation

### Option A – Use the Packaged Executable (Recommended)

1. Navigate to `dist/GeocodeGUI/`
2. Double-click `GeocodeGUI.exe`

### Option B – Run from Source

1. Clone or download this repository
2. Open PowerShell in the project folder
3. Create and activate a virtual environment:
   ```powershell
   python -m venv .venv
   & ".\.venv\Scripts\Activate.ps1"
   ```
4. Install dependencies:
   ```powershell
   pip install customtkinter pandas geopy folium openpyxl geopandas shapely
   ```
5. Launch the application:
   ```powershell
   python geocode_gui.py
   ```

---

## Usage

### Step 1 – Load Your Workbook

- Click **Browse** and select an Excel file (`.xlsx`, `.xls`, `.xlsm`)
- The app reads all column headers from the first worksheet

### Step 2 – Select Address Column

- Choose the column containing address strings from the **Address column** dropdown
- The app auto-detects columns with "address" in the name

### Step 3 – Configure Map Display (Optional)

- **Map popup columns:** Select one or more columns to display when clicking map markers
- **Color markers by column:** Choose a column to color-code markers (e.g., by category, status, or region). A legend will be automatically added to the map.

### Step 4 – Start Geocoding

- Click **Start geocoding** to begin processing
- Progress bar and status text show current progress
- Click **Stop** to cancel at any time (partial results are preserved)

### Step 5 – View Results

- Output files are created in the same folder as your input file
- A dialog shows the output location and opens the containing folder
- Open the `*_map.html` file in any web browser to explore your data

---

## Output Files

| File | Description |
|------|-------------|
| `*_geocoded.xlsx` | Copy of your original data with added `latitude`, `longitude`, and `geocoded_address` columns |
| `*_map.html` | Interactive Folium map with markers (clustered unless color-coding is used) |
| `*_geocoded.shp` | ESRI Shapefile with geocoded points (includes `.shx`, `.dbf`, `.prj`, `.cpg` sidecar files) |

**Shapefile notes:**
- Uses WGS84 (EPSG:4326) coordinate reference system
- Compatible with QGIS, ArcGIS, and other GIS software
- Contains all columns from the original Excel file

---

## Tips for Best Results

- Provide complete addresses (street + city + postal code) for higher accuracy
- Remove duplicates or obviously invalid addresses before processing
- Keep batch sizes reasonable – Nominatim has usage limits and the app waits between requests
- If markers are missing, check the `_geocoded.xlsx` file for blank coordinates

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No columns found" | Ensure the first worksheet has a header row |
| Blank latitude/longitude | Address could not be resolved – verify spelling or add postal codes |
| GUI doesn't launch | Install Visual C++ Redistributables and .NET Desktop Runtime |
| Map generation failed | Confirm `latitude` and `longitude` columns exist in the geocoded file |

---

## Dependencies

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) – Modern GUI framework
- [pandas](https://pandas.pydata.org/) – Data manipulation
- [geopy](https://geopy.readthedocs.io/) – Geocoding via Nominatim
- [folium](https://python-visualization.github.io/folium/) – Interactive maps
- [geopandas](https://geopandas.org/) – Geospatial data handling
- [openpyxl](https://openpyxl.readthedocs.io/) – Excel file support
- [shapely](https://shapely.readthedocs.io/) – Geometric operations
