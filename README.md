# Excel Address Geocoder – GUI Edition

This application lets you geocode rows in an Excel workbook and instantly explore the results on an interactive map. Upload a spreadsheet, choose the address columns that should be processed, and decide which fields appear when you click on a map marker – all from a streamlined dark-themed interface.

---

## 1. System Requirements
- Windows 10 or later
- Python 3.11 (only required if you run from source)
- Internet connection (the app queries the Nominatim service)

---

## 2. Getting the Application

### Option A – Use the packaged executable (recommended)
1. Navigate to `dist/GeocodeGUI/`.
2. Double-click `GeocodeGUI.exe` and wait a moment for the window to appear.

### Option B – Run from source
1. Open PowerShell in the project folder.
2. Activate the local virtual environment:
   ```powershell
   & ".\.venv\Scripts\Activate.ps1"
   ```
3. Launch the GUI:
   ```powershell
   python geocode_gui.py
   ```

---

## 3. First-Time Setup (source users only)
If you didn’t receive a pre-configured virtual environment, install the required libraries:

```powershell
python -m venv .venv
& ".\.venv\Scripts\Activate.ps1"
pip install customtkinter pandas geopy folium openpyxl geopandas shapely
```

---

## 4. Using the GUI

### Step 1 – Load your workbook
- Click **Browse** and choose an Excel file (`.xlsx`, `.xls`, `.xlsm`).
- The app lists all column headers from the first worksheet.

### Step 2 – Choose address information
- **Address column:** Select the field that contains complete address strings. The app attempts to guess a column for you.
- **Optional fields:** If your data splits street, postal code, or city into separate columns, choose them as well. They improve geocoding accuracy but are not required.

### Step 3 – Configure marker popups
- Open the **Map popup columns** dropdown.
- Tick one or more columns whose values you want to see when clicking a map point. You can select multiple entries – the preview text updates to show what will be displayed.

### Step 4 – Start geocoding
- Click **Start geocoding**.
- The status bar shows progress, including the current row being processed.
- When finished, the app creates a workbook suffixed with `_geocoded.xlsx` and a map HTML file.
- A message box summarizes output locations and opens the containing folder.

### Step 5 – Explore the map
- Open the generated `*_map.html` file in any web browser.
- Click markers to view the columns you selected in the multi-select dropdown.

---

## 5. Output Files
- **`*_geocoded.xlsx`** – Copy of your original sheet with these extra columns:
  - `latitude`
  - `longitude`
  - `geocoded_address`
- **`*_map.html`** – Interactive Folium map with clustered markers.
- **`*_geocoded.shp`** – Shapefile with all geocoded points (and associated `.shx`, `.dbf`, `.prj`, `.cpg` files).
  - Ready to import into any GIS software (QGIS, ArcGIS, etc.)
  - Uses WGS84 (EPSG:4326) coordinate reference system
  - Contains all columns from the original Excel file plus coordinates

---

## 6. Tips for Best Results
- Provide as much context as possible (street + postal code + city).
- Remove obvious duplicates or invalid addresses before processing.
- Keep geocoding batches modest to respect Nominatim’s usage policy. The app waits between requests, but heavy usage can still be throttled.
- If the map shows fewer points than expected, inspect the `_geocoded.xlsx` file for missing coordinates.

---

## 7. Troubleshooting
| Issue | What to check |
| --- | --- |
| “No columns found” | Ensure the first worksheet has a header row. |
| Blank latitude/longitude | The address could not be resolved – verify spelling or add postal codes. |
| GUI doesn’t launch | Install the latest Visual C++ redistributables and make sure `.NET Desktop Runtime` is available (PyInstaller builds sometimes require them). |
| “Map generation failed” message | Confirm that `latitude` and `longitude` columns exist in the output workbook. |

---

## 8. Advanced Usage (optional)
- Run `python geocode_addresses.py --help` for the command-line interface used by the GUI.
- Generate the executable yourself with:
  ```powershell
  & ".\.venv\Scripts\python.exe" -m PyInstaller GeocodeGUI.spec
  ```

---

## 9. Support & Feedback
Encounter a bug or have a feature suggestion? Share:
- The Excel columns you selected (screenshots help).
- The exact error message.
- The generated `_geocoded.xlsx` and `_map.html` if possible.

Thanks for using the Excel Address Geocoder GUI!
