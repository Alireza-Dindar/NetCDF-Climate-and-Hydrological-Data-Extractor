# GeoNetX - NetCDF Extractor with Map Preview for Climate and Hydrological Data

Download:
https://github.com/Alireza-Dindar/NetCDF-Climate-and-Hydrological-Data-Extractor/releases/download/Beta/GeoNetXSetup_Beta.exe


<img width="443" height="358" alt="historical avg" src="https://github.com/user-attachments/assets/9a588bd8-623c-428b-9eb3-66d95ff59cc5" />

<img width="443" height="358" alt="Screenshot 2026-01-02 211034" src="https://github.com/user-attachments/assets/3969d904-36c9-4fdb-9f7d-b8f7654dbce3" />

<img width="750" height="437" alt="image" src="https://github.com/user-attachments/assets/877515d9-0603-47db-8eae-3c3159e7b6d6" />


**New updates and fixes:**
- Improved UI
- Added aggregation options for extraction
- New feature which enables user to clip the maps and data using Shapfile(.shp) boundaries 
- Fixed issues with metric coordinate projections (e.g. UTM, Lambert, ...)
- Added multiple GeoTiff series exporting option
- Added scale factor for enhance shapefile clipping in visualization part
- Now app supports Persian (Farsi) language
- App is tested for NASA Daymet, GLEAM, ERA5, CHIRPS, GCM and RCM datasets. 

----------------------------------------------------------------------------------------------------------------------
# A desktop application built with Python (PyQt5) to visualize and extract data from NetCDF (.nc) files.

## Features
- 🌍 **Interactive Grid Preview Map:** Preview detected nodes on an interactive Folium map before extraction. Large node sets are handled with fast clustering and browser fallback for stability.
- 🗺️ **Flexible Spatial Filtering:** Supports manual `lat/lon` bounding boxes and optional shapefile-based filtering/clipping (WGS84-aware).
- 🧭 **Robust CRS Handling:** Detects projected datasets (e.g., Lambert Conformal / Daymet), applies CRS fallback when needed, and reprojects to geographic coordinates for consistent outputs.
- 📅 **Temporal Detection and Slicing:** Automatically detects available date ranges (for multi-file cases, constrained by common overlap) and supports start/end subsetting.
- 🎞️ **Advanced Visualization and Export:** Includes map visualization, GeoTIFF export, GeoTIFF series export, and time-series animation with optional smoothing and configurable grid scale.
- 📊 **Extraction Outputs:** Exports node lists to CSV and extracted climate variables to Excel/CSV, with configurable frequency/statistics and skip-existing behavior.
- 🚀 **Batch Processing:** Processes multiple folders/files in one run with live progress/log reporting and log export for debugging.

## How to Run 

1. Install and launch `GeoNetX-beta.exe` (if Windows blocks permissions, run once as Administrator).
2. Click `Browse Data Root` and select the parent folder containing your NetCDF subfolders.
3. Select the subfolder(s) that contain `.nc` files.
4. Define your area of interest:
   - Enter `Lat/Lon` bounds manually, or
   - Load a shapefile (`.shp`) and enable shapefile filtering.
5. (Optional) Click `Detect Spatial Extent` to auto-fill coordinate bounds from selected data/shapefile.
6. Click `Detect Temporal Extent` (or `Detect available date range`) to read available dates.
7. Set `Start Date` and `End Date`.
8. (Optional) Click `Load Grids` to preview available nodes, then `Show Nodes on Map`.
9. In map preview:
   - For small/medium grids, nodes are shown directly.
   - For very large grids, the app may switch to a faster clustered view and open in system browser for stability.
10. Select variables/bands to extract (`Select Columns` / band filter).
11. Set output options:
   - Target folder
   - Output frequency/statistic (if needed)
   - Skip existing files (optional)
12. Click `Start Extraction`.
13. (Optional) Use `Visualize Data` for:
   - `Visualize Map`
   - `Save GeoTIFF`
   - `Save GeoTIFF Series`
   - `Animate Time Series`
   - `Approx grid scale (deg, 0=auto)` for coarse datasets.
14. (Optional) Click `Export Grid List to CSV` to save node coordinates within current bounds.
15. Check the `Log` panel for progress/errors; use `Export Log` for support/debugging.
 
---------------------------------------------------------------
!!! Feel free to give feedbacks on GitHub or via email: alireza.dindar.1998@gmail.com !!!
