"""
netcdf_extractor_with_map.py

PyQt5 desktop app to preview and extract NetCDF grid nodes inside a bounding box,
and to visualize node locations on an interactive world map (folium) embedded via QWebEngineView.

Features:
- Browse data root and select subfolders
- Preview available lat/lon nodes inside bounding box using the first .nc file in a folder
- Export grid list to CSV
- Run extraction to write .xlsx files (same logic as before) in background thread
"""



import os
import time
import tracemalloc
import gc
import tempfile
import webbrowser
from pathlib import Path
import cftime
import numpy as np
import pandas as pd
import xarray as xr
from PyQt5 import QtWidgets, QtCore, QtGui
import folium
from folium.plugins import MarkerCluster, HeatMap, MousePosition
import datetime
try:
    import cftime
except Exception:
    cftime = None
import platform
import sys
import traceback





# temporary diagnostic block



print(f"Library versions: PyQt5 {QtCore.QT_VERSION_STR}, xarray {xr.__version__}, pandas {pd.__version__}, numpy {np.__version__}, cftime {getattr(cftime, '__version__', 'N/A')}, folium {folium.__version__}")
app = QtWidgets.QApplication(sys.argv)

# Use native Windows style when available for native controls
if platform.system() == "Windows":
    app.setStyle("windowsvista")  # "windowsvista" or "Windows" depending on Qt version

qss_text = """
/* --- MAIN WINDOW --- */
QMainWindow, QWidget {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #f6f9fc, stop:0.5 #e9f0f7, stop:1 #dfe9f4);
  color: #222;
  font-family: Segoe UI, Arial, sans-serif;
  font-size: 10pt;
}

/* --- GROUP BOXES --- */
QGroupBox {
  border: 1px solid #bfcfe6;
  border-radius: 6px;
  margin-top: 20px;
  background: rgba(255,255,255,0.4);
}
QGroupBox::title {
  subcontrol-origin: margin;
  subcontrol-position: top left;
  left: 8px;
  padding: 0 4px;
  color: #1f4f7a;
  font-weight: 600;
}

/* --- BUTTONS (Glossy) --- */
QPushButton {
  border: 1px solid #7aa7d9;
  border-radius: 4px;
  padding: 5px 15px;
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #eaf6ff, stop:0.5 #d6ecff, stop:1 #c7e6ff);
  color: #0b3b5a;
  font-weight: bold;
}
QPushButton:hover {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #ffffff, stop:0.5 #eaf6ff, stop:1 #d6ecff);
  border-color: #5d95d1;
}
QPushButton:pressed {
  background: #b7dbff;
  border-color: #5d95d1;
}
QPushButton:disabled {
  color: #9aaec0;
  background: #e9eef4;
  border: 1px solid #d0dbe6;
}

/* --- INPUTS (Clean & Simple) --- 
   We style the border/color, but let Windows draw the arrows so they always work.
*/
QLineEdit, QAbstractSpinBox, QDateEdit, QComboBox {
  border: 1px solid #bfcfe6;
  border-radius: 4px;
  padding: 4px;
  background: #ffffff;
  selection-background-color: #cfe9ff;
  min-height: 22px; 
}

/* --- LISTS & SCROLLBARS --- */
QListWidget, QTreeWidget, QTableWidget {
  border: 1px solid #bfcfe6;
  border-radius: 4px;
  background: #ffffff;
}
QHeaderView::section {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f0f6fb, stop:1 #dfeefb);
  padding: 5px;
  border: 1px solid #c7dff0;
  color: #1f4f7a;
  font-weight: bold;
}

QScrollBar:vertical {
    border: none; background: #e9eef4; width: 12px; border-radius: 6px; margin: 0px;
}
QScrollBar::handle:vertical {
    background: #a3c1da; min-height: 20px; border-radius: 6px; margin: 2px;
}
QScrollBar::handle:vertical:hover { background: #7aa7d9; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }

QScrollBar:horizontal {
    border: none; background: #e9eef4; height: 12px; border-radius: 6px; margin: 0px;
}
QScrollBar::handle:horizontal {
    background: #a3c1da; min-width: 20px; border-radius: 6px; margin: 2px;
}
QScrollBar::handle:horizontal:hover { background: #7aa7d9; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }

/* --- PROGRESS BAR --- */
QProgressBar {
  border: 1px solid #7aa7d9;
  border-radius: 6px;
  text-align: center;
  background: #eaf6ff;
  color: #0b3b5a;
  font-weight: bold;
}
QProgressBar::chunk {
  background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6fb3e6, stop:1 #2f8fcf);
  border-radius: 5px;
}
"""

with open("glossy.qss", "w", encoding="utf-8") as f:
    f.write(qss_text)
with open("glossy.qss", "r") as f:
    app.setStyleSheet(f.read())





try:
    from PyQt5 import QtWebEngineWidgets
    WEB_ENGINE_AVAILABLE = True
except Exception:
    WEB_ENGINE_AVAILABLE = False
    
def _to_cftime_like_sample(ts, sample):
    """
    Convert ts (datetime / pandas.Timestamp / datetime.date) to a cftime instance
    of the same class as `sample` (a dataset time sample). Returns None on failure.
    """
    if ts is None or sample is None or cftime is None:
        return None

    # normalize ts to components
    if isinstance(ts, datetime.date) and not isinstance(ts, datetime.datetime):
        y, m, d = ts.year, ts.month, ts.day
        hh = mm = ss = 0
    elif isinstance(ts, datetime.datetime):
        y, m, d = ts.year, ts.month, ts.day
        hh, mm, ss = ts.hour, ts.minute, ts.second
    else:
        try:
            if isinstance(ts, pd.Timestamp):
                dt = ts.to_pydatetime()
                y, m, d = dt.year, dt.month, dt.day
                hh, mm, ss = dt.hour, dt.minute, dt.second
            else:
                y = int(getattr(ts, "year")); m = int(getattr(ts, "month")); d = int(getattr(ts, "day"))
                hh = int(getattr(ts, "hour", 0)); mm = int(getattr(ts, "minute", 0)); ss = int(getattr(ts, "second", 0))
        except Exception:
            return None

    cls = type(sample)
    try:
        if cls is cftime.DatetimeNoLeap:
            return cftime.DatetimeNoLeap(y, m, d, hh, mm, ss)
        if cls is cftime.Datetime360Day:
            return cftime.Datetime360Day(y, m, d, hh, mm, ss)
        if cls is cftime.DatetimeGregorian:
            return cftime.DatetimeGregorian(y, m, d, hh, mm, ss)
        if cls is cftime.DatetimeProlepticGregorian:
            return cftime.DatetimeProlepticGregorian(y, m, d, hh, mm, ss)
        if cls is cftime.DatetimeJulian:
            return cftime.DatetimeJulian(y, m, d, hh, mm, ss)
        # fallback: attempt constructor
        return cls(y, m, d, hh, mm, ss)
    except Exception:
        return None

def _detect_band_coord(self, ds):
    """
    Return (coord_name, values_list) for a band-like coordinate, or (None, None).
    Excludes bounds-like names (bnds, *_bnds) and requires >2 values.
    """
    candidates = ["band", "bands", "band_index", "bnd", "bnds", "layer", "level"]
    # 1) coords
    for cand in candidates:
        if cand in ds.coords:
            vals = np.asarray(ds[cand].values).ravel()
            if vals.size > 2 and cand.lower() not in ("bnds", "time_bnds", "lat_bnds", "lon_bnds"):
                return cand, list(vals)
    # 2) dims with a coordinate variable
    for cand in candidates:
        if cand in ds.dims:
            size = ds.sizes.get(cand, 0)
            if size > 2 and cand.lower() not in ("bnds", "time_bnds", "lat_bnds", "lon_bnds"):
                # try to get coord values if present
                vals = ds[cand].values if cand in ds.coords else np.arange(size)
                return cand, list(np.asarray(vals).ravel())
    # 3) heuristic: any coord/dim containing 'band' or 'layer' with size>2
    for name in list(ds.coords) + list(ds.dims):
        lname = name.lower()
        if ("band" in lname or "layer" in lname or "level" in lname) and lname not in ("bnds", "time_bnds", "lat_bnds", "lon_bnds"):
            vals = np.asarray(ds[name].values) if name in ds.coords else np.arange(ds.sizes.get(name, 0))
            if getattr(vals, "size", len(vals)) > 2:
                return name, list(np.asarray(vals).ravel())
    return None, None

def timestamp_to_cftime(ts, calendar):
    """
    Convert a pandas.Timestamp / datetime to a cftime object matching `calendar`.
    Returns None on failure or if cftime is not available.
    """
    if ts is None:
        return None
    # normalize to datetime components
    if isinstance(ts, datetime.date) and not isinstance(ts, datetime.datetime):
        year, month, day = ts.year, ts.month, ts.day
        hour = minute = second = 0
    elif isinstance(ts, datetime.datetime):
        year, month, day = ts.year, ts.month, ts.day
        hour, minute, second = ts.hour, ts.minute, ts.second
    else:
        # try pandas Timestamp
        try:
            if isinstance(ts, pd.Timestamp):
                dt = ts.to_pydatetime()
                year, month, day = dt.year, dt.month, dt.day
                hour, minute, second = dt.hour, dt.minute, dt.second
            else:
                # fallback: try attributes
                year = int(getattr(ts, "year"))
                month = int(getattr(ts, "month"))
                day = int(getattr(ts, "day"))
                hour = int(getattr(ts, "hour", 0))
                minute = int(getattr(ts, "minute", 0))
                second = int(getattr(ts, "second", 0))
        except Exception:
            return None

    if cftime is None:
        return None

    cal = (calendar or "").lower()
    try:
        if "noleap" in cal or "365_day" in cal or "no_leap" in cal:
            return cftime.DatetimeNoLeap(year, month, day, hour, minute, second)
        if "360_day" in cal or "360" in cal:
            return cftime.Datetime360Day(year, month, day, hour, minute, second)
        if "proleptic_gregorian" in cal:
            return cftime.DatetimeProlepticGregorian(year, month, day, hour, minute, second)
        if "gregorian" in cal or "standard" in cal or cal == "" or "greg" in cal:
            return cftime.DatetimeGregorian(year, month, day, hour, minute, second)
        if "julian" in cal:
            return cftime.DatetimeJulian(year, month, day, hour, minute, second)
        # fallback
        return cftime.DatetimeGregorian(year, month, day, hour, minute, second)
    except Exception:
        return None

class ColumnSelectionDialog(QtWidgets.QDialog):
    def __init__(self, available_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Output Columns")
        self.resize(300, 400)
        self.selected_columns = []
        
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("Uncheck variables you don't need:"))

        self.checkboxes = []
        
        # Scroll area in case there are many variables
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QtWidgets.QWidget()
        self.vbox = QtWidgets.QVBoxLayout(content_widget)
        
        for col in available_columns:
            chk = QtWidgets.QCheckBox(col)
            chk.setChecked(True) # Default to all selected
            self.vbox.addWidget(chk)
            self.checkboxes.append(chk)
            
        self.vbox.addStretch()
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        # Buttons
        btn_box = QtWidgets.QHBoxLayout()
        ok_btn = QtWidgets.QPushButton("OK")
        ok_btn.clicked.connect(self.accept_selection)
        cancel_btn = QtWidgets.QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        btn_box.addWidget(ok_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box)

    def accept_selection(self):
        self.selected_columns = [chk.text() for chk in self.checkboxes if chk.isChecked()]
        self.accept()


# Worker signals
class WorkerSignals(QtCore.QObject):
    progress = QtCore.pyqtSignal(int)
    log = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal()
    busy = QtCore.pyqtSignal(bool)



class ExtractWorker(QtCore.QRunnable):
    def __init__(self, params):
        super().__init__()
        self.params = params
        self.signals = WorkerSignals()

    @QtCore.pyqtSlot()
    def run(self):
        p = self.params
        # We now expect a list of full file paths, not folders
        files_to_process = p["files"] 
        
        date_start = p.get("date_start", None)
        date_end = p.get("date_end", None)
        lat_min, lat_max = p["lat_min"], p["lat_max"]
        lon_min, lon_max = p["lon_min"], p["lon_max"]
        band_param = p.get("band", None)
        band_coord_name = p.get("band_coord_name", None)
        chunk_lat, chunk_lon = p["chunk_lat"], p["chunk_lon"]
        skip_existing = p["skip_existing"]
        save_root = Path(p["save_root"])
        user_cols = p.get("columns")

        total = len(files_to_process)
        if total == 0:
            self.signals.log.emit("No files selected to process.")
            self.signals.finished.emit()
            return

        processed = 0
        for file_path_str in files_to_process:
            processed += 1
            file_path = Path(file_path_str)
            
            # Determine folder name for output structure
            # Assuming structure: Root/Folder/File.nc
            # We want output: SaveRoot/Folder/File.xlsx
            folder_name = file_path.parent.name
            
            try:
                save_path = save_root / folder_name
                save_path.mkdir(parents=True, exist_ok=True)
                
                # Logic to determine extension (csv vs xlsx) happens later, start with xlsx
                out_name = save_path / f"{file_path.stem}.xlsx"
                
                # Check for existing (checking both xlsx and csv versions)
                if skip_existing:
                    if out_name.exists() or out_name.with_suffix('.csv').exists():
                        self.signals.log.emit(f"Skipping existing: {file_path.stem}")
                        self.signals.progress.emit(int(processed / total * 100))
                        continue

                self.signals.log.emit(f"Opening: {file_path.name}")
                self.signals.busy.emit(True)
                tracemalloc.start()
                start_time = time.time()

                ds = xr.open_dataset(file_path.as_posix(), chunks={"lat": int(chunk_lat), "lon": int(chunk_lon)})

                # --- FIX: Time Slicing Logic ---
                if date_start is not None or date_end is not None:
                    # 1. Detect the time coordinate name (e.g., 'time', 'valid_time')
                    time_name = None
                    for cand in ["time", "valid_time", "Time", "times", "date"]:
                        if cand in ds.coords:
                            time_name = cand
                            break
                    
                    if time_name:
                        try:
                            # 2. Check if data uses cftime (non-standard calendar)
                            times_vals = ds[time_name].values
                            sample = times_vals.ravel()[0] if times_vals.size > 0 else None
                            is_cftime = False
                            if cftime is not None and sample is not None:
                                is_cftime = isinstance(sample, (
                                    cftime.datetime, cftime.DatetimeNoLeap, cftime.Datetime360Day,
                                    cftime.DatetimeGregorian, cftime.DatetimeJulian, cftime.DatetimeProlepticGregorian
                                ))

                            # 3. Convert user dates to match the dataset format
                            start_val = date_start
                            end_val = date_end

                            if is_cftime:
                                # Convert pandas Timestamp -> cftime
                                start_val = _to_cftime_like_sample(date_start, sample)
                                end_val = _to_cftime_like_sample(date_end, sample)
                            
                            # 4. Apply the slice
                            if start_val is not None and end_val is not None:
                                ds = ds.sel({time_name: slice(start_val, end_val)})
                                self.signals.log.emit(f"Sliced time: {start_val} to {end_val}")
                            else:
                                self.signals.log.emit(f"Warning: Could not convert dates for {file_path.name}. Skipping time slice.")
                        
                        except Exception as e:
                            self.signals.log.emit(f"Failed to slice time for {file_path.name}: {e}")
               

                self.signals.busy.emit(False)

                # ... [Keep your existing Band Filtering logic here] ...
                # ... (Copy the band filtering block from your previous code) ...

                # --- Detect coordinate names dynamically (FROM PREVIOUS FIX) ---
                lat_name = None
                lon_name = None
                for cand in ["lat", "latitude", "y"]:
                    if cand in ds.coords:
                        lat_name = cand
                        break
                for cand in ["lon", "longitude", "x"]:
                    if cand in ds.coords:
                        lon_name = cand
                        break

                if not lat_name or not lon_name:
                    self.signals.log.emit(f"Skipping {file_path.name}: Could not find lat/lon coordinates.")
                    self.signals.progress.emit(int(processed / total * 100))
                    continue

                # Filter spatial
                filtered = ds.sortby([lat_name, lon_name]).where(
                    (ds[lat_name] >= float(lat_min)) & (ds[lat_name] <= float(lat_max)) &
                    (ds[lon_name] >= float(lon_min)) & (ds[lon_name] <= float(lon_max)),
                    drop=True
                ).unify_chunks()

                if filtered[lat_name].size < 2 or filtered[lon_name].size < 2:
                    self.signals.log.emit(f"Insufficient nodes in {file_path.name}; skipping.")
                    tracemalloc.stop()
                    self.signals.progress.emit(int(processed / total * 100))
                    continue

                self.signals.log.emit("DEBUG: Starting conversion to DataFrame...")
                df = filtered.load().to_dataframe().reset_index()
                
                # Normalize column names
                df.rename(columns={lat_name: "lat", lon_name: "lon"}, inplace=True)

                # Filter Columns (User Selection)
                if user_cols:
                    final_cols = [c for c in user_cols if c in df.columns]
                    if final_cols:
                        df = df[final_cols]

                # Fix cftime objects
                for col in df.columns:
                    if df[col].dtype == object:
                        try:
                            first_valid = df[col].dropna().iloc[0]
                            if "cftime" in str(type(first_valid)):
                                df[col] = df[col].astype(str)
                        except IndexError:
                            pass

                # Save (Check limit)
                if len(df) > 1000000:
                    self.signals.log.emit("Data exceeds Excel limit. Switching to CSV.")
                    out_name = out_name.with_suffix('.csv')
                    df.to_csv(out_name.as_posix(), index=False)
                else:
                    df.to_excel(out_name.as_posix(), index=False)
                
                gc.collect()
                elapsed = time.time() - start_time
                _, peak = tracemalloc.get_traced_memory()
                tracemalloc.stop()
                self.signals.log.emit(f"Saved {out_name.name} — time: {elapsed:.1f}s")

            except Exception as e:
                self.signals.busy.emit(False)
                full_error = traceback.format_exc()
                print(full_error)
                self.signals.log.emit(f"Error processing {file_path.name}:\n{full_error}")
            
            self.signals.progress.emit(int(processed / total * 100))

        self.signals.log.emit("All done.")
        self.signals.finished.emit()

                


class MainWindow(QtWidgets.QWidget):
    def _to_date(self, t):
            """Return a datetime.date or None for input t (handles date, datetime, pandas.Timestamp)."""
            if t is None:
                return None
            # already a date (but not datetime)
            if isinstance(t, datetime.date) and not isinstance(t, datetime.datetime):
                return t
            # datetime -> date
            if isinstance(t, datetime.datetime):
                return t.date()
            # pandas Timestamp
            try:
                if isinstance(t, pd.Timestamp):
                    return t.to_pydatetime().date()
            except Exception:
                pass
            # fallback: try attributes
            if hasattr(t, "year") and hasattr(t, "month") and hasattr(t, "day"):
                return datetime.date(int(t.year), int(t.month), int(t.day))
            return None
    
    def populate_folders(self, root):
        """Populates the tree with folders and .nc files."""
        self.file_tree.clear()
        p = Path(root)
        if not p.exists():
            return

        # Iterate over subdirectories
        for subdir in sorted([x for x in p.iterdir() if x.is_dir()]):
            # Create Parent Item (Folder)
            folder_item = QtWidgets.QTreeWidgetItem([subdir.name])
            folder_item.setFlags(folder_item.flags() | QtCore.Qt.ItemIsTristate | QtCore.Qt.ItemIsUserCheckable)
            folder_item.setCheckState(0, QtCore.Qt.Unchecked)
            
            # Find .nc files inside
            nc_files = sorted(list(subdir.glob("*.nc")))
            
            if nc_files:
                self.file_tree.addTopLevelItem(folder_item)
                for f in nc_files:
                    # Create Child Item (File)
                    file_item = QtWidgets.QTreeWidgetItem([f.name])
                    file_item.setFlags(file_item.flags() | QtCore.Qt.ItemIsUserCheckable)
                    file_item.setCheckState(0, QtCore.Qt.Unchecked)
                    # Store full path in UserRole for easy retrieval later
                    file_item.setData(0, QtCore.Qt.UserRole, str(f.absolute()))
                    folder_item.addChild(file_item)

        self.file_tree.expandAll()

    def toggle_tree_selection(self, check):
        """Selects or Deselects all items in the tree."""
        state = QtCore.Qt.Checked if check else QtCore.Qt.Unchecked
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            item.setCheckState(0, state)
            # Children are auto-handled if ItemIsTristate is set on parent, 
            # but sometimes explicit loop is safer:
            for j in range(item.childCount()):
                item.child(j).setCheckState(0, state)

    def get_selected_files(self):
        """Returns a list of absolute paths (str) of checked files."""
        selected_files = []
        for i in range(self.file_tree.topLevelItemCount()):
            folder_item = self.file_tree.topLevelItem(i)
            for j in range(folder_item.childCount()):
                file_item = folder_item.child(j)
                if file_item.checkState(0) == QtCore.Qt.Checked:
                    path = file_item.data(0, QtCore.Qt.UserRole)
                    if path:
                        selected_files.append(path)
        return selected_files
    
    
    def open_column_selector(self):
        if not self._detected_columns:
            self.append_log("No columns detected. Run preview first.")
            return
            
        dlg = ColumnSelectionDialog(self._detected_columns, self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._user_selected_cols = dlg.selected_columns
            if self._user_selected_cols:
                self.append_log(f"Extraction will include: {', '.join(self._user_selected_cols)}")
            else:
                self.append_log("No columns selected! Extraction might be empty.")
    
    def update_selected_label(self):
        if self.start_date_edit.isEnabled():
            s = self.start_date_edit.date().toString("yyyy-MM-dd")
            e = self.end_date_edit.date().toString("yyyy-MM-dd")
            self.selected_range_label.setText(f"Selected range: {s} to {e}")
        else:
            self.selected_range_label.setText("Selected range: All (Time slicing disabled)")
        
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NetCDF Extractor with Map Preview")

        self._last_grid_df_full = None   # full grid inside bounding box (if available)
        self._last_grid_df_sample = None # sampled subset for global display
        self.resize(1100, 760)
        layout = QtWidgets.QVBoxLayout(self)

        # Path selection
        path_layout = QtWidgets.QHBoxLayout()
        self.path_edit = QtWidgets.QLineEdit()
        self.path_btn = QtWidgets.QPushButton("Browse data root")
        self.path_btn.clicked.connect(self.browse_path)
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(self.path_btn)
        layout.addLayout(path_layout)

        layout.addWidget(QtWidgets.QLabel("Files to Process:"))
        
        # 1. The Tree Widget
        self.file_tree = QtWidgets.QTreeWidget()
        self.file_tree.setHeaderHidden(True)
        layout.addWidget(self.file_tree)

        # 2. Select All / Deselect All Buttons
        sel_btn_layout = QtWidgets.QHBoxLayout()
        self.btn_sel_all = QtWidgets.QPushButton("Select All")
        self.btn_sel_all.clicked.connect(lambda: self.toggle_tree_selection(True))
        self.btn_desel_all = QtWidgets.QPushButton("Deselect All")
        self.btn_desel_all.clicked.connect(lambda: self.toggle_tree_selection(False))
        
        sel_btn_layout.addWidget(self.btn_sel_all)
        sel_btn_layout.addWidget(self.btn_desel_all)
        sel_btn_layout.addStretch()
        layout.addLayout(sel_btn_layout)

        # Coordinates and chunk sizes
        grid = QtWidgets.QGridLayout()
        grid.addWidget(QtWidgets.QLabel("Lat min"), 0, 0)
        self.lat_min = QtWidgets.QDoubleSpinBox(); self.lat_min.setDecimals(6); self.lat_min.setRange(-360, 360); self.lat_min.setValue(36.01)
        grid.addWidget(self.lat_min, 0, 1)
        grid.addWidget(QtWidgets.QLabel("Lat max"), 0, 2)
        self.lat_max = QtWidgets.QDoubleSpinBox(); self.lat_max.setDecimals(6); self.lat_max.setRange(-360, 360); self.lat_max.setValue(39.12)
        grid.addWidget(self.lat_max, 0, 3)

        grid.addWidget(QtWidgets.QLabel("Lon min"), 1, 0)
        self.lon_min = QtWidgets.QDoubleSpinBox(); self.lon_min.setDecimals(6); self.lon_min.setRange(-360, 360); self.lon_min.setValue(44.99)
        grid.addWidget(self.lon_min, 1, 1)
        grid.addWidget(QtWidgets.QLabel("Lon max"), 1, 2)
        self.lon_max = QtWidgets.QDoubleSpinBox(); self.lon_max.setDecimals(6); self.lon_max.setRange(-360, 360); self.lon_max.setValue(48.8)
        grid.addWidget(self.lon_max, 1, 3)

        grid.addWidget(QtWidgets.QLabel("Chunk lat"), 2, 0)
        self.chunk_lat = QtWidgets.QSpinBox(); self.chunk_lat.setRange(1, 100000); self.chunk_lat.setValue(5000)
        grid.addWidget(self.chunk_lat, 2, 1)
        grid.addWidget(QtWidgets.QLabel("Chunk lon"), 2, 2)
        self.chunk_lon = QtWidgets.QSpinBox(); self.chunk_lon.setRange(1, 100000); self.chunk_lon.setValue(5000)
        grid.addWidget(self.chunk_lon, 2, 3)

        layout.addLayout(grid)
         # Date pickers for user selection (disabled until detection)
        date_layout = QtWidgets.QHBoxLayout()
        self.start_date_edit = QtWidgets.QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setEnabled(False)
        self.end_date_edit = QtWidgets.QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setEnabled(False)
        date_layout.addWidget(QtWidgets.QLabel("Start date:"))
        date_layout.addWidget(self.start_date_edit)
        date_layout.addWidget(QtWidgets.QLabel("End date:"))
        date_layout.addWidget(self.end_date_edit)
        layout.addLayout(date_layout)
        
        self.date_range_label = QtWidgets.QLabel("Available range: N/A")
        layout.addWidget(self.date_range_label)
        
        self.selected_range_label = QtWidgets.QLabel("Selected range: All")
        self.selected_range_label.setStyleSheet("color: #1f4f7a; font-weight: bold;")
        layout.addWidget(self.selected_range_label)

        self.start_date_edit.dateChanged.connect(self.update_selected_label)
        self.end_date_edit.dateChanged.connect(self.update_selected_label)

      
        opts_layout = QtWidgets.QHBoxLayout()
        self.skip_chk = QtWidgets.QCheckBox("Skip existing")
        self.skip_chk.setChecked(True)
        opts_layout.addWidget(self.skip_chk)

        self.preview_btn = QtWidgets.QPushButton("Preview grids for selected folder")
        self.preview_btn.clicked.connect(self.preview_grids)
        opts_layout.addWidget(self.preview_btn)

        self.map_btn = QtWidgets.QPushButton("Show nodes on world map")
        self.map_btn.clicked.connect(self.show_map)
        self.map_btn.setEnabled(False)
        opts_layout.addWidget(self.map_btn)

        opts_layout.addStretch()
        layout.addLayout(opts_layout)

        # Save root
        save_layout = QtWidgets.QHBoxLayout()
        self.save_edit = QtWidgets.QLineEdit()
        self.save_edit.setPlaceholderText("Save extracted to (default: <data_root>/Extracted)")
        self.save_btn = QtWidgets.QPushButton("target folder")
        self.save_btn.clicked.connect(self.browse_save)
        save_layout.addWidget(self.save_edit)
        save_layout.addWidget(self.save_btn)
        layout.addLayout(save_layout)
        
        self.col_btn = QtWidgets.QPushButton("Select Columns")
        self.col_btn.clicked.connect(self.open_column_selector)
        self.col_btn.setEnabled(False)
        opts_layout.addWidget(self.col_btn)

        # Start button and progress
        btn_layout = QtWidgets.QHBoxLayout()
        self.start_btn = QtWidgets.QPushButton("Start extraction")
        self.start_btn.clicked.connect(self.start_extraction)
        start_row = QtWidgets.QHBoxLayout()
        start_row.addStretch()
        start_row.addWidget(self.start_btn)
        self.start_btn.setMinimumWidth(140)
        layout.addStretch()             # push everything above so the button stays at the bottom
        layout.addLayout(start_row)


        btn_layout.addWidget(self.start_btn)
        self.progress = QtWidgets.QProgressBar()
        btn_layout.addWidget(self.progress)
        layout.addLayout(btn_layout)

        # Grid preview area
        layout.addWidget(QtWidgets.QLabel("Grid Preview Summary"))
        self.grid_summary = QtWidgets.QLabel("")
        layout.addWidget(self.grid_summary)

        # Table for sample coordinates
        self.grid_table = QtWidgets.QTableWidget()
        self.grid_table.setColumnCount(2)
        self.grid_table.setHorizontalHeaderLabels(["lat", "lon"])
        self.grid_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.grid_table, stretch=1)

        # Export CSV button
        export_layout = QtWidgets.QHBoxLayout()
        self.export_btn = QtWidgets.QPushButton("Export grid list to CSV")
        self.export_btn.clicked.connect(self.export_grid_csv)
        self.export_btn.setEnabled(False)
        export_layout.addWidget(self.export_btn)
        export_layout.addStretch()
        layout.addLayout(export_layout)

        # Log area
        layout.addWidget(QtWidgets.QLabel("Log:"))
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, stretch=1)

        # Threadpool
        self.pool = QtCore.QThreadPool()

        # internal storage for last previewed grid
        self._last_grid_df = None
        
        #Date range detector
        self.detect_dates_btn = QtWidgets.QPushButton("Detect available date range")
        self.detect_dates_btn.clicked.connect(self.detect_date_range_for_selected)
        opts_layout.addWidget(self.detect_dates_btn)
        
        # Show detected range

        
        
        # Band selector (add near other controls in __init__)
        band_layout = QtWidgets.QHBoxLayout()
        band_layout.addWidget(QtWidgets.QLabel("Band:"))
        self.band_combo = QtWidgets.QComboBox()
        self.band_combo.addItem("All")            # default
        self.band_combo.setEnabled(False)         # enabled after preview/detection
        band_layout.addWidget(self.band_combo)
        opts_layout.addLayout(band_layout)        # or layout.addLayout(band_layout) where appropriate
                
        


       
        
        # Export grids to CSV (choose path) - allow user to pick a target path before extraction
        self._last_grid_df_full = None
        self._last_grid_df_sample = None
        
        # store detected overlapping date range (pandas.Timestamp or None)
        self._detected_start = None
        self._detected_end = None
        
        self._detected_columns = []   # List of all available columns
        self._user_selected_cols = None # List of columns user wants to keep
    
    def export_grid_csv_to_path(self):
        if (self._last_grid_df_full is None or len(self._last_grid_df_full) == 0) and \
           (self._last_grid_df_sample is None or len(self._last_grid_df_sample) == 0):
            self.append_log("No grid data available to export. Please run Preview grids first.")
            return
    
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save grid CSV", "grid_list.csv", "CSV files (*.csv)")
        if not fname:
            return
        try:
            # prefer full if available
            df_to_save = self._last_grid_df_full if (self._last_grid_df_full is not None and len(self._last_grid_df_full) > 0) else self._last_grid_df_sample
            df_to_save.to_csv(fname, index=False)
            self.append_log(f"Grid CSV saved to {fname}")
        except Exception as e:
            self.append_log(f"Failed to save CSV: {e}")
            
    
    def _get_time_range_for_file(self, file_path):
        """
        Return (start_date, end_date) as datetime.date for the given NetCDF file.
        Returns (None, None) if no time coordinate found or on error.
        Handles both numpy datetime64/pandas timestamps and cftime objects.
        """
        try:
            # Open dataset but avoid loading heavy data; decode_times=True to get cftime if present
            ds = xr.open_dataset(file_path.as_posix(), decode_times=True, chunks={})
            # detect time coordinate name
            time_name = None
            for cand in ["time", "Time", "times", "date", "dates", "valid_time"]:
                if cand in ds.coords:
                    time_name = cand
                    break
            if time_name is None:
                ds.close()
                return (None, None)
    
            times = ds[time_name].values
            if getattr(times, "size", 0) == 0:
                ds.close()
                return (None, None)
    
            # Helper to convert a single time element to datetime.date
            def to_date(t):
                # cftime objects
                if isinstance(t, (cftime.datetime, cftime.DatetimeNoLeap, cftime.Datetime360Day, cftime.DatetimeGregorian)):
                    return datetime.date(t.year, t.month, t.day)
                # numpy datetime64 or pandas Timestamp
                try:
                    # convert via pandas (works for numpy.datetime64 and pandas.Timestamp)
                    ts = pd.to_datetime(t)
                    return ts.date()
                except Exception:
                    # fallback: try attributes
                    if hasattr(t, "year") and hasattr(t, "month") and hasattr(t, "day"):
                        return datetime.date(int(t.year), int(t.month), int(t.day))
                    return None
    
            # compute min/max safely without converting entire array to pandas
            # times may be numpy array of cftime or datetime64
            # use Python min/max which works for cftime and numpy.datetime64
            try:
                raw_min = times.min()
                raw_max = times.max()
            except Exception:
                # fallback: iterate
                raw_min = None
                raw_max = None
                for v in times.ravel():
                    if raw_min is None or v < raw_min:
                        raw_min = v
                    if raw_max is None or v > raw_max:
                        raw_max = v
    
            start_date = to_date(raw_min)
            end_date = to_date(raw_max)
            ds.close()
            return (start_date, end_date)
        except Exception as e:
            # log and return None
            self.append_log(f"Failed to read time range from {file_path.name}: {e}")
            return (None, None)


    def _compute_overlapping_range(self, file_paths):
        """
        Given an iterable of Path objects (nc files), compute the overlapping (intersection)
        time range across all files. Returns (start_date, end_date) as datetime.date or (None, None).
        """
        overall_start = None
        overall_end = None
        any_valid = False
    
        for p in file_paths:
            s, e = self._get_time_range_for_file(p)
            if s is None or e is None:
                # skip files without valid time info
                continue
            any_valid = True
            # intersection: start = max(starts), end = min(ends)
            if overall_start is None or s > overall_start:
                overall_start = s
            if overall_end is None or e < overall_end:
                overall_end = e
    
        if not any_valid:
            return (None, None)
    
        # if intersection is empty
        if overall_start is not None and overall_end is not None and overall_start > overall_end:
            return (None, None)
    
        return (overall_start, overall_end)


    
    def detect_date_range_for_selected(self):
        
        file_paths_str = self.get_selected_files()
        if not file_paths_str:
            self.append_log("Please select files to detect date range.")
            return
            
        file_paths = [Path(p) for p in file_paths_str]
        
        self.append_log("Detecting overlapping date range...")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        
        try:
            start, end = self._compute_overlapping_range(file_paths)
            start_date = self._to_date(start)
            end_date = self._to_date(end)
            
            if start_date is None or end_date is None:
                self.date_range_label.setText("Available range: N/A or no common overlap")
                self._detected_start = None
                self._detected_end = None
                self.start_date_edit.setEnabled(False)
                self.end_date_edit.setEnabled(False)
                self.append_log("Could not determine a common overlapping date range across files.")
                return
            
            self._detected_start = start_date
            self._detected_end = end_date
            self.date_range_label.setText(f"Available range: {start_date}  —  {end_date}")
            
            qstart = QtCore.QDate(start_date.year, start_date.month, start_date.day)
            qend   = QtCore.QDate(end_date.year, end_date.month, end_date.day)
            self.start_date_edit.setDateRange(qstart, qend)
            self.end_date_edit.setDateRange(qstart, qend)
            self.start_date_edit.setDate(qstart)
            self.end_date_edit.setDate(qend)
            self.start_date_edit.setEnabled(True)
            self.end_date_edit.setEnabled(True)
            self.update_selected_label() 
            self.append_log(f"Detected overlapping range: {start_date} to {end_date}")


        finally:
            QtWidgets.QApplication.restoreOverrideCursor()


    def browse_path(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Select data root")
        if d:
            self.path_edit.setText(d)
            self.populate_folders(d)

    def browse_save(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Select save root")
        if d:
            self.save_edit.setText(d)

    def populate_folders(self, root):

        self.file_tree.clear()
        
        p = Path(root)
        if not p.exists():
            return

        # Iterate over subdirectories
        for subdir in sorted([x for x in p.iterdir() if x.is_dir()]):
            # Create Parent Item (Folder)
            folder_item = QtWidgets.QTreeWidgetItem([subdir.name])
            # Set Tristate: checking parent auto-checks children
            folder_item.setFlags(folder_item.flags() | QtCore.Qt.ItemIsTristate | QtCore.Qt.ItemIsUserCheckable)
            folder_item.setCheckState(0, QtCore.Qt.Unchecked)
            
            # Find .nc files inside this folder
            nc_files = sorted(list(subdir.glob("*.nc")))
            
            if nc_files:
                # Only add the folder if it contains NC files
                self.file_tree.addTopLevelItem(folder_item)
                
                for f in nc_files:
                    # Create Child Item (File)
                    file_item = QtWidgets.QTreeWidgetItem([f.name])
                    file_item.setFlags(file_item.flags() | QtCore.Qt.ItemIsUserCheckable)
                    file_item.setCheckState(0, QtCore.Qt.Unchecked)
                    
                    # Store full path in UserRole (invisible data) for easy retrieval later
                    file_item.setData(0, QtCore.Qt.UserRole, str(f.absolute()))
                    
                    folder_item.addChild(file_item)

        self.file_tree.expandAll()
    


    def start_extraction(self):
        data_root = self.path_edit.text().strip()
        if not data_root:
            self.append_log("Please select data root path.")
            return

      
        files_to_process = self.get_selected_files()
        
        if not files_to_process:
            self.append_log("Please select at least one file to process.")
            return
            
        self.append_log(f"Starting extraction for {len(files_to_process)} files...")
        
        data_root = self.path_edit.text().strip()
        if not data_root:
            self.log.appendPlainText("Please select data root path.")
            return
    
   
        save_root = self.save_edit.text().strip() or os.path.join(data_root, "Extracted")
        selected_start = None
        selected_end = None
        if self.start_date_edit.isEnabled() and self.end_date_edit.isEnabled():
            sd = self.start_date_edit.date()
            ed = self.end_date_edit.date()
            selected_start = pd.Timestamp(year=sd.year(), month=sd.month(), day=sd.day())
            selected_end = pd.Timestamp(year=ed.year(), month=ed.month(), day=ed.day())
        band_sel = None
        if self.band_combo.isEnabled():
            band_sel = self.band_combo.currentText()
            if band_sel == "All":
                band_sel = None

        params = {
            "data_path": data_root,
            "files": files_to_process,
            
            "lat_min": self.lat_min.value(),
            "lat_max": self.lat_max.value(),
            "lon_min": self.lon_min.value(),
            "lon_max": self.lon_max.value(),
            "chunk_lat": self.chunk_lat.value(),
            "chunk_lon": self.chunk_lon.value(),
            "skip_existing": self.skip_chk.isChecked(),
            "save_root": save_root,
            "date_start": selected_start,
            "date_end": selected_end,
            "band": band_sel, 
            "columns": self._user_selected_cols 
        }


        self.start_btn.setEnabled(False)
        worker = ExtractWorker(params)
        worker.signals.progress.connect(self.progress.setValue)
        worker.signals.busy.connect(self.on_worker_busy)
        worker.signals.log.connect(self.append_log)
        worker.signals.finished.connect(self.on_finished)
        self.pool.start(worker)

    def append_log(self, text):
        self.log.appendPlainText(text)
    def on_worker_busy(self,is_busy):
        if is_busy:
           self.progress.setRange(0, 0)
        else:
           self.progress.setRange(0, 100)


    def on_finished(self):
        self.start_btn.setEnabled(True)
        self.append_log("Finished.")

    def preview_grids(self):
        """
        Preview available lat/lon grid points inside the bounding box.
        Uses the first selected file from the tree.
        """
        # --- NEW: Get files from the tree instead of folder list ---
        selected_files = self.get_selected_files()
        
        if not selected_files:
            self.append_log("Please select at least one file to preview.")
            return

        # Just pick the first selected file to serve as the "Sample"
        sample_file = Path(selected_files[0])
        # -----------------------------------------------------------

        self.append_log(f"Previewing grids from {sample_file.name} ...")
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            lat_min = float(self.lat_min.value())
            lat_max = float(self.lat_max.value())
            lon_min = float(self.lon_min.value())
            lon_max = float(self.lon_max.value())
    
            ds = xr.open_dataset(sample_file.as_posix(), chunks={})
            
            # Detect coordinate names dynamically
            lat_name = None
            lon_name = None
            for candidate in ["lat", "latitude", "y"]:
                if candidate in ds.coords:
                    lat_name = candidate
                    break
            for candidate in ["lon", "longitude", "x"]:
                if candidate in ds.coords:
                    lon_name = candidate
                    break

            if lat_name is None or lon_name is None:
                self.append_log("Could not find lat/lon coordinates in the dataset.")
                self.grid_summary.setText("No coordinates found")
                QtWidgets.QApplication.restoreOverrideCursor()
                return

            # --- Detect available columns for the "Select Columns" UI ---
            time_name = None
            # Look for common time variable names
            for cand in ["time", "valid_time", "Time", "times", "date"]:
                if cand in ds.coords:
                    time_name = cand
                    break
            
            # Start list with lat/lon and time
            self._detected_columns = ["lat", "lon"]
            if time_name:
                self._detected_columns.append(time_name)

            # Add band/level coords
            for cand in ["band", "bands", "level", "depth"]:
                if cand in ds.coords or cand in ds.dims:
                    self._detected_columns.append(cand)
                    break
            
            # Add actual data variables
            self._detected_columns.extend(list(ds.data_vars.keys()))
            
            # Remove duplicates
            self._detected_columns = list(dict.fromkeys(self._detected_columns))
            
            self.col_btn.setEnabled(True)
            self._user_selected_cols = None # Reset previous selection
            self.append_log(f"Detected variables: {', '.join(ds.data_vars.keys())}")
            # -----------------------------------------------------------
    
            lats = ds[lat_name].values
            lons = ds[lon_name].values
    
            full_df = None
            sample_df = pd.DataFrame(columns=["lat", "lon"])
            
            # Calculate grid preview
            if lats.ndim == 2 and lons.ndim == 2:
                lat_flat = lats.ravel()
                lon_flat = lons.ravel()
                coords = pd.DataFrame({"lat": lat_flat, "lon": lon_flat})
                coords = coords[
                    (coords["lat"] >= lat_min) & (coords["lat"] <= lat_max) &
                    (coords["lon"] >= lon_min) & (coords["lon"] <= lon_max)
                ].reset_index(drop=True)
                count = len(coords)
                sample_df = coords.head(200)
                full_df = coords
            else:
                lat_vals = lats.ravel()
                lon_vals = lons.ravel()
                # avoid full expansion for huge grids
                if lat_vals.size * lon_vals.size > 5_000_000:
                    self.append_log("Grid is very large; summarizing without full expansion.")
                    lat_in = lat_vals[(lat_vals >= lat_min) & (lat_vals <= lat_max)]
                    lon_in = lon_vals[(lon_vals >= lon_min) & (lon_vals <= lon_max)]
                    count = lat_in.size * lon_in.size
                    if lat_in.size == 0 or lon_in.size == 0:
                        sample_df = pd.DataFrame(columns=["lat", "lon"])
                        full_df = None
                    else:
                        sample_df = pd.DataFrame({
                            "lat": np.repeat(lat_in[:50], min(50, lon_in.size)),
                            "lon": list(lon_in[:50]) * min(50, lat_in.size)
                        }).head(200)
                        full_df = None
                else:
                    mesh_lat, mesh_lon = np.meshgrid(lat_vals, lon_vals, indexing="ij")
                    coords = pd.DataFrame({"lat": mesh_lat.ravel(), "lon": mesh_lon.ravel()})
                    coords = coords[
                        (coords["lat"] >= lat_min) & (coords["lat"] <= lat_max) &
                        (coords["lon"] >= lon_min) & (coords["lon"] <= lon_max)
                    ].reset_index(drop=True)
                    count = len(coords)
                    sample_df = coords.head(200)
                    full_df = coords
            
                # Band detection logic
                band_values = None
                for cand in ["band", "bands", "band_index", "bnd", "bnds"]:
                    if cand in ds.coords:
                        band_values = ds[cand].values
                        break
                    if cand in ds.dims:
                        if cand in ds.coords:
                            band_values = ds[cand].values
                        else:
                            band_values = np.arange(ds.sizes.get(cand, 0))
                        break
                
                if band_values is None:
                    for cand in ["band", "bands", "band_index", "bnd", "bnds"]:
                        if cand in ds.variables:
                            try:
                                band_values = ds[cand].values
                                break
                            except Exception:
                                pass
                
                if band_values is not None:
                    try:
                        band_list = [str(v) for v in np.asarray(band_values).ravel()]
                    except Exception:
                        band_list = [str(band_values)]
                else:
                    band_list = []

            self.band_combo.clear()
            self.band_combo.addItem("All")
            if len(band_list) > 0:
                for v in band_list:
                    self.band_combo.addItem(v)
                self.band_combo.setEnabled(True)
                self.append_log(f"Detected band values: {', '.join(band_list)}")
            else:
                self.band_combo.setEnabled(False)
                self.append_log("No band coordinate detected.")
            
            if full_df is not None:
                self._last_grid_df_full = full_df
                self._last_grid_df_sample = full_df.sample(n=min(len(full_df), 5000), random_state=1).reset_index(drop=True)
            else:
                self._last_grid_df_full = None
                self._last_grid_df_sample = sample_df

            if (self._last_grid_df_full is not None and len(self._last_grid_df_full) > 0) or \
               (self._last_grid_df_sample is not None and len(self._last_grid_df_sample) > 0):
                self.export_btn.setEnabled(True)
                self.map_btn.setEnabled(True)
            else:
                self.export_btn.setEnabled(False)
                self.map_btn.setEnabled(False)
    
            self.grid_table.setRowCount(0)
            for r, row in sample_df.iterrows():
                self.grid_table.insertRow(r)
                lat_item = QtWidgets.QTableWidgetItem(f"{row['lat']:.6f}")
                lon_item = QtWidgets.QTableWidgetItem(f"{row['lon']:.6f}")
                self.grid_table.setItem(r, 0, lat_item)
                self.grid_table.setItem(r, 1, lon_item)
    
            summary_text = f"Sample file: {sample_file.name}  |  Grid points inside box: {count}"
            self.grid_summary.setText(summary_text)
            self.append_log(f"Preview complete. Found {count} grid points inside the bounding box.")

        except Exception as e:
            self.append_log(f"Error during preview: {e}")
            import traceback
            print(traceback.format_exc())
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()


    def export_grid_csv(self):
        """
        Export the last previewed grid to CSV.
        Prefer the full grid (if available), otherwise export the sampled grid.
        """
        # choose the dataframe to export
        df_full = getattr(self, "_last_grid_df_full", None)
        df_sample = getattr(self, "_last_grid_df_sample", None)
    
        # prefer full, else sample
        if df_full is not None and len(df_full) > 0:
            df_to_save = df_full
            source = "full"
        elif df_sample is not None and len(df_sample) > 0:
            df_to_save = df_sample
            source = "sample"
        else:
            self.append_log("No grid data to export. Please run Preview grids first.")
            return
    
        # ask user where to save
        fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save grid CSV", "grid_list.csv", "CSV files (*.csv)")
        if not fname:
            self.append_log("Export cancelled by user.")
            return
    
        try:
            df_to_save.to_csv(fname, index=False)
            self.append_log(f"Grid CSV ({source}) saved to: {fname}")
        except Exception as e:
            self.append_log(f"Failed to save CSV: {e}")
    
            
    def show_map(self):
        """
        Create a folium map from the last previewed grid and show it in a QWebEngineView dialog.
        Falls back to opening the HTML in the system browser if QWebEngine is unavailable.
        """
        # prefer sample for global display, but keep full for zoomed-in layer
        display_df = self._last_grid_df_sample.copy() if self._last_grid_df_sample is not None else None
        full_df = self._last_grid_df_full
    
        if display_df is None and full_df is not None:
            display_df = full_df.sample(n=min(len(full_df), 5000), random_state=1).reset_index(drop=True)
    
        if full_df is not None and len(full_df) > 0:
            center_lat = float(full_df['lat'].mean())
            center_lon = float(full_df['lon'].mean())
        elif display_df is not None and len(display_df) > 0:
            center_lat = float(display_df['lat'].mean())
            center_lon = float(display_df['lon'].mean())
        else:
            self.append_log("No grid data available to map.")
            return
    
        m = folium.Map(location=[center_lat, center_lon], zoom_start=6, tiles="CartoDB positron")
        formatter = "function(num) {return L.Util.formatNum(num, 5);};"
        MousePosition(
            position='topright',
            separator=' | ',
            empty_string='NaN',
            lng_first=False,
            num_digits=20,
            prefix='Coordinates:',
            lat_formatter=formatter,
            lng_formatter=formatter,
        ).add_to(m)
    
        lat_min = float(self.lat_min.value()); lat_max = float(self.lat_max.value())
        lon_min = float(self.lon_min.value()); lon_max = float(self.lon_max.value())
        folium.Rectangle(bounds=[[lat_min, lon_min], [lat_max, lon_max]],
                         color="#ff7800", weight=2, fill=False).add_to(m)
    

    
        # add full nodes as a separate layer (only if available)
        if full_df is not None and len(full_df) > 0:
            full_layer = folium.FeatureGroup(name="All nodes in bbox", show=False)
            for _, row in full_df.iterrows():
                folium.CircleMarker(location=[row["lat"], row["lon"]],
                                    radius=2,
                                    color="#ff0000",
                                    fill=True,
                                    fill_opacity=0.9).add_to(full_layer)
            full_layer.add_to(m)
    
        folium.LayerControl().add_to(m)
    
        tmp = tempfile.NamedTemporaryFile(suffix=".html", delete=False)
        tmp_name = tmp.name
        tmp.close()
        m.save(tmp_name)
    
        if WEB_ENGINE_AVAILABLE:
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Grid Map Preview")
            dlg.resize(1000, 700)
            layout = QtWidgets.QVBoxLayout(dlg)
            web = QtWebEngineWidgets.QWebEngineView()
            web.load(QtCore.QUrl.fromLocalFile(tmp_name))
            layout.addWidget(web)
    
            toolbar = QtWidgets.QHBoxLayout()
            open_btn = QtWidgets.QPushButton("Open in browser")
            def open_browser():
                webbrowser.open(tmp_name)
            open_btn.clicked.connect(open_browser)
            toolbar.addWidget(open_btn)
    
            save_btn = QtWidgets.QPushButton("Save HTML As")
            def save_html():
                fname, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save map HTML", "grid_map.html", "HTML files (*.html)")
                if fname:
                    try:
                        import shutil
                        shutil.copy(tmp_name, fname)
                        self.append_log(f"Map HTML saved to {fname}")
                    except Exception as e:
                        self.append_log(f"Failed to save HTML: {e}")
            save_btn.clicked.connect(save_html)
            toolbar.addWidget(save_btn)
            toolbar.addStretch()
            layout.addLayout(toolbar)
    
            dlg.exec_()
        else:
            self.append_log("PyQtWebEngine not available; opening map in system browser.")
            webbrowser.open(tmp_name)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def main():
    if not QtWidgets.QApplication.instance():
        app = QtWidgets.QApplication(sys.argv)
    else:
        app = QtWidgets.QApplication.instance()

    app.setStyleSheet(qss_text)
    
    icon_path = resource_path("icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QtGui.QIcon(icon_path))
    else:
        print("Icon file not found:", icon_path)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()