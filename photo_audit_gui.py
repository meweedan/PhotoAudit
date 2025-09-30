#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Photo Audit GUI — macOS & Windows
Compares spreadsheet reference numbers with image filenames to find:
  1) Items NOT photographed (no matching image) -> *_not_photographed.xlsx
  2) Items photographed, enriched with image metadata -> *_photographed.xlsx

Matching logic:
- Normalizes both reference values and filenames by stripping non-alphanumerics.
- By default, considers a match if the (normalized) reference appears as a substring of the (normalized) filename.
- You can switch to "Exact match" if your filenames are exactly the reference.

Dependencies:
  pip install pandas openpyxl pillow pillow-heif
"""

import os
import sys
import re
import traceback
import datetime as dt
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

# GUI
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Data / Excel
import pandas as pd

# Images / EXIF
from PIL import Image, ExifTags

# Optional HEIC support (won't crash if missing)
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:
    pass

# ---------- Config ----------

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".heic", ".heif"}
DEFAULT_OUTPUT_BASENAME = "photo_audit"

# ---------- Helpers ----------

def norm(s: str) -> str:
    """Normalize to alphanumeric lowercase for robust matching."""
    if s is None:
        return ""
    return re.sub(r"[^A-Za-z0-9]+", "", str(s)).lower()

def list_images(folder: str) -> List[str]:
    files = []
    for root, _, filenames in os.walk(folder):
        for fn in filenames:
            ext = os.path.splitext(fn)[1].lower()
            if ext in SUPPORTED_EXTS:
                files.append(os.path.join(root, fn))
    return files

def get_exif_dict(img: Image.Image) -> Dict:
    exif = {}
    try:
        raw = img._getexif() or {}
    except Exception:
        raw = {}
    for k, v in raw.items():
        tag = ExifTags.TAGS.get(k, k)
        exif[tag] = v
    return exif

def try_parse_exif_datetime(exif: Dict) -> Optional[str]:
    # Prefer DateTimeOriginal -> DateTime -> DateTimeDigitized
    for key in ("DateTimeOriginal", "DateTime", "DateTimeDigitized"):
        val = exif.get(key)
        if val:
            # EXIF format: "YYYY:MM:DD HH:MM:SS"
            try:
                dt_obj = dt.datetime.strptime(val, "%Y:%m:%d %H:%M:%S")
                return dt_obj.isoformat(sep=" ", timespec="seconds")
            except Exception:
                # Some cameras add sub-seconds or different patterns; keep raw if it fails
                return str(val)
    return None

def get_file_times(path: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Returns (created_time, modified_time) as ISO strings where possible.
    Note: 'created' availability differs by OS.
    """
    try:
        stat = os.stat(path)
        modified = dt.datetime.fromtimestamp(stat.st_mtime).isoformat(sep=" ", timespec="seconds")
        created = None
        # macOS provides st_birthtime; Windows maps st_ctime to creation
        if hasattr(stat, "st_birthtime"):  # macOS
            created = dt.datetime.fromtimestamp(stat.st_birthtime).isoformat(sep=" ", timespec="seconds")
        else:
            # On Windows, st_ctime is creation; on Linux it's inode change (not true creation)
            if sys.platform.startswith("win"):
                created = dt.datetime.fromtimestamp(stat.st_ctime).isoformat(sep=" ", timespec="seconds")
        return created, modified
    except Exception:
        return None, None

def image_metadata(path: str) -> Dict:
    meta = {
        "image_path": path,
        "image_name": os.path.basename(path),
        "image_ext": os.path.splitext(path)[1].lower(),
        "file_created": None,
        "file_modified": None,
        "exif_datetime": None,
        "exif_make": None,
        "exif_model": None,
        "width": None,
        "height": None,
    }
    try:
        created, modified = get_file_times(path)
        meta["file_created"] = created
        meta["file_modified"] = modified

        with Image.open(path) as im:
            meta["width"], meta["height"] = getattr(im, "width", None), getattr(im, "height", None)
            exif = get_exif_dict(im)
            meta["exif_datetime"] = try_parse_exif_datetime(exif)
            meta["exif_make"] = exif.get("Make")
            meta["exif_model"] = exif.get("Model")
    except Exception:
        # Keep metadata that we could extract; log silently
        pass
    return meta

def build_match_index(images: List[str]) -> Dict[str, List[str]]:
    """
    Build a normalized filename index -> list of full paths
    """
    idx: Dict[str, List[str]] = {}
    for p in images:
        key = norm(os.path.basename(p))
        idx.setdefault(key, []).append(p)
    return idx

def find_best_match(ref_norm: str, filename_index: Dict[str, List[str]], substring: bool = True) -> Optional[str]:
    """
    Find any image whose (normalized) filename contains the ref (or equals it if substring=False).
    If multiple matches, prefer the shortest filename key containing the ref (heuristic for specificity).
    """
    if not ref_norm:
        return None
    if substring:
        candidates = []
        for fname_key, paths in filename_index.items():
            if ref_norm in fname_key:
                candidates.extend((fname_key, p) for p in paths)
        if not candidates:
            return None
        # prefer the closest length to the ref (simple heuristic)
        candidates.sort(key=lambda t: (abs(len(t[0]) - len(ref_norm)), len(t[0])))
        return candidates[0][1]
    else:
        # exact match of whole normalized filename (without extension)
        for fname_key, paths in filename_index.items():
            if ref_norm == fname_key:
                return paths[0]
        return None

@dataclass
class JobConfig:
    excel_path: str
    sheet_name: str
    ref_col: str
    image_folder: str
    substring_match: bool = True
    case_sensitive: bool = False
    output_folder: Optional[str] = None

def run_audit(cfg: JobConfig) -> Tuple[str, str]:
    """
    Main audit function. Returns (photographed_xlsx, not_photographed_xlsx) paths.
    """
    # Load excel
    df = pd.read_excel(cfg.excel_path, sheet_name=cfg.sheet_name, dtype=str)
    if cfg.ref_col not in df.columns:
        raise ValueError(f"Reference column '{cfg.ref_col}' not found in sheet '{cfg.sheet_name}'.")

    # Normalize the reference series (keep original too)
    df["_ref_original"] = df[cfg.ref_col].astype(str)
    df["_ref_norm"] = df["_ref_original"].apply(norm)

    # Scan images
    images = list_images(cfg.image_folder)
    if not images:
        raise ValueError("No images found in the selected folder with the supported extensions.")

    # Build filename index
    fname_index = build_match_index(images)

    # Match
    matched_paths: List[Optional[str]] = []
    for ref in df["_ref_norm"]:
        p = find_best_match(ref, fname_index, substring=cfg.substring_match)
        matched_paths.append(p)
    df["_image_path"] = matched_paths
    df["_has_image"] = df["_image_path"].notna()

    # Split photographed vs not
    df_photo = df[df["_has_image"]].copy()
    df_missing = df[~df["_has_image"]].copy()

    # Enrich photographed with metadata
    meta_rows = []
    for p in df_photo["_image_path"]:
        meta_rows.append(image_metadata(p))
    meta_df = pd.DataFrame(meta_rows)

    # Merge back by image_path
    df_photo = df_photo.merge(meta_df, left_on="_image_path", right_on="image_path", how="left")

    # Prepare outputs (drop helper cols except keep _ref_original maybe)
    keep_all_cols = [c for c in df.columns if not c.startswith("_")]  # original spreadsheet columns
    # For photographed: keep original cols + metadata + handy columns
    ordered_photo_cols = keep_all_cols + [
        "image_name", "image_ext", "image_path",
        "file_created", "file_modified",
        "exif_datetime", "exif_make", "exif_model",
        "width", "height",
        # also include the reference used for matching
        # (keep for auditing)
    ]
    # For missing: only original spreadsheet columns
    df_missing_out = df_missing[keep_all_cols].copy()

    # Output filenames
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = DEFAULT_OUTPUT_BASENAME
    out_dir = cfg.output_folder or cfg.image_folder
    os.makedirs(out_dir, exist_ok=True)

    photographed_xlsx = os.path.join(out_dir, f"{base}_{ts}_photographed.xlsx")
    missing_xlsx = os.path.join(out_dir, f"{base}_{ts}_not_photographed.xlsx")

    # Write
    with pd.ExcelWriter(photographed_xlsx, engine="openpyxl") as writer:
        df_photo[ordered_photo_cols].to_excel(writer, index=False, sheet_name="photographed")

    with pd.ExcelWriter(missing_xlsx, engine="openpyxl") as writer:
        df_missing_out.to_excel(writer, index=False, sheet_name="not_photographed")

    return photographed_xlsx, missing_xlsx

# ---------- GUI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Photo Audit")
        self.geometry("720x560")
        self.minsize(680, 520)

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.ref_col = tk.StringVar()
        self.image_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.match_mode = tk.StringVar(value="substring")
        self.case_sensitive = tk.BooleanVar(value=False)

        self.sheet_combo: Optional[ttk.Combobox] = None
        self.ref_combo: Optional[ttk.Combobox] = None

        self._build_ui()

    def _build_ui(self):
        pad = 10

        frm = ttk.Frame(self, padding=pad)
        frm.pack(fill="both", expand=True)

        # Excel picker
        excel_label = ttk.Label(frm, text="Excel file:")
        excel_label.grid(row=0, column=0, sticky="w")
        excel_entry = ttk.Entry(frm, textvariable=self.excel_path, width=60)
        excel_entry.grid(row=0, column=1, sticky="we", padx=(5, 5))
        excel_btn = ttk.Button(frm, text="Select Excel", command=self.pick_excel)
        excel_btn.grid(row=0, column=2, sticky="e")

        # Sheet
        sheet_label = ttk.Label(frm, text="Sheet:")
        sheet_label.grid(row=1, column=0, sticky="w", pady=(pad, 0))
        self.sheet_combo = ttk.Combobox(frm, textvariable=self.sheet_name, state="readonly", width=40)
        self.sheet_combo.grid(row=1, column=1, sticky="w", padx=(5, 5), pady=(pad, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        # Ref column
        ref_label = ttk.Label(frm, text="Reference column:")
        ref_label.grid(row=2, column=0, sticky="w", pady=(pad, 0))
        self.ref_combo = ttk.Combobox(frm, textvariable=self.ref_col, state="readonly", width=40)
        self.ref_combo.grid(row=2, column=1, sticky="w", padx=(5, 5), pady=(pad, 0))

        # Image folder
        img_label = ttk.Label(frm, text="Image folder:")
        img_label.grid(row=3, column=0, sticky="w", pady=(pad, 0))
        img_entry = ttk.Entry(frm, textvariable=self.image_folder, width=60)
        img_entry.grid(row=3, column=1, sticky="we", padx=(5, 5), pady=(pad, 0))
        img_btn = ttk.Button(frm, text="Select Folder", command=self.pick_folder)
        img_btn.grid(row=3, column=2, sticky="e", pady=(pad, 0))

        # Output folder
        out_label = ttk.Label(frm, text="Output folder (optional):")
        out_label.grid(row=4, column=0, sticky="w", pady=(pad, 0))
        out_entry = ttk.Entry(frm, textvariable=self.output_folder, width=60)
        out_entry.grid(row=4, column=1, sticky="we", padx=(5, 5), pady=(pad, 0))
        out_btn = ttk.Button(frm, text="Select Output", command=self.pick_output)
        out_btn.grid(row=4, column=2, sticky="e", pady=(pad, 0))

        # Options
        opt_frame = ttk.LabelFrame(frm, text="Matching options", padding=pad)
        opt_frame.grid(row=5, column=0, columnspan=3, sticky="we", pady=(pad, 0))

        ttk.Radiobutton(opt_frame, text="Filename contains reference (recommended)", value="substring",
                        variable=self.match_mode).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opt_frame, text="Exact match (filename equals reference, ignoring punctuation)", value="exact",
                        variable=self.match_mode).grid(row=1, column=0, sticky="w")

        ttk.Checkbutton(opt_frame, text="Case sensitive (normally off)", variable=self.case_sensitive).grid(
            row=2, column=0, sticky="w", pady=(5, 0)
        )

        # Run
        run_btn = ttk.Button(frm, text="Run", command=self.on_run)
        run_btn.grid(row=6, column=0, columnspan=3, pady=(pad*1.5, 0))

        # Log / status
        self.log = tk.Text(frm, height=12, wrap="word")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(pad, 0))
        frm.rowconfigure(7, weight=1)
        frm.columnconfigure(1, weight=1)

        self._log("Select your Excel, sheet, reference column, and image folder; then click Run.")

    def _log(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def pick_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        self.excel_path.set(path)
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
            self.sheet_combo["values"] = sheets
            if sheets:
                self.sheet_name.set(sheets[0])
                self.on_sheet_selected()
            self._log(f"Loaded workbook with sheets: {', '.join(sheets)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel: {e}")
            self._log(traceback.format_exc())

    def on_sheet_selected(self, *_):
        path = self.excel_path.get()
        sheet = self.sheet_name.get()
        if not path or not sheet:
            return
        try:
            # Peek columns only
            df = pd.read_excel(path, sheet_name=sheet, nrows=1)
            cols = list(df.columns)
            self.ref_combo["values"] = cols
            if cols:
                self.ref_col.set(cols[0])
            self._log(f"Sheet '{sheet}' columns: {', '.join(cols)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheet: {e}")
            self._log(traceback.format_exc())

    def pick_folder(self):
        folder = filedialog.askdirectory(title="Select image folder")
        if folder:
            self.image_folder.set(folder)

    def pick_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder.set(folder)

    def on_run(self):
        try:
            if not self.excel_path.get() or not os.path.exists(self.excel_path.get()):
                messagebox.showwarning("Missing", "Please select a valid Excel file.")
                return
            if not self.sheet_name.get():
                messagebox.showwarning("Missing", "Please choose a sheet.")
                return
            if not self.ref_col.get():
                messagebox.showwarning("Missing", "Please choose the reference column.")
                return
            if not self.image_folder.get() or not os.path.isdir(self.image_folder.get()):
                messagebox.showwarning("Missing", "Please select a valid image folder.")
                return

            substring = (self.match_mode.get() == "substring")
            cfg = JobConfig(
                excel_path=self.excel_path.get(),
                sheet_name=self.sheet_name.get(),
                ref_col=self.ref_col.get(),
                image_folder=self.image_folder.get(),
                substring_match=substring,
                case_sensitive=self.case_sensitive.get(),
                output_folder=self.output_folder.get() or None
            )
            self._log("Running audit...")
            photographed_xlsx, missing_xlsx = run_audit(cfg)
            self._log(f"✓ Created: {photographed_xlsx}")
            self._log(f"✓ Created: {missing_xlsx}")
            messagebox.showinfo("Done", f"Created:\n{photographed_xlsx}\n{missing_xlsx}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log("ERROR:\n" + traceback.format_exc())

def main():
    app = App()
    # Better default styling on macOS
    try:
        app.style = ttk.Style()
        if sys.platform == "darwin":
            app.style.theme_use("clam")
    except Exception:
        pass
    app.mainloop()

if __name__ == "__main__":
    main()
