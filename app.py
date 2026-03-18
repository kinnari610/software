import os
import sys

import pandas as pd
from tkinter import *
from tkinter import ttk
import re
import sys

# ===== FILE PATHS =====
param_file = r"C:\Users\kinnari\Downloads\Telegram Desktop\TEST_PARAMETER.xlsx"
data_file = r"C:\Users\kinnari\Downloads\Telegram Desktop\test_data.xlsx"
# External workbook providing No-Load test values keyed by Assembly Number
no_load_file = r"C:\Users\kinnari\Downloads\Book1.xlsx"

# ===== LOAD SHEETS =====

# These will be loaded on demand in load_data()

# ===== FUNCTIONS =====

# global container for plate data rows read from workbook
plate_rows = []

# computed values (populated when loading data)
ambient_temp_val = 0.0
resistance_cold_val = 0.0
resistance_20_val = 0.0


def _parse_float(val, default=0.0):
    """Parse a value into float.

    Handles numbers stored as strings (e.g. "13.616", "13.616 Ohm", "13,616").
    Returns default if parsing fails.
    """
    if val is None:
        return default
    # if it is already numeric, just return it
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return default

    # Normalize common separators
    s = s.replace(',', '.')

    # Try direct float conversion first
    try:
        return float(s)
    except Exception:
        pass

    # Try to extract first numeric token (e.g. "13.616 ohm")
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if m:
        try:
            return float(m.group(0))
        except Exception:
            pass

    return default


def _normalize_assembly_id(val):
    """Normalize a value for assembly number comparison.

    This helps when assembly numbers are stored as numbers in Excel (e.g. 123.0)
    but are entered/shown as "123".
    """
    if val is None:
        return ""

    if isinstance(val, (int, float)):
        # Convert floats like 123.0 to "123".
        try:
            f = float(val)
            if f.is_integer():
                return str(int(f))
            return str(f)
        except Exception:
            return str(val).strip()

    s = str(val).strip()
    if not s:
        return ""

    # If value looks like a float with a trailing .0 (e.g. "123.0"), strip it
    m = re.match(r"^(-?\d+)\.0+$", s)
    if m:
        return m.group(1)

    return s


def _extract_no_load_table(assembly_no: str):
    """Extract the no-load test table for the given assembly from the no-load workbook.

    The workbook is expected to contain a section like this (one of many assembly blocks):

        assembly no: 2604810-H
        %Voltage  Volts  AMP  WATT  Current Ratio  Value of Type Tested motor  Remark
        110       ...
        100       ...
        90        ...

    This function searches all sheets in the workbook and returns the first matching
    table as a list-of-lists (rows), or None if no matching block is found.
    """
    try:
        sheets = pd.read_excel(no_load_file, sheet_name=None, header=None)
    except Exception as e:
        print("Warning: failed to load no-load test file:", e)
        return None

    asm_norm = _normalize_assembly_id(assembly_no)
    if not asm_norm:
        return None

    def _find_in_df(df):
        # Locate the row containing the assembly number (and ideally the label 'assembly no').
        asm_row = None
        for idx, row in df.iterrows():
            joined = " ".join(str(x) if not pd.isna(x) else "" for x in row)
            if re.search(r"assembly\s*no", joined, flags=re.I) and asm_norm in joined:
                asm_row = idx
                break
            # fallback: match just the assembly id itself if it's in any cell
            if any(_normalize_assembly_id(x) == asm_norm for x in row):
                asm_row = idx
                break

        if asm_row is None:
            return None

        # Find the header row (contains %Voltage or Volts/AMP/WATT)
        header_row = None
        for i in range(asm_row + 1, min(len(df), asm_row + 10)):
            joined = " ".join(str(x).strip() for x in df.iloc[i].tolist() if not pd.isna(x))
            if "%Voltage" in joined or ("Volts" in joined and "AMP" in joined) or "WATT" in joined:
                header_row = i
                break

        if header_row is None:
            header_row = asm_row + 1

        # Collect rows from header_row until blank row or next assembly marker.
        rows = []
        for i in range(header_row, len(df)):
            row = df.iloc[i].tolist()
            # stop when completely blank row (no non-empty cells)
            if all(pd.isna(x) or str(x).strip() == "" for x in row):
                break

            joined = " ".join(str(x) if not pd.isna(x) else "" for x in row)
            # stop if we hit another assembly section or new test section
            if re.search(r"assembly\s*no", joined, flags=re.I) or re.search(r"locked\s+rotor", joined, flags=re.I):
                break

            rows.append(["" if pd.isna(x) else str(x).strip() for x in row])

        return rows or None

    # Try every sheet until we find a matching block.
    for sheet_name, df in sheets.items():
        table = _find_in_df(df)
        if table:
            return table

    return None


def _get_no_load_test_values(assembly_no: str):
    """Return the first row of no-load test values for the given assembly.

    This is used for the small UI display (no_load_var).
    """
    tbl = _extract_no_load_table(assembly_no)
    if not tbl or len(tbl) < 2:
        return None

    # Take the first data row after the header row.
    # Attempt to map columns based on the expected format.
    header = [c.strip().upper() for c in tbl[0]]
    data = tbl[1]

    def _col(val):
        try:
            return header.index(val)
        except ValueError:
            return None

    volts = data[_col("VOLTS")] if _col("VOLTS") is not None and _col("VOLTS") < len(data) else ""
    amp = data[_col("AMP")] if _col("AMP") is not None and _col("AMP") < len(data) else ""
    watt = data[_col("WATT")] if _col("WATT") is not None and _col("WATT") < len(data) else ""

    return volts, amp, watt

def load_assembly_numbers():
    """Fetch all unique assembly numbers for the selected motor model."""
    model = motor_var.get().strip().upper()
    if not model:
        print("Please enter the Motor Model Number first")
        return
    
    try:
        data_df = pd.read_excel(data_file, header=None)
    except Exception as e:
        print("Error loading file:", e)
        return
    
    # Extract base model number (before any space)
    base_model = model.split()[0] if ' ' in model else model
    
    # Search in Column 2 for test_data (use base model number only)
    data_match = data_df[data_df.iloc[:, 2].astype(str).str.contains(base_model, case=False, na=False)]
    
    if data_match.empty:
        print(f"✗ Model code '{base_model}' not found in test_data")
        assembly_combo['values'] = []
        return
    
    # Extract unique assembly numbers from Column 3 and normalize them.
    assembly_numbers = sorted({
        _normalize_assembly_id(x)
        for x in data_match.iloc[:, 3].tolist()
        if _normalize_assembly_id(x)
    })
    assembly_combo['values'] = assembly_numbers
    
    print(f"Found {len(assembly_numbers)} assembly numbers for model '{base_model}'")
    if assembly_numbers:
        assembly_combo.current(0)  # Select first one by default


def load_data():
    """Load motor data based on Model Number and selected Assembly Number.
    
    - TEST_PARAMETER.xlsx: Column 0 has Model number, contains name plate data
    - test_data.xlsx: Column 2 has Model code, Column 3 has Assembly number
    """
    model = motor_var.get().strip()
    assembly_no = assembly_combo.get().strip()

    # computed values from data_file
    global ambient_temp_val, resistance_cold_val, resistance_20_val
    
    if not model:
        print("Please enter the Motor Model Number")
        return
    
    if not assembly_no:
        print("Please select an Assembly Number")
        return

    try:
        param_df = pd.read_excel(param_file, header=None)
        data_df = pd.read_excel(data_file, header=None)
    except Exception as e:
        print("Error loading files:", e)
        return

    # Extract base model number (before any space)
    base_model = model.split()[0] if ' ' in model else model
    
    # Search in Column 0 for TEST_PARAMETER (use full model name)
    param_match = param_df[param_df.iloc[:, 0].astype(str).str.contains(model, case=False, na=False)]
    
    # Search in Column 2 for test_data (use base model number only)
    data_match = data_df[data_df.iloc[:, 2].astype(str).str.contains(base_model, case=False, na=False)]
    
    # Further filter by Assembly Number (Column 3)
    if not data_match.empty:
        # Keep a copy so we can report what assembly numbers were available.
        pre_filter_match = data_match

        assembly_no_norm = _normalize_assembly_id(assembly_no)
        data_match = data_match[data_match.iloc[:, 3].apply(_normalize_assembly_id) == assembly_no_norm]

        if data_match.empty:
            # For debugging: show which assembly numbers were found when the selected one doesn't match.
            found_asm = sorted({
                _normalize_assembly_id(x)
                for x in pre_filter_match.iloc[:, 3].tolist()
                if _normalize_assembly_id(x)
            })
            print(
                f"✗ No exact match for assembly '{assembly_no}' (normalized '{assembly_no_norm}').",
                f"Found assemblies: {found_asm}"
            )

    global plate_rows
    plate_rows = []
    
    if not param_match.empty:
        param_row = param_match.iloc[0]
        plate_rows.append(("Frame", param_row.iloc[1]))
        plate_rows.append(("Power (kW)", param_row.iloc[2]))
        plate_rows.append(("Voltage", param_row.iloc[3]))
        plate_rows.append(("Frequency", param_row.iloc[4]))
        plate_rows.append(("Speed", param_row.iloc[5]))
        print(f"Parameters found for '{model}'")
    else:
        print(f"✗ Model '{model}' not found in TEST_PARAMETER")

    if not data_match.empty:
        data_row = data_match.iloc[0]  # Get the matching row
        row_idx = data_match.index[0]
        print(f"Using row {row_idx} for assembly {assembly_no} (model {model})")
        print("Raw cols: G/I/K/M ->", 
              data_row.iloc[6] if len(data_row)>6 else None,
              data_row.iloc[8] if len(data_row)>8 else None,
              data_row.iloc[10] if len(data_row)>10 else None,
              data_row.iloc[12] if len(data_row)>12 else None)

        date_var.set(str(data_row.iloc[0]))

        # Load no-load test values from the external sheet (Book1.xlsx) based on assembly number
        no_load_values = _get_no_load_test_values(assembly_no)
        if no_load_values:
            volts, amp, watt = no_load_values
            no_load_var.set(f"{volts} V / {amp} A / {watt} W")
        else:
            no_load_var.set(str(data_row.iloc[26] if len(data_row) > 26 else "N/A"))

        locked_var.set(str(data_row.iloc[30] if len(data_row) > 30 else "N/A"))

        # === Compute resistance & ambient temperature from test_data.xlsx ===
        # Ambient temperature is in column G (index 6)
        ambient_temp_val = _parse_float(data_row.iloc[6] if len(data_row) > 6 else None)

        # Resistance per phase values are in columns I (9), K (11), M (13)
        resist_vals = []
        for idx in (8, 10, 12):
            if len(data_row) > idx:
                v = _parse_float(data_row.iloc[idx], None)
                if v is not None:
                    resist_vals.append(v)
        # compute average (ignore missing values)
        resistance_cold_val = sum(resist_vals) / len(resist_vals) if resist_vals else 0.0

        # Temperature correction (Copper default)
        alpha = 0.00393
        resistance_20_val = resistance_cold_val / (1 + alpha * (ambient_temp_val - 20)) if resistance_cold_val else 0.0

        # Update UI fields so user can see these values
        res_cold_var.set(f"{round(resistance_cold_val, 3)}")
        res_20deg_var.set(f"{round(resistance_20_val, 3)}")

        print(f"Test data found for Assembly {assembly_no} - Date: {date_var.get()} - Ambient {round(ambient_temp_val,1)}°C")
    else:
        print(f"✗ Data not found for model '{base_model}' with assembly '{assembly_no}'")

def extract_table_for_assembly(df, asm):
    """Return the vibration table rows for the given assembly from *df*."""
    asm_str = str(asm or "").strip()

    # Prefer a numeric assembly identifier when available.
    # We'll match against the sheet's "ASSEMBLY NO" rows by finding a
    # digit substring that appears in both the request and the sheet.
    groups = re.findall(r"\d+", asm_str)
    candidates = set(groups)

    # Also include a suffix match for the most common assembly digit length seen
    # in the sheet (e.g. if sheet uses 4-digit assembly IDs, allow matching the
    # last 4 digits of a longer input like MOTOR2501642 -> 1642).
    found_ids = []
    for _, r in df.iterrows():
        joined = " ".join(str(x) if not pd.isna(x) else "" for x in r.tolist())
        m = re.search(r"ASSEMBLY\s*NO[-\s]*(\d+)", joined, flags=re.I)
        if m:
            found_ids.append(m.group(1))

    if found_ids:
        # Determine the most common digit length among available assembly IDs.
        lengths = [len(x) for x in found_ids if x.isdigit()]
        if lengths:
            most_common_len = max(set(lengths), key=lengths.count)
            for g in groups:
                if len(g) > most_common_len:
                    candidates.add(g[-most_common_len:])

    rows = []
    capturing = False
    start_col = None
    end_col = None

    def _matches_candidate(text: str) -> bool:
        if not candidates:
            return True
        text = str(text).upper()
        return any(c.upper() in text for c in candidates)

    for _, r in df.iterrows():
        text_cells = [str(x) if not pd.isna(x) else "" for x in r.tolist()]
        joined = " ".join(text_cells).upper()

        if capturing:
            # stop when a new assembly header starts in the same column
            # and it does not match the requested digits (if any)
            if start_col is not None and start_col < len(text_cells):
                colval = str(text_cells[start_col]).upper()
                if re.match(r"^\s*ASSEMBLY\s*NO", colval):
                    if candidates:
                        if not _matches_candidate(colval) and rows:
                            break
                    else:
                        # if no digits were requested, stop at the next assembly header
                        if rows:
                            break
            if all(not cell.strip() for cell in text_cells):
                continue
            rows.append(text_cells)
        else:
            # start capturing when we find the correct assembly header
            if "ASSEMBLY NO" in joined and _matches_candidate(joined):
                capturing = True
                for idx, cell in enumerate(text_cells):
                    if _matches_candidate(cell):
                        start_col = idx
                        break
                # if no digit was found, just start at the first "ASSEMBLY NO" column
                if start_col is None:
                    for idx, cell in enumerate(text_cells):
                        if re.match(r"^\s*ASSEMBLY\s*NO", str(cell), flags=re.I):
                            start_col = idx
                            break
                if start_col is None:
                    start_col = 0
                end_col = start_col + 1
                while end_col < len(text_cells) and \
                      not re.search(r"ASSEMBLY\s*NO", text_cells[end_col], flags=re.I):
                    end_col += 1
                rows.append(text_cells)

    if rows and start_col is not None:
        sliced = []
        for r in rows:
            segment = r[start_col:end_col]
            i = len(segment)
            while i > 0 and not str(segment[i-1]).strip():
                i -= 1
            sliced.append(segment[:i])
        rows = sliced
    return rows


def generate_pdf(output_path=None, open_pdf=True):
    try:
        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, HRFlowable, Image
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        import os

        
        model = motor_var.get().strip()
        assembly_no = assembly_combo.get().strip()
        
        if not model or not assembly_no:
            print("Please enter Motor Model and select Assembly Number")
            return
        
        base_model = model.split()[0] if ' ' in model else model
        
        # Load data
        param_df = pd.read_excel(param_file, header=None)
        data_df = pd.read_excel(data_file, header=None)
        
        param_match = param_df[param_df.iloc[:, 0].astype(str).str.contains(model, case=False, na=False)]
        data_match = data_df[data_df.iloc[:, 2].astype(str).str.contains(base_model, case=False, na=False)]
        
        # Further filter by Assembly Number (Column 3)
        if not data_match.empty:
            data_match = data_match[data_match.iloc[:, 3].astype(str) == assembly_no]
        
        if param_match.empty or data_match.empty:
            print("ERROR: Data not found for model", model, "with assembly", assembly_no)
            return
        
        param_row = param_match.iloc[0]
        data_row = data_match.iloc[0]
        row_idx = data_match.index[0]
        print(f"Generating PDF using row {row_idx} for assembly {assembly_no} (model {model})")
        print("Raw cols: G/I/K/M ->", 
              data_row.iloc[6] if len(data_row)>6 else None,
              data_row.iloc[8] if len(data_row)>8 else None,
              data_row.iloc[10] if len(data_row)>10 else None,
              data_row.iloc[12] if len(data_row)>12 else None)

        # === Compute values from test_data.xlsx ===
        ambient_temp = _parse_float(data_row.iloc[6] if len(data_row) > 6 else None)

        # Resistance per phase values are in columns I (8), K (10), M (12)
        res_vals = []
        for idx in (8, 10, 12):
            if len(data_row) > idx:
                v = _parse_float(data_row.iloc[idx], None)
                if v is not None:
                    res_vals.append(v)

        # If any values exist, average them; else leave as 0.0
        resistance_cold = sum(res_vals) / len(res_vals) if res_vals else 0.0

        # Temperature correction (Copper)
        alpha = 0.00393
        resistance_20 = resistance_cold / (1 + alpha * (ambient_temp - 20)) if resistance_cold else 0.0

        ambient_temp = round(ambient_temp, 1)
        resistance_cold = round(resistance_cold, 3)
        resistance_20 = round(resistance_20, 3)

        # Sync UI values so they reflect the data_file values
        res_cold_var.set(str(resistance_cold))
        res_20deg_var.set(str(resistance_20))

        # Format raw values for display in the PDF (and debug output)
        raw_res_str = ", ".join(f"{v:.3f}" for v in res_vals) if res_vals else "n/a"
        if res_vals:
            print(f"Raw resistance values (I/K/M): {raw_res_str} -> avg {resistance_cold}")
        else:
            print("Raw resistance values (I/K/M): <none>")

        # Save to Desktop
        desktop = os.path.expanduser("~\\Desktop")
        os.makedirs(desktop, exist_ok=True)
        pdf_path = output_path or os.path.join(desktop, "certificate.pdf")
        
        # Use larger margins to reserve space for the page header.
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=(8.27*inch, 11.69*inch),
            topMargin=1.25*inch,
            bottomMargin=0.75*inch,
            leftMargin=0.4*inch,
            rightMargin=0.75*inch
        )
        styles = getSampleStyleSheet()
        styles['Heading3'].fontSize = 11
        styles['Heading3'].leading = 13
        styles['Heading3'].spaceBefore = 2
        styles['Heading3'].spaceAfter = 3
        styles['Heading4'].fontSize = 10
        styles['Heading4'].leading = 12
        styles['Heading4'].spaceBefore = 2
        styles['Heading4'].spaceAfter = 3
        
        # Custom style for centered title
        title_style = ParagraphStyle(name='CustomTitle', parent=styles['Heading1'],
                                     alignment=1, fontSize=16, textColor=colors.black)

        # Custom styles for Visual Inspection section
        inspection_heading_style = ParagraphStyle(
            name='InspectionHeading',
            parent=styles['Heading3'],
            fontSize=11,
            leading=14,
            spaceBefore=4,
            spaceAfter=2,
        )

        inspection_body_style = ParagraphStyle(
            name='InspectionBody',
            parent=styles['Normal'],
            fontSize=9,
            leading=11,
            leftIndent=8,
            spaceBefore=2,
            spaceAfter=4,
        )
        table_bold_small_style = ParagraphStyle(
            name='TableBoldSmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        locked_bold_small_style = ParagraphStyle(
            name='LockedBoldSmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        dielectric_bold_small_style = ParagraphStyle(
            name='DielectricBoldSmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        dielectric_body_small_style = ParagraphStyle(
            name='DielectricBodySmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        vibration_bold_small_style = ParagraphStyle(
            name='VibrationBoldSmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        vibration_body_small_style = ParagraphStyle(
            name='VibrationBodySmall',
            parent=styles['Normal'],
            fontSize=7,
            leading=8,
        )
        bottom_bold_small_style = ParagraphStyle(
            name='BottomBoldSmall',
            parent=styles['Normal'],
            fontSize=6.5,
            leading=7.5,
        )
        bottom_body_small_style = ParagraphStyle(
            name='BottomBodySmall',
            parent=styles['Normal'],
            fontSize=6.5,
            leading=7.5,
        )
        
        # ===== LOGO / STAMP SETTINGS =====
        # Place 'logo.png' and 'stamp.png' next to this script (same folder), or bundle via PyInstaller.
        def _find_resource(filename: str) -> str:
            """Locate a bundled resource (PyInstaller, exe folder, script folder, cwd)."""
            candidates = []

            # PyInstaller --onefile or --onedir
            if hasattr(sys, "_MEIPASS"):
                candidates.append(sys._MEIPASS)

            # When frozen, look next to the executable
            if getattr(sys, "frozen", False):
                candidates.append(os.path.dirname(sys.executable))

            # Running as a script
            candidates.append(os.path.dirname(os.path.abspath(__file__)))

            # Running from a working directory (e.g., dist folder)
            candidates.append(os.getcwd())

            for base in candidates:
                if not base:
                    continue
                candidate = os.path.join(base, filename)
                if os.path.exists(candidate):
                    return candidate

            # Fallback (will likely fail later if resource is missing)
            print(f"WARNING: resource {filename!r} not found in any expected location: {candidates}")
            return filename

        logo_path = _find_resource("logo.png")
        stamp_path = _find_resource("stamp.png")

        def _fit_image_dims(img, maxw, maxh):
            iw, ih = img.getSize()
            scale = min(maxw / iw, maxh / ih)
            return iw * scale, ih * scale

        def draw_header(canvas, doc):
            def _log(msg: str):
                try:
                    log_path = os.path.join(os.path.expanduser("~"), "Desktop", "telema_resource_log.txt")
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(f"{datetime.datetime.now().isoformat()} {msg}\n")
                except Exception:
                    pass

            try:
                from reportlab.lib.utils import ImageReader
                page_top = doc.pagesize[1]
                header_left = doc.leftMargin
                header_right = doc.pagesize[0] - doc.rightMargin

                logo_width = 0
                logo_gap = 0.18 * inch
                logo_x = header_left
                logo_y = page_top - 0.92 * inch

                # Draw logo on all pages (top-left)
                if os.path.exists(logo_path):
                    img = ImageReader(logo_path)
                    width, height = _fit_image_dims(img, 1.55*inch, 0.7*inch)
                    canvas.drawImage(img, logo_x, logo_y, width=width, height=height, preserveAspectRatio=True, mask='auto')
                    logo_width = width
                    _log(f"logo drawn at ({logo_x:.1f},{logo_y:.1f}) size {width:.1f}x{height:.1f}")

                canvas.saveState()
                text_x = header_left + logo_width + (logo_gap if logo_width else 0)

                canvas.setFont("Helvetica-Bold", 17)
                canvas.drawString(text_x, page_top - 0.46*inch, "Routine Test Report")

                canvas.setFont("Helvetica", 10)
                canvas.drawString(text_x, page_top - 0.67*inch, plain_motor_spec)

                canvas.setFont("Helvetica", 9)
                canvas.drawString(
                    text_x,
                    page_top - 0.86*inch,
                    "Ref : RDSO : E - 10/3/09 / IS 12615 / EN 60034 / Customer Specifications"
                )

                # Separator line
                line_y = page_top - 0.97*inch
                canvas.setLineWidth(0.8)
                canvas.line(header_left, line_y, header_right, line_y)

                # Company info below separator
                canvas.setFont("Helvetica", 10)
                company_text = "Power Drives (Guj) Pvt. Ltd., Vadodara - 390 010."
                canvas.drawString(header_left, page_top - 1.11*inch, company_text)

                canvas.restoreState()

                # Draw stamp on all pages (bottom-right, slightly inset)
                if os.path.exists(stamp_path):
                    img = ImageReader(stamp_path)
                    width, height = _fit_image_dims(img, 1.8*inch, 1.2*inch)
                    x_offset = 2 * inch  # move stamp left from right margin
                    y_offset = 0.1 * inch  # move stamp up from bottom margin
                    x = doc.pagesize[0] - doc.rightMargin - width - x_offset
                    y = doc.bottomMargin + y_offset
                    canvas.drawImage(img, x, y, width=width, height=height, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                # Don't interrupt PDF generation on image errors
                print("WARNING: failed to draw header/stamp:", e)

        elements = []

        # ===== HEADER helper =====
        # Prepare the motor spec string used by the header drawing function
        kw = str(param_row.iloc[2]) if len(param_row) > 2 else ""
        hp = str(param_row.iloc[3]) if len(param_row) > 3 else ""
        voltage = str(param_row.iloc[5]) if len(param_row) > 5 else ""
        freq = str(param_row.iloc[7]) if len(param_row) > 7 else ""
        plain_motor_spec = f"21.5kw-31.0kw / 2 pole , {voltage}V, {freq}Hz, 3 Phase induction motor"
        motor_spec = f"<b>{plain_motor_spec}</b>"

        elements.append(Spacer(1, 0.08*inch))
        
        def _build_motor_line():
            t = Table(
                [[
                    Paragraph(f"<b>Motor Sr. No:</b> {motor_sr_var.get()}", styles['Normal']),
                    Paragraph(f"<u><b>Date:</b></u> {date_var.get()}", styles['Normal'])
                ]],
                colWidths=[3.8*inch, 2.2*inch]
            )
            t.setStyle([
                ('ALIGN', (0,0), (0,0), 'LEFT'),
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ])
            t.hAlign = 'LEFT'
            return t

        # Motor details row below the header separator
        motor_line = _build_motor_line()

        elements.append(motor_line)
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("<u><b>Name Plate Data</b></u>", styles['Normal']))
        elements.append(Spacer(1, 0.05*inch))
        
        # ===== NAME PLATE DATA TABLE =====
        name_plate = [
            ["KW", str(param_row.iloc[2]) if len(param_row) > 2 else "", 
             "VOLTS±10%", str(param_row.iloc[5]) if len(param_row) > 5 else "", 
             "PH", "3", 
             "RPM", str(param_row.iloc[6]) if len(param_row) > 6 else ""],
            ["HP", str(param_row.iloc[3]) if len(param_row) > 3 else "", 
             "Hz±5%", str(param_row.iloc[7]) if len(param_row) > 7 else "", 
             "DUTY", str(param_row.iloc[14]) if len(param_row) > 14 else"", 
             "AMP", str(param_row.iloc[4]) if len(param_row) > 4 else""],
            ["η %", str(param_row.iloc[8]) if len(param_row) > 8 else "", 
             "CONN", str(param_row.iloc[11]) if len(param_row) > 11 else"", 
             "", "", 
             "INS.CLS", str(param_row.iloc[15]) if len(param_row) > 15 else""],
            ["COS∅", str(param_row.iloc[10]) if len(param_row) > 10 else "", 
             "ENCL", str(param_row.iloc[17]) if len(param_row) > 17 else "", 
             "", "", "IP", "65"]
        ]
        t_plate = Table(
            name_plate,
            colWidths=[0.7*inch, 0.9*inch, 0.9*inch, 0.9*inch, 0.7*inch, 0.9*inch, 0.7*inch, 0.9*inch],
            rowHeights=[0.22*inch] * len(name_plate)
        )
        t_plate.hAlign = 'LEFT'
        t_plate.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ])
        elements.append(t_plate)
        elements.append(Spacer(1, 0.01*inch))
        
        # ===== VISUAL INSPECTION =====
<<<<<<< HEAD
        elements.append(Paragraph("<u><b>Visual Inspection</b></u>", inspection_heading_style))
        elements.append(Paragraph(
            "1) Motor Paint found accepted<br/>"
            "2) High Speed & Low Speed sleeves red and blue respectively with U,V,W ferruls found ok<br/>"
            "3) Fan secured with key, circlip & wash<br/>"
            "4) Cable Gland provision on left side looking from DE side<br/>"
            "5) Min DFT found on overall motor 135 micron<br/>"
            "6) Flange orientation off center wrt terminal box<br/>"
            "7) Earthing tapped hole provided inside terminal box and outer frame",
            inspection_body_style
        ))
        elements.append(Spacer(1, 0.08*inch))
=======
        elements.append(Paragraph("<u><b>Visual Inspection</b></u>", styles['Heading3']))
        elements.append(Paragraph("1) Motor Painting &amp; casting finish found accepted", styles['Normal']))
        elements.append(Paragraph("2) Direction of Rotation found clockwise for R.Y. B .", styles['Normal']))
        elements.append(Paragraph("3) 'V' ring Provided.", styles['Normal']))
        elements.append(Paragraph("4) Vibration Velocity found <= 0.9 mm/sec( With Half Key ) in suspended condition.", styles['Normal']))
        elements.append(Spacer(1, 0.05*inch))
        
        # Dimensions table - use manual entries
        dim_table = [
            ["Dimension", "Shaft Diameter", "A - Distance", "B - Distance", "Mounting hole diameter\n at foot", "Total Length", "Pcd", "Mounting hole Diameter \nat flange"],
            ["Tolerance", "+0.010 / -0.00", "+0.1 / -0.1", "+0.1 / -0.1", "+0.1 / -0.1", "+3.0 / -3.0", "+0.5 / -0.5" , "+0.05 / -0.00"],
            ["Actual dimension", shaft_dia_var.get(), a_dist_var.get(), b_dist_var.get(), mount_hole_var.get(), total_length_var.get(), pcd_var.get() , flange_var.get()], 
            ["Results", "Accepted", "Accepted", "Accepted", "Accepted", "Accepted", "Accepted" , "Accepted"]
        ]
        t_dim = Table(dim_table, colWidths=[0.9*inch, 0.9*inch, 0.8*inch, 0.8*inch, 1.5*inch, 0.8*inch, 0.6*inch , 1.5*inch])     
        t_dim.setStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black),
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('ALIGN',(0,0),(-1,-1),'CENTER'),
                        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                        ('FONTSIZE',(0,0),(-1,-1),8)])
        elements.append(t_dim)
        elements.append(Paragraph("Note : All dimensions are in mm.", styles['Normal']))
        elements.append(Spacer(1, 0.1*inch))
        
        # ===== DIELECTRIC TEST =====
        # ===== DIELECTRIC TEST =====
        elements.append(Spacer(1,0.05*inch))
>>>>>>> 2027a469b8d87a48931d90e21b1adacc2f3d61a2

        elements.append(Paragraph("<u><b>0.5 Kw/ 4Pole - Low Speed Test Results of Duel Speed motor</b></u>", styles['Heading4']))
        # Resistance Test table (like screenshot)
        res_mode = resistance_mode_var.get()
        conn_type = connection_type_var.get()

        # Raw resistance values (from test_data.xlsx)
        u_raw = _parse_float(data_row.iloc[8] if len(data_row) > 8 else None, None)
        v_raw = _parse_float(data_row.iloc[10] if len(data_row) > 10 else None, None)
        w_raw = _parse_float(data_row.iloc[12] if len(data_row) > 12 else None, None)

        def _safe_round(val):
            return round(val, 3) if val is not None else ""

        def _to_20deg(val):
            if val in ("", None):
                return ""
            return _safe_round(val / (1 + alpha * (ambient_temp - 20)))

        # Determine displayed line/phase values based on selection
        if res_mode == "Line":
            # Raw readings are assumed to be line resistances
            line_u, line_v, line_w = u_raw, v_raw, w_raw

            if conn_type == "Star":
                # In Star connection, phase resistance is typically half of line resistance
                line_u = _safe_round(u_raw / 2) if u_raw is not None else ""
                line_v = _safe_round(v_raw / 2) if v_raw is not None else ""
                line_w = _safe_round(w_raw / 2) if w_raw is not None else ""
                phase_u = line_u
                phase_v = line_v
                phase_w = line_w
                conn_desc = "Line resistance (Star)"

                # 20°C correction for line values is derived from phase correction
                line_resistance_20 = _safe_round(resistance_20 / 2)
            else:
                # TODO: apply delta conversion formula here once provided.
                # For now, use raw values as placeholder.
                line_u = _safe_round(u_raw * 3) if u_raw is not None else ""
                phase_u = line_u
                line_v = _safe_round(v_raw * 3) if v_raw is not None else ""
                phase_v = line_v
                line_w = _safe_round(w_raw * 3) if w_raw is not None else ""
                phase_w = line_w
                conn_desc = "Line resistance (Delta)"

                # Placeholder: treat 20°C correction as same as phase for now
                line_resistance_20 = _to_20deg(_safe_round((line_u + line_v + line_w) / 3))

            # Use line resistance as the displayed column values in Line mode
            display_u, display_v, display_w = line_u, line_v, line_w
            display_20 = line_resistance_20
            display_label = "Resistance per Line (cold)"

        else:
            # Phase mode: raw values are phase resistance
            phase_u = _safe_round(u_raw / 2) if u_raw is not None else ""
            phase_v = _safe_round(v_raw / 2) if v_raw is not None else ""
            phase_w = _safe_round(w_raw / 2) if w_raw is not None else ""
            conn_desc = "Phase resistance"

            # Convert to line (assuming star connection) for completeness
            line_u = _safe_round(u_raw / 2) if u_raw is not None else ""
            line_v = _safe_round(v_raw / 2) if v_raw is not None else ""
            line_w = _safe_round(w_raw / 2) if w_raw is not None else ""

            display_u, display_v, display_w = phase_u, phase_v, phase_w
            display_20 = _to_20deg(_safe_round((phase_u + phase_v + phase_w) / 3))
            display_label = "Resistance per Phase (cold)"

        def _avg(vals):
            nums = [v for v in vals if isinstance(v, (int, float))]
            return _safe_round(sum(nums) / len(nums)) if nums else ""

        line_avg = _avg([line_u, line_v, line_w])
        phase_avg = _avg([phase_u, phase_v, phase_w])
        display_u_20 = _to_20deg(phase_u)
        display_v_20 = _to_20deg(phase_v)
        display_w_20 = _to_20deg(phase_w)
        display_avg = _avg([display_u, display_v, display_w])
        display_avg_20 = _to_20deg(phase_avg)

        # Build a single-table Resistance Test layout matching the screenshot.
        # This uses fixed column widths so the measurement text takes the full middle column.

        resist_data = [
            [
                Paragraph("<b>Resistance Test</b>", styles['Normal']),
                Paragraph("<b>RDSO E-10/3/09 - 19.7.1</b>", styles['Normal']),
                Paragraph("Measurement of resistance (cold). The\n resistance of each phase winding of the\n stator, when cold, shall be measured\n either by bridge or by voltage drop \nmethod", styles['Normal']),
                Paragraph(f"<b>Ambient Temperature {round(ambient_temp,1)} Deg C</b>", styles['Normal']),
                "", ""
            ],
            [
                "", "", "", Paragraph(f"<b>{display_label}</b>", styles['Normal']), "", ""
            ],
            [
                "", "", "", Paragraph("<b>Phase</b>", styles['Normal']), Paragraph("<b> Phase \nResistance</b>", styles['Normal']), Paragraph("<b>Resistance at 20 Deg C</b>", styles['Normal'])
            ],
            ["", "", "", "U", display_u, display_u_20],
            ["", "", "", "V", display_v, display_v_20],
            ["", "", "", "W", display_w, display_w_20],
            ["", "", "", "Avg", display_avg, display_avg_20],
            [
                Paragraph(f"Resistance per phase of Type Tested motor 6.54 Ohms at 20 Deg C, should be ± 5%", styles['Normal']),
                "", "", "", "", Paragraph("- Accepted", styles['Normal'])
            ]
        ]

        # Make row 3 (measurement text row) slightly taller to allow wrapping,
        # but keep the table compact so the No Load test can remain on page 1.
        row_heights = [0.22*inch] * len(resist_data)
        if len(row_heights) > 2:
            row_heights[2] = 0.35*inch
        if len(row_heights) > 7:
            row_heights[7] = 0.34*inch

        t_resist = Table(
            resist_data,
            colWidths=[1*inch, 0.9*inch, 2.3*inch, 0.65*inch, 0.9*inch, 1*inch],
            rowHeights=row_heights
        )
        t_resist.hAlign = 'LEFT'
        t_resist.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('SPAN', (0,0), (0,6)),
            ('SPAN', (1,0), (1,6)),
            ('SPAN', (2,0), (2,6)),
            ('SPAN', (0,7), (2,7)),
            ('SPAN', (3,0), (5,0)),
            ('SPAN', (3,1), (5,1)),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('ALIGN', (2,0), (2,6), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
        ])
        elements.append(t_resist)
        # Give more breathing room between the Resistance table and the Direction of Rotation table
        elements.append(Spacer(1, 0.08*inch))

        # Direction of Rotation table (placed after resistance test)
        direction_data = [
            [
                Paragraph("<b>Direction of Rotation</b>", styles['Normal']),
                Paragraph("<b>RDSO E-10/3/09 - 19.7.2</b>", styles['Normal']),
                Paragraph(
                    "Direction of rotation should be same as that marked on the motor when supply phase sequence RYB is connected to U,V,W terminals",
                    styles['Normal']
                
                )
            ],
            [
                "", "", Paragraph(
                    "Motor rotates in Anticlock direction (DE) when R,Y,B phase are connected to respective terminals marked U, V, W.",
                    styles['Normal']
                )
            ]
        ]

        # Align direction table width with the resistance table width
        resist_table_width = 1*inch + 0.9*inch + 2.3*inch + 0.65*inch + 0.9*inch + 0.9*inch
        scale = resist_table_width / (1.2*inch + 1.6*inch + 2.6*inch)
        t_direction = Table(
            direction_data,
            colWidths=[1.2*inch*scale, 1.6*inch*scale, 2.6*inch*scale]
        )
        t_direction.hAlign = 'LEFT'
        t_direction.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('SPAN', (0,0), (0,1)),
            ('SPAN', (1,0), (1,1)),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('ALIGN', (0,0), (-1,-1), 'CENTER')
        ])
        elements.append(t_direction)
        elements.append(Spacer(1, 0.08*inch))

        # Insert a page break so the No Load Test begins cleanly on a new page.
        elements.append(PageBreak())
        elements.append(Spacer(1, 0.12*inch))
        elements.append(_build_motor_line())
        elements.append(Spacer(1, 4))

        # ---- No Load Test ----
        elements.append(Paragraph("<u><b>No Load Test</b></u>", styles['Heading3']))

        tbl = _extract_no_load_table(assembly_no)

        def _find_header_row(table):
            for idx, row in enumerate(table):
                joined = " ".join(str(c).strip().lower() for c in row if c is not None)
                if "%voltage" in joined or ("volts" in joined and "amp" in joined) or "watt" in joined:
                    return idx
            return None

        def _find_column_indices(header_row):
            # Desired columns in order of appearance
            desired = [
                "%voltage", "volts", "amp", "watt",
                "current ratio", "value of type tested motor", "remark"
            ]
            indices = []
            row = [str(x).strip().lower() for x in header_row]
            for key in desired:
                idx = next((i for i, v in enumerate(row) if key in v), None)
                indices.append(idx)
            return indices

        def _build_noload_table(table):
            hdr_idx = _find_header_row(table)
            if hdr_idx is None:
                return None

            header_row = table[hdr_idx]
            lower_row = [str(x).strip().lower() for x in header_row]

            def _find_col(*keys):
                for key in keys:
                    idx = next((i for i, v in enumerate(lower_row) if key in v), None)
                    if idx is not None:
                        return idx
                return None

            idx_pct = _find_col("%voltage", "% voltage")
            idx_volts = _find_col("volts", "volt")
            idx_amp = _find_col("amp", "amps")
            idx_watt = _find_col("watt", "watts")
            idx_cur_ratio = _find_col("current ratio")
            idx_type_tested = _find_col("value of type tested motor", "type tested motor")
            idx_remark = _find_col("remark")

            if None in (idx_pct, idx_volts, idx_amp, idx_watt, idx_cur_ratio):
                return None

            type_amp_idx = idx_type_tested
            type_ratio_idx = None
            if idx_type_tested is not None:
                candidate = idx_type_tested + 1
                if idx_remark is None or candidate < idx_remark:
                    type_ratio_idx = candidate

            data_start = hdr_idx + 1
            if data_start < len(table):
                maybe_subheader = " ".join(str(x).strip().lower() for x in table[data_start] if x is not None)
                if "amp" in maybe_subheader and "current ratio" in maybe_subheader:
                    data_start += 1

            data_rows = []
            for row in table[data_start:]:
                if all(not str(x).strip() for x in row):
                    continue

                def _cell(idx):
                    return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

                if not any(_cell(idx) for idx in (idx_pct, idx_volts, idx_amp, idx_watt)):
                    continue

                data_rows.append([
                    _cell(idx_pct),
                    _cell(idx_volts),
                    _cell(idx_amp),
                    _cell(idx_watt),
                    _cell(idx_cur_ratio),
                    _cell(type_amp_idx),
                    _cell(type_ratio_idx),
                    _cell(idx_remark),
                ])
                if len(data_rows) >= 3:
                    break

            if not data_rows:
                return None

            final = [
                [
                    Paragraph("<b>No Load Test</b>", table_bold_small_style),
                    Paragraph("<b>RDSO E-10/3/09 -<br/>19.7.3</b>", table_bold_small_style),
                    Paragraph(
                        "The ratio of no load\ncurrent at 457 V shall\nnot exceed the value\nachieved during type\ntest",
                        styles['Normal'],
                    ),
                    Paragraph("<b>%Voltage</b>", table_bold_small_style),
                    Paragraph("<b>Volts</b>", table_bold_small_style),
                    Paragraph("<b>AMP</b>", table_bold_small_style),
                    Paragraph("<b>WATT</b>", table_bold_small_style),
                    Paragraph("<b>Current Ratio</b>", table_bold_small_style),
                    Paragraph("<b>Values of Type Tested<br/>motor</b>", table_bold_small_style),
                    "",
                    Paragraph("<b>Remark</b>", table_bold_small_style),
                ],
                [
                    "", "", "",
                    "", "", "", "", "",
                    Paragraph("<b>Amp</b>", table_bold_small_style),
                    Paragraph("<b>Current Ratio</b>", table_bold_small_style),
                    ""
                ],
            ]

            for row_idx, dr in enumerate(data_rows):
                remark_value = dr[7] if row_idx == 0 and dr[7] else ("Accepted" if row_idx == 0 else "")
                final.append([
                    "", "", "",
                    dr[0], dr[1], dr[2], dr[3], dr[4], dr[5], dr[6], remark_value
                ])

            col_widths = [
                0.7 * inch, 0.72 * inch, 1.25 * inch,
                0.66 * inch, 0.54 * inch, 0.5 * inch, 0.57 * inch, 0.6 * inch,
                0.5 * inch, 0.7 * inch, 0.54 * inch
            ]
            row_heights = [0.28 * inch, 0.2 * inch] + [0.18 * inch] * len(data_rows)

            t = Table(final, colWidths=col_widths, rowHeights=row_heights, repeatRows=2, splitByRow=1)
            t.hAlign = 'LEFT'
            t.setStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('LEADING', (0,0), (-1,-1), 9),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('ALIGN', (2,0), (2,-1), 'LEFT'),
                ('ALIGN', (10,0), (10,-1), 'CENTER'),
                ('SPAN', (0,0), (0, len(final)-1)),
                ('SPAN', (1,0), (1, len(final)-1)),
                ('SPAN', (2,0), (2, len(final)-1)),
                ('SPAN', (3,0), (3,1)),
                ('SPAN', (4,0), (4,1)),
                ('SPAN', (5,0), (5,1)),
                ('SPAN', (6,0), (6,1)),
                ('SPAN', (7,0), (7,1)),
                ('SPAN', (8,0), (9,0)),
                ('SPAN', (10,0), (10,1)),
                ('SPAN', (10,2), (10, len(final)-1)),
            ])
            return t

        t_noload = _build_noload_table(tbl) if tbl else None
        if not t_noload:
            t_noload = Table([
                ["No Load Test", "(no data found for this assembly)"]
            ], colWidths=[2*inch, 4*inch])
            t_noload.hAlign = 'LEFT'
            t_noload.setStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ])

        elements.append(t_noload)
        elements.append(Spacer(1, 0.05*inch))
        
        # Locked Rotor Test (formatted to match screenshot layout)
        elements.append(Paragraph("<u><b>Locked Rotor Test</b></u>", styles['Heading3']))

        locked = [
            [
                Paragraph("<b>Locked Rotor Test</b>", locked_bold_small_style),
                Paragraph("<b>RDSO E-10/3/09 - 19.7.4</b>", locked_bold_small_style),
                Paragraph("A single test at any reduced balanced input voltage as per IS 4029-1991", styles['Normal']),
                Paragraph("<b>Volts</b>", locked_bold_small_style),
                Paragraph("<b>Amps</b>", locked_bold_small_style),
                Paragraph("<b>Watts</b>", locked_bold_small_style),
                Paragraph("<b>Kg.1 mtr</b>", locked_bold_small_style),
                Paragraph("<b>Declared Value</b>", locked_bold_small_style),
                Paragraph("<b>Remark</b>", locked_bold_small_style)
            ],
            [
                "", "", "",
                str(data_row.iloc[56] if len(data_row) > 56 else ""),
                str(data_row.iloc[61] if len(data_row) > 61 else ""),
                str(data_row.iloc[63] if len(data_row) > 63 else ""),
                str(data_row.iloc[66] if len(data_row) > 66 else ""),
                
            ],
             [
        "", "", "",
        Paragraph("<b>% starting torque</b>", locked_bold_small_style),
        "",
        "210%","",                # 👈 your value
        "200 % minimum",
        "Accepted"
    ]
        ]

        t_locked = Table(locked, colWidths=[1*inch, 0.9*inch, 1.9*inch, 0.6*inch, 0.6*inch, 0.55*inch, 0.5*inch, 0.9*inch, 0.8*inch])
        t_locked.hAlign = 'LEFT'
        t_locked.setStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),

        # Left side vertical merge
        ('SPAN', (0,0), (0,2)),  # Locked Rotor Test
        ('SPAN', (1,0), (1,2)),  # RDSO
        ('SPAN', (2,0), (2,2)),  # Description

        # Right side vertical merge
        ('SPAN', (7,0), (7,1)),  # Declared Value ✅ FIXED
        ('SPAN', (8,0), (8,1)),  # Remark

        # % starting torque row merge (middle columns)
        ('SPAN', (3,2), (4,2)),

        # Alignment
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),

        ('FONTSIZE', (0,0), (-1,-1), 8),
    ])
        elements.append(t_locked)
        elements.append(Spacer(1, 0.2*inch))

        # ===== DIELECTRIC + INSULATION TABLE =====

        dielec_table = [
            [
                Paragraph("<b>Dielectric (High Voltage) Test & Insulation Resistance Test</b>", dielectric_bold_small_style),
                Paragraph("<b>RDSO E-10/3/09 - 19.7.6</b><br/>On hot winding's, perform the electrical test at 3300V, 50 Hz applied for one minutes between each insulated stator winding and the frame of the motor and record the insulation resistance before and after the high voltage tests in terms of IS : 325-1991. There should be no appreciable difference in insulation resistance values.", dielectric_body_small_style),
                Paragraph("<b>Tests</b>", dielectric_bold_small_style),
                Paragraph("<b>Required</b>", dielectric_bold_small_style),
                Paragraph("<b>Result</b>", dielectric_bold_small_style),
                Paragraph("<b>Remark</b>", dielectric_bold_small_style)
            ],
            [
                "", "",
                "Dielectric Test 2.64kV at 1 min (HV)",
                "Withstand",
                "Withstood",
                "Accepted"
            ],
            [
                "", "",
                "Insulation Resistance test (500v DC) before HV",
                "≥ 200 M Ohm",
                "677 M Ohm",
                "Accepted"
            ],
            [
                "", "",
                "Insulation Resistance test (500v DC) after HV",
                "≥ 200 M Ohm",
                "645 M Ohm",
                "Accepted"
            ]
        ]
        t_dielec = Table(
            dielec_table,
            colWidths=[1.2*inch, 1.95*inch, 2.43*inch, 0.8*inch, 0.7*inch, 0.67*inch]
        )
        t_dielec.hAlign = 'LEFT'

        t_dielec.setStyle([
            ('GRID', (0,0), (-1,-1), 0.05, colors.black),

            # LEFT SIDE FULL MERGE
            ('SPAN', (0,0), (0,3)),  # Test name
            ('SPAN', (1,0), (1,3)),  # Description

            # ALIGNMENT
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),

            # LEFT ALIGN description text
            ('ALIGN', (1,0), (1,3), 'LEFT'),
            ('ALIGN', (2,1), (2,3), 'LEFT'),

            # FONT
            ('FONTSIZE', (0,0), (-1,-1), 8),

            # HEADER BACKGROUND
        ])

        # Add dielectric table after it is created
        elements.append(t_dielec)
        elements.append(Spacer(1, 0.2*inch))

        vibration_data = [
        [
            Paragraph("<b>Vibration and Noise Level Test</b>", vibration_bold_small_style),
            Paragraph("<b>RDSO E-10/3/09 - 19.7.7</b>", vibration_bold_small_style),
            Paragraph("The vibration levels on the motors shall not exceed 15/10 microns refer clause 4.9 when tested as per IS : 4729-1968.", vibration_body_small_style),
            Paragraph("<b>Vibration at Prescribed locations</b>", vibration_bold_small_style),
            "", "", "",
            Paragraph("<b>Maximum Permissible</b>", vibration_bold_small_style),
            Paragraph("<b>Result</b>", vibration_bold_small_style)
        ],
        [
            "", "", "",
            Paragraph("<b>Displacement (Micro)</b>", vibration_bold_small_style),
            "4", "4", "5",
            "15",
            "Accepted"
        ],
        [
            "", "", "",
            Paragraph("<b>Velocity in mm/sec</b>", vibration_bold_small_style),
            "0.4", "0.4", "0.5",
            "1.3",
            "Accepted"
        ]
    ]
        # Target column widths for the vibration/noise/surge tables
        vib_col_widths = [0.6*inch, 0.6*inch, 0.8*inch, 1.0*inch, 0.4*inch, 0.4*inch, 0.4*inch, 0.7*inch, 0.4*inch]
        noise_col_widths = [0.92*inch, 0.92*inch]
        surge_col_widths = [1.15*inch, 0.95*inch, 1.35*inch, 2.0*inch]
        page_content_width = doc.pagesize[0] - doc.leftMargin - doc.rightMargin
        desired_vib_width = sum(vib_col_widths)
        vib_width = page_content_width
        vib_scale = vib_width / desired_vib_width if desired_vib_width else 1.0
        vib_col_widths = [w * vib_scale for w in vib_col_widths]

        t_vibration = Table(
            vibration_data,
            colWidths=vib_col_widths,
            rowHeights=[None]*len(vibration_data)

        )
        t_vibration.hAlign = 'LEFT'
        t_vibration.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),

            # LEFT SIDE MERGE
            ('SPAN', (0,0), (0,2)),
            ('SPAN', (1,0), (1,2)),
            ('SPAN', (2,0), (2,2)),

            # HEADER MERGE (Vibration title)
            ('SPAN', (3,0), (6,0)),

            # ALIGNMENT
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),

            # LEFT ALIGN description
            ('ALIGN', (2,0), (2,2), 'LEFT'),
            ('ALIGN', (3,1), (3,2), 'LEFT'),

            # FONT
            ('FONTSIZE', (0,0), (-1,-1), 7),

            # HEADER BG
        ])

        elements.append(t_vibration)
        elements.append(Spacer(1, 0.2*inch))

        noise_data = [
            [
                Paragraph("<b>Noise Level</b>", bottom_bold_small_style),
                Paragraph("<b>Maximum Allowed</b>", bottom_bold_small_style),
            ],
            ["63.5 db", "90 db"],
            ["Accepted"]
        ]

        t_noise = Table(noise_data, colWidths=noise_col_widths, rowHeights=[0.22*inch] * len(noise_data))
        t_noise.hAlign = 'LEFT'

        t_noise.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 6.5),
            ('SPAN', (0,2), (1,2)),
        ])

        surge_data = [[
            Paragraph("<b>Surge Test without<br/>rotor in position</b>", bottom_bold_small_style),
            Paragraph("<b>RDSO E-10/3/09 -<br/>19.7.8</b>", bottom_bold_small_style),
            Paragraph("Should be conducted<br/>without rotor in<br/>position at 5 kV Pk-Pk", bottom_body_small_style),
            Paragraph("No intern turn short found -<br/><b>Accepted</b>", bottom_body_small_style),
        ]]

        t_surge = Table(
            surge_data,
            colWidths=surge_col_widths,
            rowHeights=[0.42*inch]
        )
        t_surge.hAlign = 'LEFT'
        t_surge.setStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 6.5),
            ('ALIGN', (3,0), (3,0), 'LEFT'),
        ])

        elements.append(t_noise)
        elements.append(Spacer(1, 0.2*inch))
        elements.append(t_surge)
        # before building, attempt to append a second page with the
        # assembly-specific vibration table read from the same workbook.
        # ------------------------------------------------------------------

        # Debug: try wrapping each flowable to identify which table fails due to width/height issues.
        from reportlab.platypus import Table

        max_width = doc.pagesize[0] - doc.leftMargin - doc.rightMargin
        max_height = doc.pagesize[1] - doc.topMargin - doc.bottomMargin

        def _debug_wrap(flowable, prefix=""):
            try:
                if isinstance(flowable, Table):
                    summary = f"Table({len(getattr(flowable, '_cellvalues', []))} rows)"
                    print(f"DEBUG: Wrapping {prefix}{summary}")
                    # try to wrap this table with available page width
                    flowable.wrap(max_width, max_height)
                elif hasattr(flowable, 'wrap'):
                    flowable.wrap(max_width, max_height)
            except Exception as e:
                print(f"DEBUG: Flowable failing wrap: {prefix}{type(flowable).__name__} -> {e}")
                if isinstance(flowable, Table):
                    try:
                        # dump a preview of the table
                        rows = getattr(flowable, '_cellvalues', [])
                        print(f"DEBUG: Table first row: {rows[0] if rows else None}")
                    except Exception:
                        pass
                raise

            if isinstance(flowable, Table):
                for row in getattr(flowable, '_cellvalues', []):
                    for cell in row:
                        _debug_wrap(cell, prefix + "  ")
            elif hasattr(flowable, 'flowables'):
                for f in getattr(flowable, 'flowables', []):
                    _debug_wrap(f, prefix + "  ")

        print('DEBUG: elements types:', [type(e).__name__ for e in elements])
       # for flow in elements:
        #    _debug_wrap(flow)

        # Build and save the PDF (no vibration test table included)
        doc.build(elements, onFirstPage=draw_header, onLaterPages=draw_header)
        
        print(f"PDF saved to: {pdf_path}")
        
        # Try to open it automatically
        import subprocess
        try:
            if open_pdf:
                if os.name == 'nt':
                    os.startfile(pdf_path)
                else:
                    subprocess.Popen(['xdg-open' if shutil.which('xdg-open') else 'open', pdf_path])
                print("Opening PDF...")
        except Exception:
            pass

    except Exception as e:
        import traceback
        traceback.print_exc()
        print("ERROR:", e)

# ===== GUI SETUP =====
root = Tk()
root.title("Standard Radiator report generator")
root.state('zoomed')

report_sections = {}
active_report_id = 1


def _merge_pdfs(pdf_paths, output_path):
    merger = None
    try:
        try:
            from pypdf import PdfWriter, PdfReader
        except ImportError:
            try:
                from PyPDF2 import PdfWriter, PdfReader
            except ImportError:
                import subprocess
                cmd = [
                    "py", "-c",
                    (
                        "from pypdf import PdfWriter, PdfReader; "
                        "import sys; "
                        "writer = PdfWriter(); "
                        "paths = sys.argv[1:-1]; "
                        "out = sys.argv[-1]; "
                        "[writer.add_page(page) for p in paths for page in PdfReader(p).pages]; "
                        "writer.write(out)"
                    ),
                    *pdf_paths,
                    output_path,
                ]
                subprocess.check_call(cmd)
                return

        merger = PdfWriter()
        for pdf_path in pdf_paths:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                merger.add_page(page)
        with open(output_path, "wb") as f:
            merger.write(f)
    finally:
        if merger and hasattr(merger, "close"):
            merger.close()


def _activate_report(report_id):
    global active_report_id
    global motor_var, date_var, motor_sr_var, no_load_var, locked_var
    global res_cold_var, res_20deg_var, resistance_mode_var, connection_type_var
    global assembly_combo, rb_star, rb_delta

    cfg = report_sections[report_id]
    active_report_id = report_id
    motor_var = cfg["motor_var"]
    date_var = cfg["date_var"]
    motor_sr_var = cfg["motor_sr_var"]
    no_load_var = cfg["no_load_var"]
    locked_var = cfg["locked_var"]
    res_cold_var = cfg["res_cold_var"]
    res_20deg_var = cfg["res_20deg_var"]
    resistance_mode_var = cfg["resistance_mode_var"]
    connection_type_var = cfg["connection_type_var"]
    assembly_combo = cfg["assembly_combo"]
    rb_star = cfg["rb_star"]
    rb_delta = cfg["rb_delta"]


def _run_for_report(report_id, func, *args, **kwargs):
    previous = active_report_id
    try:
        _activate_report(report_id)
        _update_connection_controls()
        return func(*args, **kwargs)
    finally:
        _activate_report(previous)
        _update_connection_controls()


def _update_connection_controls(*args):
    for cfg in report_sections.values():
        state = 'normal' if cfg["resistance_mode_var"].get() == 'Line' else 'disabled'
        cfg["rb_star"].config(state=state)
        cfg["rb_delta"].config(state=state)


def _generate_combined_pdf():
    import tempfile
    import subprocess

    required = []
    for report_id, cfg in report_sections.items():
        model = cfg["motor_var"].get().strip()
        assembly = cfg["assembly_combo"].get().strip()
        if not model or not assembly:
            required.append(str(report_id))

    if required:
        print("Please complete Motor Model and Assembly Number for report(s):", ", ".join(required))
        return

    desktop = os.path.expanduser("~\\Desktop")
    os.makedirs(desktop, exist_ok=True)
    temp_dir = tempfile.mkdtemp(prefix="telema_reports_")
    pdf_1 = os.path.join(temp_dir, "report_1.pdf")
    pdf_2 = os.path.join(temp_dir, "report_2.pdf")
    final_pdf = os.path.join(desktop, "certificate_combined.pdf")

    _run_for_report(1, generate_pdf, pdf_1, False)
    _run_for_report(2, generate_pdf, pdf_2, False)
    _merge_pdfs([pdf_1, pdf_2], final_pdf)
    print(f"Combined PDF saved to: {final_pdf}")

    try:
        if os.name == 'nt':
            os.startfile(final_pdf)
        else:
            subprocess.Popen(['xdg-open' if shutil.which('xdg-open') else 'open', final_pdf])
    except Exception:
        pass


def _build_report_section(parent, title, report_id):
    frame = LabelFrame(parent, text=title, padx=8, pady=8)
    frame.grid_columnconfigure(1, weight=1, minsize=220)
    frame.grid_columnconfigure(3, weight=1, minsize=220)

    cfg = {
        "motor_var": StringVar(master=root),
        "date_var": StringVar(master=root),
        "motor_sr_var": StringVar(master=root),
        "no_load_var": StringVar(master=root),
        "locked_var": StringVar(master=root),
        "res_cold_var": StringVar(master=root),
        "res_20deg_var": StringVar(master=root),
        "resistance_mode_var": StringVar(master=root, value="Line"),
        "connection_type_var": StringVar(master=root, value="Star"),
    }

    Label(frame, text="Motor Model No").grid(row=0, column=0, sticky="w", padx=4, pady=2)
    Entry(frame, textvariable=cfg["motor_var"], width=25).grid(row=0, column=1, sticky="ew", padx=4, pady=2)
    Button(frame, text="Fetch Assemblies", command=lambda rid=report_id: _run_for_report(rid, load_assembly_numbers)).grid(row=0, column=2, padx=4, pady=2)

    Label(frame, text="Motor Sr No").grid(row=1, column=0, sticky="w", padx=4, pady=2)
    Entry(frame, textvariable=cfg["motor_sr_var"], width=25).grid(row=1, column=1, sticky="ew", padx=4, pady=2)

    Label(frame, text="Date").grid(row=2, column=0, sticky="w", padx=4, pady=2)
    Entry(frame, textvariable=cfg["date_var"], width=20).grid(row=2, column=1, sticky="ew", padx=4, pady=2)

    Label(frame, text="Assembly No").grid(row=2, column=2, sticky="w", padx=4, pady=2)
    cfg["assembly_combo"] = ttk.Combobox(frame, width=22, state="readonly")
    cfg["assembly_combo"].grid(row=2, column=3, sticky="ew", padx=4, pady=2)

    Label(frame, text="--- RESISTANCE TEST (Manual Entry) ---", font=("Arial", 10, "bold")).grid(
        row=3, column=0, columnspan=4, sticky="w", pady=6
    )

    Label(frame, text="Resistance at Ambient Temp").grid(row=4, column=0, sticky="w", padx=4, pady=2)
    Entry(frame, textvariable=cfg["res_cold_var"], width=20, state="readonly").grid(row=4, column=1, sticky="w", padx=4, pady=2)

    Label(frame, text="Resistance at -20 Deg C").grid(row=4, column=2, sticky="w", padx=4, pady=2)
    Entry(frame, textvariable=cfg["res_20deg_var"], width=20, state="readonly").grid(row=4, column=3, sticky="w", padx=4, pady=2)

    Label(frame, text="Resistance Type").grid(row=5, column=0, sticky="w", padx=4, pady=2)
    Radiobutton(frame, text="Line", variable=cfg["resistance_mode_var"], value="Line").grid(row=5, column=1, sticky="w", padx=4, pady=2)
    Radiobutton(frame, text="Phase", variable=cfg["resistance_mode_var"], value="Phase").grid(row=5, column=2, sticky="w", padx=4, pady=2)

    Label(frame, text="Connection").grid(row=6, column=0, sticky="w", padx=4, pady=2)
    cfg["rb_star"] = Radiobutton(frame, text="Star", variable=cfg["connection_type_var"], value="Star")
    cfg["rb_star"].grid(row=6, column=1, sticky="w", padx=4, pady=2)
    cfg["rb_delta"] = Radiobutton(frame, text="Delta", variable=cfg["connection_type_var"], value="Delta")
    cfg["rb_delta"].grid(row=6, column=2, sticky="w", padx=4, pady=2)

    cfg["resistance_mode_var"].trace_add('write', _update_connection_controls)

    Button(frame, text="Load Data", command=lambda rid=report_id: _run_for_report(rid, load_data)).grid(row=7, column=0, padx=4, pady=8)

    return frame, cfg


root.grid_columnconfigure(0, weight=1)

section_1, report_sections[1] = _build_report_section(root, "Report 1", 1)
section_1.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))

section_2, report_sections[2] = _build_report_section(root, "Report 2", 2)
section_2.grid(row=1, column=0, sticky="ew", padx=10, pady=6)

footer = Frame(root)
footer.grid(row=2, column=0, sticky="w", padx=10, pady=10)
Button(footer, text="Generate PDF", command=_generate_combined_pdf).grid(row=0, column=0, padx=4)

_activate_report(1)
_update_connection_controls()

if __name__ == "__main__":
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.destroy()
        sys.exit(0)
