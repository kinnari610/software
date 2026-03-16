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
# workbook containing the vibration tables; used for page‑2 output
vibration_file = r"C:\Users\kinnari\Downloads\Telegram Desktop\TELEMA MOTOR VIBRATION TEST REPORT 10-53HZ.xlsx"

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
        dimension_var.set(assembly_no)
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


def generate_pdf():
    try:
        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, HRFlowable
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
        pdf_path = os.path.join(desktop, "certificate.pdf")
        
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=(8.27*inch, 11.69*inch),
            topMargin=20,
            bottomMargin=20,
            leftMargin=20,
            rightMargin=20
        )
        styles = getSampleStyleSheet()
        
        # Custom style for centered title
        title_style = ParagraphStyle(name='CustomTitle', parent=styles['Heading1'],
                                     alignment=1, fontSize=16, textColor=colors.black)
        
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

        def draw_logo_and_stamp(canvas, doc):
            def _log(msg: str):
                try:
                    log_path = os.path.join(os.path.expanduser("~"), "Desktop", "telema_resource_log.txt")
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(f"{datetime.datetime.now().isoformat()} {msg}\n")
                except Exception:
                    pass

            try:
                from reportlab.lib.utils import ImageReader

                # Draw logo on second page (top-right)
                if canvas.getPageNumber() == 1 and os.path.exists(logo_path):
                    img = ImageReader(logo_path)
                    width, height = _fit_image_dims(img, 2.0*inch, 0.8*inch)
                    x = doc.pagesize[0] - doc.rightMargin - width
                    y = doc.pagesize[1] - doc.topMargin - height + 6
                    canvas.drawImage(img, x, y, width=width, height=height, preserveAspectRatio=True, mask='auto')
                    _log(f"logo drawn at ({x:.1f},{y:.1f}) size {width:.1f}x{height:.1f}")

                # Draw stamp on second page (bottom-right, slightly inset)
                if canvas.getPageNumber() == 2 and os.path.exists(stamp_path):
                    img = ImageReader(stamp_path)
                    width, height = _fit_image_dims(img, 1.8*inch, 1.2*inch)
                    # offsets help place the stamp where the user circled it
                    x_offset = 2 * inch  # move stamp left from right margin
                    y_offset = 0.1 * inch  # move stamp up from bottom margin
                    x = doc.pagesize[0] - doc.rightMargin - width - x_offset
                    y = doc.bottomMargin + y_offset
                    canvas.drawImage(img, x, y, width=width, height=height, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                # Don't interrupt PDF generation on image errors
                print("WARNING: failed to draw logo/stamp:", e)

        elements = []

        
        # ===== HEADER =====
        elements.append(Paragraph("<b>Routine Test Report</b>", title_style))
        elements.append(Spacer(1, 0.1*inch))
        
        # Motor spec line
        kw = str(param_row.iloc[2]) if len(param_row) > 2 else ""
        hp = str(param_row.iloc[3]) if len(param_row) > 3 else ""
        voltage = str(param_row.iloc[5]) if len(param_row) > 5 else ""
        freq = str(param_row.iloc[7]) if len(param_row) > 7 else ""
        motor_spec = f"<b>21.5kw-31.0kw / 2 pole , {voltage}V, {freq}Hz, 3 Phase induction motor</b>"
        elements.append(Paragraph(motor_spec, styles['Normal']))
        


        
        # Company info
        elements.append(Paragraph("Ref : RDSO : E - 10/3/09 / IS 12615 / EN 60034/ Customer Specifications", styles['Normal']))

        # line ABOVE the Ref text
        elements.append(Spacer(1,5))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.black))
        elements.append(Spacer(1,5))

        elements.append(Paragraph(
        "<para alignment='right'><b>Power Drives (Guj) Pvt. Ltd., Vadodara - 390 010.</b></para>",
        styles['Normal']))
        elements.append(Spacer(1, 0.1*inch))
        
        # Motor details header
        motor_line = Table(
    [[f"Motor Sr. No: {motor_sr_var.get()}", f"Date: {date_var.get()}"]],
    colWidths=[4*inch, 2*inch]
)

        motor_line.setStyle([
        ('ALIGN',(1,0),(1,0),'RIGHT')
        ])

        elements.append(motor_line)
        elements.append(Spacer(1,10))
        elements.append(Paragraph("<u><b>Name Plate Data</b></u>", styles['Normal']))
        elements.append(Spacer(1, 0.1*inch))
        
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
        t_plate = Table(name_plate, colWidths=[0.7*inch, 0.9*inch, 0.9*inch, 0.9*inch, 
                                                0.7*inch, 0.9*inch, 0.7*inch, 0.9*inch])
        t_plate.setStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black)])
        elements.append(t_plate)
        elements.append(Spacer(1, 0.05*inch))
        
        # ===== VISUAL INSPECTION =====
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

        d_title = Paragraph("<u><b>Dielectric Test</b></u>", styles['Heading3'])

        dielec_data = [
        ["Voltage", "Time", "Result"],
        ["2.0 kV", "60 Sec", "Withstood"]
        ]

        t_dielec = Table(dielec_data, colWidths=[0.6*inch,0.6*inch,0.8*inch], hAlign='LEFT')
        t_dielec.setStyle([
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('ALIGN',(0,0),(-1,-1),'LEFT')
        ])


        # ===== INSULATION RESISTANCE =====
        i_title = Paragraph("<u><b>Insulation Resistance Test</b></u>", styles['Heading3'])

        insul_data = [
        ["Voltage (DC)", "Insulation Resistance at 60 Sec."],
        ["1000 V", ">40 G Ohm"]
        ]

        t_insul = Table(insul_data, colWidths=[1*inch,2.0*inch])
        t_insul.setStyle([
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('ALIGN',(0,0),(-1,-1),'CENTER')
        ])


        # ===== SIDE BY SIDE LAYOUT =====
        layout = Table([
        [d_title, i_title],
        [t_dielec, t_insul]
        ], colWidths=[3.15*inch,3.15*inch])

        layout.setStyle([
('LEFTPADDING',(0,0),(-1,-1),0),
('RIGHTPADDING',(0,0),(-1,-1),0),
('TOPPADDING',(0,0),(-1,-1),0),
('BOTTOMPADDING',(0,0),(-1,-1),0),

('ALIGN',(0,1),(0,1),'LEFT')  # force dielectric table to left
])

        elements.append(layout)
        elements.append(Spacer(1,0.1*inch))
        # ===== TEST RESULTS =====
        result_title = f"<u><b>21.5kw-31.0kw/ 2Pole - Test Results of motor</b></u>"
        elements.append(Paragraph(result_title, styles['Heading3']))
        
        # Resistance Test
        styles = getSampleStyleSheet()

        resistance = [
            [
                Paragraph("<b>Resistance<br/>Test</b>", styles['Normal']),
                "",
                Paragraph("Resistance per Phase at Ambient Temperature (cold)", styles['Normal']),
                Paragraph(f"{resistance_cold} Ω", styles['Normal']),
                Paragraph(f"Ambient Temp {ambient_temp} °C", styles['Normal'])
            ],
            [
                "",
                "",
                Paragraph("Resistance per Phase at 20 °C", styles['Normal']),
                Paragraph(f"{resistance_20} Ω", styles['Normal']),
                
            ]
        ]

        t_resist = Table(
            resistance,
            colWidths=[1*inch, 0.8*inch, 2.2*inch, 1.2*inch, 1.6 *inch],
            rowHeights=[0.5*inch, 0.5*inch]
        )

        t_resist.setStyle([
            ('GRID',(0,0),(-1,-1),0.5,colors.black),
            ('SPAN',(0,0),(0,1)),   # Resistance Test
            ('SPAN',(1,0),(1,1)),   # blank column
            ('SPAN',(4,0),(4,1)),   # Resistance per Phase label                
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('FONTSIZE',(0,0),(-1,-1),8)
        ])

        elements.append(t_resist)
        # No-Load Test

        elements.append(Paragraph("<u><b>No load Test</b></u>", styles['Heading3']))
        no_load = [
            ["Volts", "Amps", "Watts", "Pf", "Hz"],
            [str(data_row.iloc[37] if len(data_row) > 37 else ""), 
             str(data_row.iloc[42] if len(data_row) > 42 else ""),
             str(data_row.iloc[44] if len(data_row) > 44 else ""),
             str(data_row.iloc[45] if len(data_row) > 45 else ""),
             str(data_row.iloc[46] if len(data_row) > 46 else "")]
        ]
        t_noload = Table(no_load, colWidths=[1.2*inch]*5)
        t_noload.setStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black)])
        elements.append(t_noload)
        elements.append(Spacer(1, 0.1*inch))
        
        # Locked Rotor Test
        elements.append(Paragraph("<u><b>Locked Rotor Test</b></u>", styles['Heading3']))
        locked = [
            ["Volts", "Amps", "Watts", "N.m at 1 mtr", "% of Starting Torque"],
            [str(data_row.iloc[56] if len(data_row) > 56 else ""),
             str(data_row.iloc[61] if len(data_row) > 61 else ""),
             str(data_row.iloc[63] if len(data_row) > 63 else ""),
             str(data_row.iloc[66] if len(data_row) > 66 else ""),
             str(data_row.iloc[67] if len(data_row) > 67 else ""),],
            ["Rated Voltage =", "415", "", "", ""]
        ]
        t_locked = Table(locked, colWidths=[1.4*inch]*5)
        t_locked.setStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black)])
        elements.append(t_locked)
        elements.append(Spacer(1, 0.05*inch))        # before building, attempt to append a second page with the
        # assembly-specific vibration table read from the same workbook.
        # ------------------------------------------------------------------

        # attempt to read the sheet containing the small tables; the user
        # indicates these are stored in the vibration workbook rather than our
        # main data file.
        try:
            small_df = pd.read_excel(vibration_file, sheet_name=0, header=None)
        except Exception as exc:
            print("WARNING: could not open vibration file:", exc)
            small_df = pd.DataFrame()



        assembly_table = []
        if not small_df.empty:
            assembly_table = extract_table_for_assembly(small_df, assembly_no)

            # if extraction failed, build a list of available assembly numbers
            
        assembly_table = []
        if not small_df.empty:
            assembly_table = extract_table_for_assembly(small_df, assembly_no)

            # if extraction failed, build a list of available assembly numbers
            if not assembly_table:
                found = []
                for _, r in small_df.iterrows():
                    joined = " ".join(str(x) if not pd.isna(x) else "" for x in r.tolist())
                    m = re.search(r"ASSEMBLY\s*NO[-\s]*(\d+)", joined, flags=re.I)
                    if m:
                        found.append(m.group(1))
                found = sorted(set(found))
                print("Available vibration assemblies in sheet:", found)
                try:
                    from tkinter import messagebox
                    messagebox.showinfo("Vibration data",
                                        f"No table for {assembly_no}.\n"
                                        f"Available assemblies: {', '.join(found)}")
                except Exception:
                    pass


        if assembly_table:
            # drop any rows that are just headers or notes; these usually
            # contain the words "ASSEMBLY NO" or "OVERALL" in the first
            # column and are not part of the numeric data.
            cleaned = []
            for r in assembly_table:
                # ignore completely empty rows
                if not r or all(not str(c).strip() for c in r):
                    continue
                first = str(r[0]).strip().upper()
                # drop header markers we don't want
                if first.startswith('ASSEMBLY NO') or 'OVERALL' in first:
                    continue
                # keep frequency header row (it usually contains 'HZ')
                if not re.match(r"^\d", first) and 'HZ' not in first:
                    # any other non‑numeric row is noise
                    continue
                cleaned.append(r)
            assembly_table = cleaned

            # debug: report how many data rows we have
            print(f"Vibration rows after cleaning: {len(assembly_table)}")

            # prepended rows as requested by user (top-first order)
            header_rows = [
                ["TMB UNIT"],
                [f"Sr no : {motor_sr_var.get() or ''}"],
                ["Overall vibration measurement at non-driving end mm/s"]
            ]
            # drop any completely empty columns (they only add narrow gaps)
            if assembly_table:
                maxcols = max(len(r) for r in assembly_table)
                keep = [False] * maxcols
                for col in range(maxcols):
                    for r in assembly_table:
                        if col < len(r) and str(r[col]).strip():
                            keep[col] = True
                            break
                # rebuild rows keeping only non-empty columns
                new_rows = []
                for r in assembly_table:
                    new_rows.append([r[c] for c in range(len(r)) if keep[c]])
                assembly_table = new_rows
                cols = maxcols = max(len(r) for r in assembly_table) if assembly_table else 0
            else:
                cols = 1
            # remove any stray CU UNIT rows from the extracted data so we
            # don't end up with another copy at the bottom of the frequency
            # column.  We'll prepend a single header row below.
            assembly_table = [r for r in assembly_table if not (r and str(r[0]).strip().upper() == 'CU UNIT')]

            # insert header rows in reverse order so the list order represents top-to-bottom
            for hr in reversed(header_rows):
                if len(hr) < cols:
                    hr += [''] * (cols - len(hr))
                elif len(hr) > cols:
                    hr = hr[:cols]
                assembly_table.insert(0, hr)

            elements.append(PageBreak())
            elements.append(Paragraph(
                f"<u><b>Vibration Test </b></u>",
                styles['Heading2']))
            # build a table using column widths proportional to the
            # maximum content in each column (so every column can have a
            # different size).
            if assembly_table:
                maxcols = max(len(r) for r in assembly_table)
            else:
                maxcols = 0
            # compute width for each column based on longest string
            colwidths = []
            for col in range(maxcols):
                maxlen = 0
                for r in assembly_table:
                    if col < len(r):
                        l = len(str(r[col]))
                        if l > maxlen:
                            maxlen = l
                # fudge factor: 0.12 inch per character + 0.2 inch padding
                width_in_inches = 0.12 * maxlen + 0.2
                # convert to points and enforce minimum
                width_pts = width_in_inches * inch
                # cap frequency column to 1 inch (reduce breadth)
                if col == 0:
                    width_pts = min(width_pts, 1 * inch)
                colwidths.append(max(width_pts, 0.5 * inch))

            # if the table is too wide for the page, scale columns proportionally
            available = doc.width  # already in points
            total = sum(colwidths)
            if total > available and total > 0:
                factor = available / total
                colwidths = [w * factor for w in colwidths]

            # use a minimum row height to prevent vertical overlap
            row_heights = [0.2 * inch] * len(assembly_table)
            vib_table = Table(assembly_table, colWidths=colwidths, rowHeights=row_heights)
            # span the header rows (first few rows we prepended) across all
            # columns so their text doesn't wrap in a narrow first column.
            header_count = len(header_rows)
            style_list = [
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]
            for r in range(header_count):
                style_list.append(('SPAN', (0, r), (-1, r)))
            vib_table.setStyle(style_list)
            elements.append(vib_table)
        else:
            # no table found; still add a page so the caller knows
            elements.append(PageBreak())
            elements.append(Paragraph(
                f"<b>No vibration table found for assembly {assembly_no}</b>",
                styles['Normal']))
        
        doc.build(elements, onFirstPage=draw_logo_and_stamp, onLaterPages=draw_logo_and_stamp)
        
        print(f"PDF saved to: {pdf_path}")
        
        # Try to open it automatically
        import subprocess
        try:
            if os.name == 'nt':
                os.startfile(pdf_path)
            else:
                subprocess.Popen(['xdg-open' if shutil.which('xdg-open') else 'open', pdf_path])
            print("Opening PDF...")
        except Exception:
            pass

    except Exception as e:
        print("ERROR:", e)

# ===== GUI SETUP =====
# build a very simple form
root = Tk()
root.title("TELEMA B-35 - Routine Test Report Generator")
root.state('zoomed') 


# Define the Tkinter variables
motor_var = StringVar(master=root)
date_var = StringVar(master=root)
dimension_var = StringVar(master=root)
motor_sr_var = StringVar(master=root)
no_load_var = StringVar(master=root)
locked_var = StringVar(master=root)
#inspector_var = StringVar(master=root)

# Resistance test variables (manual entry)
res_cold_var = StringVar(master=root)
res_20deg_var = StringVar(master=root)

# Dimension variables
shaft_dia_var = StringVar(master=root)
a_dist_var = StringVar(master=root)
b_dist_var = StringVar(master=root)
mount_hole_var = StringVar(master=root)
total_length_var = StringVar(master=root)
pcd_var = StringVar(master=root)
flange_var = StringVar(master=root)

## GUI Layout - improved grid to avoid overlapping
root.columnconfigure(0, weight=1, minsize=80)
root.columnconfigure(1, weight=2, minsize=160)
root.columnconfigure(2, weight=1, minsize=80)
root.columnconfigure(3, weight=2, minsize=160)

# Motor Model
Label(root, text="Motor Model No").grid(row=0, column=0, sticky="w", padx=4, pady=2)
Entry(root, textvariable=motor_var, width=25).grid(row=0, column=1, sticky="ew", padx=4, pady=2)

Button(root, text="Fetch Assemblies", command=load_assembly_numbers).grid(row=0, column=2, padx=4, pady=2)

# Motor Serial No
Label(root, text="Motor Sr No").grid(row=1, column=0, sticky="w", padx=4, pady=2)
Entry(root, textvariable=motor_sr_var, width=25).grid(row=1, column=1, sticky="ew", padx=4, pady=2)

# Date
Label(root, text="Date").grid(row=2, column=0, sticky="w", padx=4, pady=2)
Entry(root, textvariable=date_var, width=20).grid(row=2, column=1, sticky="ew", padx=4, pady=2)

# Assembly dropdown
Label(root, text="Assembly No").grid(row=2, column=2, sticky="w", padx=4, pady=2)
assembly_combo = ttk.Combobox(root, width=20, state="readonly")
assembly_combo.grid(row=2, column=3, sticky="ew", padx=4, pady=2)

# Dimensions section
Label(root, text="--- DIMENSIONS (Manual Entry) ---",
      font=("Arial", 10, "bold")).grid(row=3, column=0, columnspan=4, sticky="w", pady=5, padx=4)

Label(root, text="Shaft Diameter").grid(row=4, column=0, sticky="w")
Entry(root, textvariable=shaft_dia_var, width=20).grid(row=4, column=1)

Label(root, text="A - Distance").grid(row=4, column=2, sticky="w")
Entry(root, textvariable=a_dist_var, width=20).grid(row=4, column=3)

Label(root, text="B - Distance").grid(row=5, column=0, sticky="w")
Entry(root, textvariable=b_dist_var, width=20).grid(row=5, column=1)

Label(root, text="Diameter at foot").grid(row=5, column=2, sticky="w")
Entry(root, textvariable=mount_hole_var, width=20).grid(row=5, column=3)

Label(root, text="Total Length").grid(row=6, column=0, sticky="w")
Entry(root, textvariable=total_length_var, width=20).grid(row=6, column=1)

Label(root, text="PCD").grid(row=6, column=2, sticky="w")
Entry(root, textvariable=pcd_var, width=20).grid(row=6, column=3)

Label(root, text="Diameter at flange").grid(row=7, column=2, sticky="w")
Entry(root, textvariable=flange_var, width=20).grid(row=7, column=3)

#Label(root, text="Inspector").grid(row=7, column=0, sticky="w")
#Entry(root, textvariable=inspector_var, width=20).grid(row=7, column=1)

# Resistance test section
Label(root, text="--- RESISTANCE TEST (Manual Entry) ---",
      font=("Arial", 10, "bold")).grid(row=8, column=0, columnspan=4, sticky="w", pady=5)

Label(root, text="Resistance at Ambient Temp").grid(row=9, column=0, sticky="w")
Entry(root, textvariable=res_cold_var, width=20, state="readonly").grid(row=9, column=1)

Label(root, text="Resistance at -20 Deg C").grid(row=9, column=2, sticky="w")
Entry(root, textvariable=res_20deg_var, width=20, state="readonly").grid(row=9, column=3)

# Buttons
Button(root, text="Load Data", command=load_data).grid(row=10, column=0, pady=8)
Button(root, text="Generate PDF", command=generate_pdf).grid(row=10, column=1, pady=8)
# start the Tk event loop when the script is executed directly
if __name__ == "__main__":
    root.mainloop()