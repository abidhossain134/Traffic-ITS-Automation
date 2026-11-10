# THIS IS MAIN CODE WITH EXCEL FORMAT CHANGED - Latest from claude
# GUI version: double-click to open (QC_Quantity_Extractor.pyw)

import os
import re
import sys
import traceback
from collections import defaultdict


# ------------------ AUTO-INSTALL MISSING LIBRARIES ------------------
def _ensure_packages():
    """Install any required third-party packages that are missing."""
    import importlib
    import subprocess

    required = {
        # import name : pip package name
        "pdfplumber": "pdfplumber",
        "openpyxl": "openpyxl",
    }

    missing = []
    for mod, pkg in required.items():
        try:
            importlib.import_module(mod)
        except ImportError:
            missing.append(pkg)

    if not missing:
        return

    # Show a small "installing" window so the user knows something is happening
    notice = None
    try:
        import tkinter as tk
        notice = tk.Tk()
        notice.title("QC Quantity Extractor")
        tk.Label(
            notice,
            text="Installing required libraries:\n" + ", ".join(missing) +
                 "\n\nThis only happens once. Please wait...",
            padx=30, pady=20, justify="center"
        ).pack()
        notice.update()
    except Exception:
        notice = None

    creationflags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW on Windows

    def _pip(args):
        return subprocess.run(
            [sys.executable, "-m", "pip", "install", "--disable-pip-version-check"] + args + missing,
            capture_output=True, text=True, creationflags=creationflags
        )

    try:
        # Make sure pip itself is available
        subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"],
                       capture_output=True, creationflags=creationflags)
        result = _pip([])
        if result.returncode != 0:
            # No admin rights? Retry as a per-user install
            result = _pip(["--user"])
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout)
    except Exception as e:
        if notice is not None:
            notice.destroy()
        try:
            from tkinter import messagebox
            messagebox.showerror(
                "Install failed",
                "Could not install: " + ", ".join(missing) +
                "\n\nCheck your internet connection, or run this in a command prompt:\n"
                "pip install " + " ".join(missing) + "\n\nDetails:\n" + str(e)[-800:]
            )
        except Exception:
            print("[ERROR] Could not install:", missing, e)
        sys.exit(1)

    if notice is not None:
        notice.destroy()
    importlib.invalidate_caches()
    # Make a fresh --user install importable in this same run
    try:
        import site
        user_site = site.getusersitepackages()
        if user_site not in sys.path:
            sys.path.append(user_site)
    except Exception:
        pass


_ensure_packages()

import pdfplumber
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ------------------ PATHS (YOUR INPUTS) ------------------
# These are set from the GUI fields when you click Run.
PDF_PATH = r"C:/Abid/Abid/AI AGENTS/qc PDFS/44033825201.pdf"
OUT_XLSX = r"C:/Abid/Abid/AI AGENTS/qc PDFS/44038.xlsx"
OUT_CSV_FALLBACK = r"C:/Abid/Abid/AI AGENTS/qc PDFS/pay_item_quantities_fallback.csv"

# ------------------ PATTERNS ------------------
PAY_ITEM_RE = re.compile(r"^\d{3}-\d{1,3}-\d{1,3}$")  # catches 700-141-140
NUM_TOKEN_RE = re.compile(r"^(?:\d{1,4}|\d,\d{3})(?:\.\d+)?$")

UNITS = {"LF", "EA", "MI", "SF", "AS", "LS", "LU", "DA", "PI", "HR", "LO"}
RUNS_WORDS = {"RUNS", "RUN"}
AT_WORDS = {"@", "AT"}

# For reading sheet numbers at bottom-right
SHEET_ID_RE = re.compile(r"\bIT\s*-\s*\d+\b", re.IGNORECASE)

# ------------------ TUNABLE PARAMETERS ------------------
Y_TOL = 4.0
MAX_HORIZONTAL_DISTANCE = 250
UNIT_GAP_MAX = 70


def autofit_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 60)


def normalize_sheet_id(text):
    return re.sub(r"\s+", "", text.upper())


def extract_sheet_ids(pdf_path):
    """
    Reads bottom-right title block area of each page and extracts sheet IDs like IT-2.
    Returns: dict {page_number: sheet_id}
    """
    page_to_sheet = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            w, h = page.width, page.height

            # Bottom-right crop region
            crop_box = (
                w * 0.78,   # x0
                h * 0.82,   # top
                w * 0.995,  # x1
                h * 0.995   # bottom
            )

            try:
                cropped = page.crop(crop_box)
                text = cropped.extract_text() or ""
            except Exception:
                text = ""

            m = SHEET_ID_RE.search(text)
            if m:
                page_to_sheet[page_num] = normalize_sheet_id(m.group(0))
            else:
                page_to_sheet[page_num] = f"PAGE-{page_num}"

    return page_to_sheet


def find_runs_patterns(words):
    """
    Scan words for the pattern: NUMBER RUNS @ NUMBER UNIT
    e.g. "2 RUNS @ 129 LF"

    Returns:
        runs_pairs   : list of (effective_qty, unit_word, anchor_word)
                       where anchor_word is the run-count word (used for
                       proximity matching to the pay item)
        excluded_idx : set of word indices that are consumed by a RUNS pattern
                       (so they are NOT matched individually as plain numbers)
    """
    runs_pairs = []
    excluded_idx = set()

    for i in range(len(words) - 4):
        w0 = words[i]       # number  (run count)
        w1 = words[i + 1]   # RUNS / RUN
        w2 = words[i + 2]   # @  / AT
        w3 = words[i + 3]   # number  (unit quantity)
        w4 = words[i + 4]   # unit

        if (NUM_TOKEN_RE.match(w0["text"]) and
                w1["text"].upper() in RUNS_WORDS and
                w2["text"].upper() in AT_WORDS and
                NUM_TOKEN_RE.match(w3["text"]) and
                w4["text"].upper() in UNITS):

            run_count = float(w0["text"].replace(",", ""))
            unit_qty  = float(w3["text"].replace(",", ""))
            effective_qty = run_count * unit_qty

            # Use the run-count word as the anchor so that when the pay item
            # is to the LEFT of "2 RUNS @ 57 LF", the "2" position is used
            # for proximity — giving the correct closest match.
            # Quantity is just the unit qty (e.g. 129 LF), NOT multiplied by run count.
            # The run count (w0) is excluded so it isn't mistakenly matched as the qty.
            runs_pairs.append((unit_qty, w4, w0))
            excluded_idx.add(i)   # suppress only the run count from plain matching

    return runs_pairs, excluded_idx


def extract_occurrences(pdf_path):
    """
    Extract occurrences across ALL pages.
    Each occurrence is a row: {page, pay_item, qty, unit}
    """
    occurrences = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"[INFO] Opened PDF. Pages: {total_pages}")

        for page_num, page in enumerate(pdf.pages, start=1):
            if page_num % 5 == 1 or page_num == total_pages:
                print(f"[INFO] Processing page {page_num}/{total_pages} ...")

            words = page.extract_words(keep_blank_chars=False, use_text_flow=True) or []
            if not words:
                continue

            pay_items, numbers, units = [], [], []

            for idx, w in enumerate(words):
                text = (w.get("text") or "").strip()
                if not text:
                    continue
                if PAY_ITEM_RE.match(text):
                    pay_items.append(w)
                elif text.upper() in UNITS:
                    units.append(w)
                elif NUM_TOKEN_RE.match(text):
                    numbers.append((idx, w))

            # --- Detect "N RUNS @ QTY UNIT" patterns first ---
            runs_pairs, excluded_idx = find_runs_patterns(words)

            # Build valid (number + unit) pairs from plain numbers,
            # skipping any number that belongs to a RUNS pattern.
            pairs = []
            for idx, q in numbers:
                if idx in excluded_idx:
                    continue
                qy = q["top"]
                for u in units:
                    if abs(u["top"] - qy) <= Y_TOL:
                        gap = u["x0"] - q["x1"]
                        if 0 <= gap <= UNIT_GAP_MAX:
                            pairs.append((q, u))

            # Add RUNS-derived pairs: (anchor_word, unit_word) with effective qty stored
            # We store them as dicts with synthetic qty baked in so the candidate
            # loop below can treat them uniformly.
            for eff_qty, unit_w, anchor_w in runs_pairs:
                pairs.append(("__runs__", unit_w, eff_qty, anchor_w))

            # Assign closest pair to each pay item
            for pi in pay_items:
                piy = pi["top"]
                pi_center = (pi["x0"] + pi["x1"]) / 2

                candidates = []
                for pair in pairs:
                    if pair[0] == "__runs__":
                        _, u, eff_qty, anchor_w = pair
                        anchor_y = anchor_w["top"]
                        if abs(anchor_y - piy) <= Y_TOL:
                            anchor_center = (anchor_w["x0"] + anchor_w["x1"]) / 2
                            dx = abs(anchor_center - pi_center)
                            if dx <= MAX_HORIZONTAL_DISTANCE:
                                candidates.append((dx, eff_qty, u["text"].upper()))
                    else:
                        q, u = pair
                        if abs(q["top"] - piy) <= Y_TOL:
                            q_center = (q["x0"] + q["x1"]) / 2
                            dx = abs(q_center - pi_center)
                            if dx <= MAX_HORIZONTAL_DISTANCE:
                                candidates.append((dx, float(q["text"].replace(",", "")), u["text"].upper()))

                if not candidates:
                    occurrences.append({
                        "page": page_num,
                        "pay_item": pi["text"],
                        "qty": None,
                        "unit": None
                    })
                    continue

                candidates.sort(key=lambda x: x[0])
                _, qty, unit = candidates[0]

                occurrences.append({
                    "page": page_num,
                    "pay_item": pi["text"],
                    "qty": qty,
                    "unit": unit
                })

    print(f"[INFO] Extracted occurrences: {len(occurrences)}")
    return occurrences


def compute_totals(occurrences):
    totals = defaultdict(float)
    for o in occurrences:
        if o["qty"] is None or o["unit"] is None:
            continue
        totals[(o["pay_item"], o["unit"])] += o["qty"]
    return totals


def organize_occurrences_expressions(occurrences):
    overall_list = defaultdict(list)
    per_page_list = defaultdict(list)

    for o in occurrences:
        if o["qty"] is None or o["unit"] is None:
            continue
        overall_list[(o["pay_item"], o["unit"])].append(o["qty"])
        per_page_list[(o["page"], o["pay_item"], o["unit"])].append(o["qty"])

    def fmt(v):
        return str(int(v)) if float(v).is_integer() else str(v)

    overall_expr = {k: "+".join(fmt(v) for v in vals) for k, vals in overall_list.items()}
    per_page_expr = {k: "+".join(fmt(v) for v in vals) for k, vals in per_page_list.items()}

    return overall_expr, per_page_expr


def sort_sheet_ids(sheet_ids):
    def key_func(s):
        m = re.match(r"([A-Z]+)-(\d+)", s)
        if m:
            return (m.group(1), int(m.group(2)))
        return (s, 999999)
    return sorted(sheet_ids, key=key_func)


def build_per_sheet_expr(per_page_expr, page_to_sheet):
    """
    Convert per-page expressions to per-sheet expressions using sheet IDs.
    Output: {(sheet_id, pay_item, unit): "36+13+273"}
    """
    grouped = defaultdict(list)

    for (page, pay_item, unit), expr in per_page_expr.items():
        sheet_id = page_to_sheet.get(page, f"PAGE-{page}")
        grouped[(sheet_id, pay_item, unit)].append(expr)

    merged = {}
    for key, exprs in grouped.items():
        parts = []
        for e in exprs:
            if e:
                parts.extend(e.split("+"))
        merged[key] = "+".join(parts)

    return merged


def apply_table_border(ws, row_start, row_end, col_start, col_end):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            ws.cell(r, c).border = border


def write_excel_tabulation(occurrences, totals, overall_expr, per_sheet_expr, page_to_sheet, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tabulation"

    # ---------------- Styles ----------------
    header_font = Font(name="Arial", size=11, bold=True, italic=True)
    body_font = Font(name="Arial", size=10, italic=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ---------------- Data prep ----------------
    sheet_ids = sort_sheet_ids(sorted(set(page_to_sheet.values())))
    pay_item_keys = sorted(set((o["pay_item"], o["unit"]) for o in occurrences if o["qty"] is not None and o["unit"] is not None))

    # ---------------- Layout ----------------
    start_row = 1

    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=3 + len(sheet_ids) * 2)
    ws.cell(1, 4).value = "TABULATION OF QUANTITIES"
    ws.cell(1, 4).font = header_font
    ws.cell(1, 4).alignment = center

    ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=3 + len(sheet_ids) * 2)
    ws.cell(2, 4).value = "SHEET NUMBERS"
    ws.cell(2, 4).font = header_font
    ws.cell(2, 4).alignment = center

    ws.merge_cells("A3:A4")
    ws["A3"] = "PAY\nITEM\nNO."
    ws["A3"].font = header_font
    ws["A3"].alignment = center

    ws.merge_cells("B3:B4")
    ws["B3"] = "DESCRIPTION"
    ws["B3"].font = header_font
    ws["B3"].alignment = center

    ws.merge_cells("C3:C4")
    ws["C3"] = "UNIT"
    ws["C3"].font = header_font
    ws["C3"].alignment = center

    col = 4
    sheet_plan_col = {}
    for sid in sheet_ids:
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        ws.cell(3, col).value = sid
        ws.cell(3, col).font = header_font
        ws.cell(3, col).alignment = center

        ws.cell(4, col).value = "PLAN"
        ws.cell(4, col + 1).value = "FINAL"
        ws.cell(4, col).font = header_font
        ws.cell(4, col + 1).font = header_font
        ws.cell(4, col).alignment = center
        ws.cell(4, col + 1).alignment = center

        sheet_plan_col[sid] = col
        col += 2

    grand_plan_col = col
    ws.merge_cells(start_row=3, start_column=grand_plan_col, end_row=3, end_column=grand_plan_col + 1)
    ws.cell(3, grand_plan_col).value = "GRAND\nTOTAL"
    ws.cell(3, grand_plan_col).font = header_font
    ws.cell(3, grand_plan_col).alignment = center

    ws.cell(4, grand_plan_col).value = "PLAN"
    ws.cell(4, grand_plan_col + 1).value = "FINAL"
    ws.cell(4, grand_plan_col).font = header_font
    ws.cell(4, grand_plan_col + 1).font = header_font
    ws.cell(4, grand_plan_col).alignment = center
    ws.cell(4, grand_plan_col + 1).alignment = center

    # ---------------- Fill rows ----------------
    row = 5
    for pay_item, unit in pay_item_keys:
        ws.cell(row, 1).value = pay_item
        ws.cell(row, 2).value = ""
        ws.cell(row, 3).value = unit

        ws.cell(row, 1).font = body_font
        ws.cell(row, 2).font = body_font
        ws.cell(row, 3).font = body_font
        ws.cell(row, 1).alignment = center
        ws.cell(row, 2).alignment = left_align
        ws.cell(row, 3).alignment = center

        for sid in sheet_ids:
            expr = per_sheet_expr.get((sid, pay_item, unit))
            if expr:
                ws.cell(row, sheet_plan_col[sid]).value = f"={expr}"

        overall = overall_expr.get((pay_item, unit))
        if overall:
            ws.cell(row, grand_plan_col).value = f"={overall}"

        row += 1

    # ---------------- Formatting ----------------
    for r in range(5, row):
        for c in range(1, grand_plan_col + 2):
            ws.cell(r, c).font = body_font
            if c == 2:
                ws.cell(r, c).alignment = left_align
            else:
                ws.cell(r, c).alignment = center

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 8

    for c in range(4, grand_plan_col + 2):
        ws.column_dimensions[get_column_letter(c)].width = 10

    apply_table_border(ws, 3, row - 1, 1, grand_plan_col + 1)

    ws2 = wb.create_sheet("Occurrences")
    ws2.append(["page", "sheet_id", "pay_item", "qty", "unit"])
    for o in occurrences:
        ws2.append([o["page"], page_to_sheet.get(o["page"], f"PAGE-{o['page']}"), o["pay_item"], o["qty"], o["unit"]])
    autofit_columns(ws2)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def write_csv_fallback(occurrences, path):
    import csv
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["page", "pay_item", "qty", "unit"])
        for o in occurrences:
            w.writerow([o["page"], o["pay_item"], o["qty"], o["unit"]])


def main():
    print("[INFO] Starting QC quantity extraction...")

    if not os.path.exists(PDF_PATH):
        print(f"[ERROR] PDF not found:\n  {PDF_PATH}")
        sys.exit(1)

    occurrences = extract_occurrences(PDF_PATH)
    totals = compute_totals(occurrences)
    overall_expr, per_page_expr = organize_occurrences_expressions(occurrences)

    page_to_sheet = extract_sheet_ids(PDF_PATH)
    per_sheet_expr = build_per_sheet_expr(per_page_expr, page_to_sheet)

    try:
        write_excel_tabulation(
            occurrences=occurrences,
            totals=totals,
            overall_expr=overall_expr,
            per_sheet_expr=per_sheet_expr,
            page_to_sheet=page_to_sheet,
            out_path=OUT_XLSX
        )
        print(f"[OK] Excel written:\n  {OUT_XLSX}")
    except Exception as e:
        print("[ERROR] Failed to write Excel. Writing CSV fallback instead.")
        print("Reason:", repr(e))
        write_csv_fallback(occurrences, OUT_CSV_FALLBACK)
        print(f"[OK] CSV fallback written:\n  {OUT_CSV_FALLBACK}")
        print("\nFull traceback:")
        traceback.print_exc()

    print("[INFO] Done.")


# =====================================================================
# ------------------------- TKINTER INTERFACE -------------------------
# =====================================================================
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


class _QueueWriter:
    """Sends print() output from the worker thread to the GUI log box."""
    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(text)

    def flush(self):
        pass


class QuantityExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("QC Quantity Extractor")
        root.geometry("720x460")
        root.minsize(600, 380)

        self.pdf_var = tk.StringVar()
        self.xlsx_var = tk.StringVar()
        self.log_q = queue.Queue()
        self.worker = None

        frm = tk.Frame(root, padx=12, pady=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        # PDF field
        tk.Label(frm, text="Plan PDF:").grid(row=0, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.pdf_var).grid(row=0, column=1, sticky="ew", padx=6)
        tk.Button(frm, text="Browse...", width=12, command=self.pick_pdf).grid(row=0, column=2)

        # Excel output field
        tk.Label(frm, text="Output Excel:").grid(row=1, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.xlsx_var).grid(row=1, column=1, sticky="ew", padx=6)
        tk.Button(frm, text="Save As...", width=12, command=self.pick_xlsx).grid(row=1, column=2)

        # Run button
        self.run_btn = tk.Button(frm, text="Run Extraction", width=18,
                                 bg="#1f6feb", fg="white", command=self.start)
        self.run_btn.grid(row=2, column=0, columnspan=3, pady=10)

        # Log box
        self.log = scrolledtext.ScrolledText(frm, height=14, state="disabled",
                                             font=("Consolas", 9))
        self.log.grid(row=3, column=0, columnspan=3, sticky="nsew")
        frm.rowconfigure(3, weight=1)

        self.root.after(100, self.drain_log)

    # ---------- file pickers ----------
    def pick_pdf(self):
        path = filedialog.askopenfilename(
            title="Select plan PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)
            # Suggest an output name next to the PDF if none chosen yet
            if not self.xlsx_var.get().strip():
                base = os.path.splitext(path)[0]
                self.xlsx_var.set(base + "_quantities.xlsx")

    def pick_xlsx(self):
        initial = self.xlsx_var.get().strip()
        path = filedialog.asksaveasfilename(
            title="Save Excel output as",
            defaultextension=".xlsx",
            initialdir=os.path.dirname(initial) if initial else None,
            initialfile=os.path.basename(initial) if initial else "quantities.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if path:
            self.xlsx_var.set(path)

    # ---------- logging ----------
    def append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def drain_log(self):
        try:
            while True:
                self.append_log(self.log_q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self.drain_log)

    # ---------- run ----------
    def start(self):
        global PDF_PATH, OUT_XLSX, OUT_CSV_FALLBACK

        pdf = self.pdf_var.get().strip().strip('"')
        xlsx = self.xlsx_var.get().strip().strip('"')

        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror("Missing PDF", "Please select a valid plan PDF.")
            return
        if not xlsx:
            messagebox.showerror("Missing output", "Please choose where to save the Excel file.")
            return
        if not xlsx.lower().endswith(".xlsx"):
            xlsx += ".xlsx"
            self.xlsx_var.set(xlsx)

        PDF_PATH = pdf
        OUT_XLSX = os.path.abspath(xlsx)
        OUT_CSV_FALLBACK = os.path.join(os.path.dirname(OUT_XLSX),
                                        "pay_item_quantities_fallback.csv")

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.run_btn.configure(state="disabled", text="Running...")

        self.worker = threading.Thread(target=self.run_job, daemon=True)
        self.worker.start()
        self.root.after(200, self.check_done)

    def run_job(self):
        writer = _QueueWriter(self.log_q)
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = writer
        self.success = False
        try:
            main()
            self.success = os.path.exists(OUT_XLSX)
        except SystemExit:
            pass
        except Exception:
            print("[ERROR] Unexpected error:")
            traceback.print_exc()
        finally:
            sys.stdout, sys.stderr = old_out, old_err

    def check_done(self):
        if self.worker and self.worker.is_alive():
            self.root.after(200, self.check_done)
            return
        self.drain_log()
        self.run_btn.configure(state="normal", text="Run Extraction")
        if self.success:
            if messagebox.askyesno("Done", f"Excel saved to:\n{OUT_XLSX}\n\nOpen it now?"):
                try:
                    os.startfile(OUT_XLSX)  # Windows
                except AttributeError:
                    import subprocess
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", OUT_XLSX])
        else:
            messagebox.showwarning("Finished with issues",
                                   "The Excel file was not created. See the log for details.")


if __name__ == "__main__":
    root = tk.Tk()
    QuantityExtractorApp(root)
    root.mainloop()
