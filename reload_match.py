import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import re
from PyPDF2 import PdfReader


# --------------------------------------------------
# Extract LPNs from PDFs
# --------------------------------------------------
def extract_lpns_from_pdfs(pdf_paths):
    """
    Extract numeric LPN values from PDFs.
    LPNs are numeric and typically 8–12 digits.
    """
    lpns = set()
    lpn_pattern = re.compile(r"\b\d{8,12}\b")

    for pdf_path in pdf_paths:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            for match in lpn_pattern.findall(text):
                lpns.add(match)

    return lpns


# --------------------------------------------------
# GUI Application
# --------------------------------------------------
class ReceiveMatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Receive Match Checker")

        self.excel_file = None
        self.pdf_files = []

        tk.Button(root, text="Select Excel File", width=60, command=self.load_excel).pack(pady=5)
        tk.Button(root, text="Select PDF Files", width=60, command=self.load_pdfs).pack(pady=5)
        tk.Button(root, text="Run Receive Match", width=60, command=self.run_match).pack(pady=15)

        self.status = tk.Label(root, text="Waiting for files...", fg="blue")
        self.status.pack()

    # ----------------------------
    # Loaders
    # ----------------------------
    def load_excel(self):
        self.excel_file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if self.excel_file:
            self.status.config(text=f"Excel loaded: {Path(self.excel_file).name}")

    def load_pdfs(self):
        self.pdf_files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if self.pdf_files:
            self.status.config(text=f"{len(self.pdf_files)} PDF(s) loaded")

    # ----------------------------
    # Matching Logic
    # ----------------------------
    def run_match(self):
        if not self.excel_file or not self.pdf_files:
            messagebox.showerror(
                "Missing Files",
                "Please select both an Excel file and PDF files."
            )
            return

        try:
            # 🔑 Header is on row 2
            df = pd.read_excel(self.excel_file, header=1, dtype=str).fillna("")

            # Normalize headers
            df.columns = df.columns.astype(str).str.strip().str.upper()

            if "PACKAGEID" not in df.columns:
                raise ValueError(
                    "Excel file must contain a 'PACKAGEID' column on row 2"
                )

            # Extract LPNs from PDFs
            lpns_in_pdfs = extract_lpns_from_pdfs(self.pdf_files)

            # Determine matches
            def match_lpn(pkg_id):
                pkg_id = str(pkg_id).strip()
                if pkg_id in lpns_in_pdfs:
                    return pkg_id, "YES"
                return "", "NO"

            pdf_lpn_col = []
            receive_match_col = []

            for pkg in df["PACKAGEID"]:
                lpn, match = match_lpn(pkg)
                pdf_lpn_col.append(lpn)
                receive_match_col.append(match)

            # Insert columns in correct order
            df["PDF LPN"] = pdf_lpn_col
            df["RECEIVE MATCH"] = receive_match_col

            # Save output
            excel_path = Path(self.excel_file)
            output_path = excel_path.with_name(
                f"{excel_path.stem}_RECEIVE_MATCH.xlsx"
            )

            df.to_excel(output_path, index=False)

            messagebox.showinfo(
                "Success",
                f"Receive match completed.\n\nOutput file:\n{output_path.name}"
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))


# --------------------------------------------------
# Run App
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("700x380")
    app = ReceiveMatchApp(root)
    root.mainloop()
