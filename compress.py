import os
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def get_gs_executable_path():
    base_path = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    return str(base_path / "tools" / "ghostscript" / "bin" / "gswin64c.exe")

def compress_pdf(input_file, output_file, dpi=150, quality="ebook"):
    gs_path = get_gs_executable_path()
    try:
        subprocess.run([
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-r{dpi}",
            f"-sOutputFile={output_file}",
            input_file
        ], check=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Ghostscript failed: {e}")
    except FileNotFoundError:
        raise RuntimeError(f"Ghostscript not found at: {gs_path}")

def compress_all_pdfs_in_directory(directory, dpi=150, quality="ebook"):
    directory = Path(directory)
    compressed_dir = directory.parent / f"{directory.name}_compressed_pdfs"
    compressed_dir.mkdir(exist_ok=True)

    pdf_files = list(directory.glob("*.pdf"))
    if not pdf_files:
        messagebox.showinfo("Info", "No PDF files found in the selected folder.")
        return

    for pdf in pdf_files:
        print("Compressing:", pdf.name)
        output_path = compressed_dir / pdf.name
        try:
            compress_pdf(str(pdf), str(output_path), dpi, quality)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compress '{pdf.name}': {e}")
    messagebox.showinfo("Done", f"All PDFs compressed into:\n{compressed_dir}")

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    try:
        folder = filedialog.askdirectory(title="Select folder with PDFs to compress")
        if not folder:
            return

        dpi = simpledialog.askinteger("DPI", "Enter DPI (e.g. 150):", initialvalue=150, minvalue=72, maxvalue=300)
        if dpi is None:
            return

        quality = "printer"
        if not quality:
            return
        print("Please wait a moment while the PDFs are being compressed...")
        compress_all_pdfs_in_directory(folder, dpi, quality)
    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))

if __name__ == "__main__":
    main()
