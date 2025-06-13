def compress_all_pdfs_in_directory_with_backup_option(directory, color_image_dpi=150, quality="ebook",
                                                      replace_originals=False, recursive=True, create_backup=True):
    """Compress all PDFs in a directory with backup option control"""
    directory = Path(directory)

    # Setup logging
    log_dir = directory / "compression_logs"
    logger = setup_logging(log_dir)

    try:
        logger.info("=" * 50)
        logger.info("PDF COMPRESSION SESSION STARTED")
        logger.info("=" * 50)
        logger.info(f"Source directory: {directory}")
        logger.info(f"Recursive processing: {recursive}")
        logger.info(f"Color Image DPI setting: {color_image_dpi}")
        logger.info(f"Quality setting: {quality}")
        logger.info(f"Replace originals: {replace_originals}")
        logger.info(f"Create backup: {create_backup}")

        # Create output directory
        if replace_originals:
            output_base_dir = directory / "temp_compressed"
            backup_base_dir = directory / "original_backups" if create_backup else None
        else:
            output_base_dir = directory.parent / f"{directory.name}_compressed_pdfs"
            backup_base_dir = None

        output_base_dir.mkdir(exist_ok=True)
        if backup_base_dir:
            backup_base_dir.mkdir(exist_ok=True)

        logger.info(f"Output base directory: {output_base_dir}")
        if backup_base_dir:
            logger.info(f"Backup base directory: {backup_base_dir}")

        # Find PDF files
        pdf_files = find_all_pdfs(directory, recursive)
        if not pdf_files:
            message = f"No PDF files found in the selected folder{' and its subdirectories' if recursive else ''}."
            logger.warning(message)
            messagebox.showinfo("Info", message)
            return

        logger.info(f"Found {len(pdf_files)} PDF files to compress")

        # Group files by directory for better logging
        files_by_dir = {}
        for pdf_file in pdf_files:
            parent_dir = pdf_file.parent
            if parent_dir not in files_by_dir:
                files_by_dir[parent_dir] = []
            files_by_dir[parent_dir].append(pdf_file)

        logger.info(f"Files distributed across {len(files_by_dir)} directories:")
        for dir_path, files in files_by_dir.items():
            rel_dir = get_relative_path(dir_path, directory)
            logger.info(f"  {rel_dir}: {len(files)} files")

        # Process each PDF
        successful_compressions = 0
        failed_compressions = 0
        total_original_size = 0
        total_compressed_size = 0

        for i, pdf in enumerate(pdf_files, 1):
            try:
                # Get relative path for better logging
                rel_path = get_relative_path(pdf, directory)
                logger.info(f"\n--- Processing file {i}/{len(pdf_files)}: {rel_path} ---")

                # Create output path maintaining directory structure
                if replace_originals:
                    # For replacement, maintain the same relative structure in temp directory
                    relative_pdf_path = pdf.relative_to(directory)
                    output_path = output_base_dir / relative_pdf_path
                    backup_dir = backup_base_dir / relative_pdf_path.parent if backup_base_dir else None
                else:
                    # For separate output, maintain directory structure
                    relative_pdf_path = pdf.relative_to(directory)
                    output_path = output_base_dir / relative_pdf_path
                    backup_dir = None

                # Create output directory if it doesn't exist
                output_path.parent.mkdir(parents=True, exist_ok=True)
                if backup_dir:
                    backup_dir.mkdir(parents=True, exist_ok=True)

                # Compress PDF
                result = compress_pdf(str(pdf), str(output_path), color_image_dpi, quality, logger)

                if result['success']:
                    total_original_size += result['original_size']
                    total_compressed_size += result['compressed_size']

                    # Replace original if requested
                    if replace_originals:
                        try:
                            if create_backup:
                                replace_original_file(pdf, output_path, backup_dir, logger)
                            else:
                                # Replace without backup
                                shutil.move(str(output_path), str(pdf))
                                logger.info(f"Original file replaced without backup: {pdf}")
                        except Exception as e:
                            logger.error(f"Failed to replace original file: {e}")
                            failed_compressions += 1
                            continue

                    successful_compressions += 1
                    logger.info(f"✓ Successfully processed: {rel_path}")

            except Exception as e:
                failed_compressions += 1
                rel_path = get_relative_path(pdf, directory)
                logger.error(f"✗ Failed to process {rel_path}: {e}")
                messagebox.showerror("Error", f"Failed to compress '{rel_path}': {e}")

        # Clean up temp directory if replacing originals
        if replace_originals and output_base_dir.exists():
            try:
                shutil.rmtree(output_base_dir)
                logger.info("Temporary directory cleaned up")
            except Exception as e:
                logger.warning(f"Failed to clean up temporary directory: {e}")

        # Summary
        logger.info("\n" + "=" * 50)
        logger.info("COMPRESSION SESSION SUMMARY")
        logger.info("=" * 50)
        logger.info(f"Total files processed: {len(pdf_files)}")
        logger.info(f"Successful compressions: {successful_compressions}")
        logger.info(f"Failed compressions: {failed_compressions}")

        if successful_compressions > 0:
            overall_compression = (
                                          1 - total_compressed_size / total_original_size) * 100 if total_original_size > 0 else 0
            logger.info(f"Total original size: {total_original_size:.2f} MB")
            logger.info(f"Total compressed size: {total_compressed_size:.2f} MB")
            logger.info(f"Overall compression ratio: {overall_compression:.1f}%")

        # Show completion message
        if replace_originals:
            if create_backup:
                message = f"Compression completed!\n\nProcessed: {successful_compressions}/{len(pdf_files)} files"
                if backup_base_dir and successful_compressions > 0:
                    message += f"\nOriginal files backed up to: {backup_base_dir}"
            else:
                message = f"Compression completed!\n\nProcessed: {successful_compressions}/{len(pdf_files)} files\nOriginal files replaced without backup"
        else:
            message = f"Compression completed!\n\nProcessed: {successful_compressions}/{len(pdf_files)} files\nCompressed files saved to: {output_base_dir}"

        messagebox.showinfo("Compression Complete", message)
        logger.info("PDF compression session completed")

    except Exception as e:
        error_msg = f"Unexpected error during batch compression: {e}"
        logger.error(error_msg)
        messagebox.showerror("Unexpected Error", error_msg)
        raise


def compress_all_pdfs_in_directory(directory, color_image_dpi=150, quality="ebook", replace_originals=False,
                                   recursive=True):
    """Legacy function - compress all PDFs in a directory with basic settings"""
    compress_all_pdfs_in_directory_with_backup_option(directory, color_image_dpi, quality, replace_originals, recursive,
                                                      True)
    import os


import sys
import subprocess
import logging
import shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# Try to import Excel libraries
try:
    import pandas as pd

    EXCEL_SUPPORT = True
except ImportError:
    try:
        import openpyxl

        EXCEL_SUPPORT = True
    except ImportError:
        EXCEL_SUPPORT = False


def setup_logging(log_dir):
    """Setup logging to both console and file"""
    log_dir = Path(log_dir)
    log_dir.mkdir(exist_ok=True)

    # Create log filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"pdf_compression_{timestamp}.log"

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized. Log file: {log_file}")
    return logger


def get_gs_executable_path():
    """Get Ghostscript executable path"""
    try:
        base_path = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
        gs_path = str(base_path / "tools" / "ghostscript" / "bin" / "gswin64c.exe")

        if not Path(gs_path).exists():
            raise FileNotFoundError(f"Ghostscript executable not found at: {gs_path}")

        return gs_path
    except Exception as e:
        raise RuntimeError(f"Failed to locate Ghostscript executable: {e}")


def get_file_size(file_path):
    """Get file size in MB"""
    try:
        return Path(file_path).stat().st_size / (1024 * 1024)
    except Exception as e:
        logging.error(f"Failed to get file size for {file_path}: {e}")
        return 0


def create_backup(file_path, backup_dir):
    """Create backup of original file"""
    try:
        backup_dir = Path(backup_dir)
        backup_dir.mkdir(exist_ok=True)

        source_path = Path(file_path)
        backup_path = backup_dir / source_path.name

        # If backup already exists, add number suffix
        counter = 1
        while backup_path.exists():
            name_stem = source_path.stem
            extension = source_path.suffix
            backup_path = backup_dir / f"{name_stem}_backup_{counter}{extension}"
            counter += 1

        shutil.copy2(source_path, backup_path)
        logging.info(f"Backup created: {backup_path}")
        return str(backup_path)

    except Exception as e:
        logging.error(f"Failed to create backup for {file_path}: {e}")
        raise


def compress_pdf(input_file, output_file, color_image_dpi=150, quality="ebook", logger=None):
    """Compress a single PDF file"""
    if logger is None:
        logger = logging.getLogger(__name__)

    gs_path = get_gs_executable_path()

    try:
        logger.info(f"Starting compression: {input_file}")
        logger.info(f"Output file: {output_file}")
        logger.info(f"Settings - Color Image DPI: {color_image_dpi}, Quality: {quality}")

        # Get original file size
        original_size = get_file_size(input_file)
        logger.info(f"Original file size: {original_size:.2f} MB")

        # Run Ghostscript compression
        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-dColorImageResolution={color_image_dpi}",
            f"-sOutputFile={output_file}",
            input_file
        ]

        logger.info(f"Executing command: {' '.join(cmd)}")

        result = subprocess.run(cmd, check=True, capture_output=True, text=True)

        if result.stderr:
            logger.warning(f"Ghostscript warnings: {result.stderr}")

        # Verify output file was created
        if not Path(output_file).exists():
            raise RuntimeError("Output file was not created")

        # Get compressed file size
        compressed_size = get_file_size(output_file)
        compression_ratio = (1 - compressed_size / original_size) * 100 if original_size > 0 else 0

        logger.info(f"Compressed file size: {compressed_size:.2f} MB")
        logger.info(f"Compression ratio: {compression_ratio:.1f}%")
        logger.info(f"Successfully compressed: {input_file}")

        return {
            'success': True,
            'original_size': original_size,
            'compressed_size': compressed_size,
            'compression_ratio': compression_ratio
        }

    except subprocess.CalledProcessError as e:
        error_msg = f"Ghostscript failed with return code {e.returncode}"
        if e.stderr:
            error_msg += f". Error: {e.stderr}"
        logger.error(error_msg)
        raise RuntimeError(error_msg)

    except FileNotFoundError as e:
        error_msg = f"Ghostscript executable not found: {e}"
        logger.error(error_msg)
        raise RuntimeError(error_msg)

    except Exception as e:
        error_msg = f"Unexpected error during compression: {e}"
        logger.error(error_msg)
        raise RuntimeError(error_msg)


def replace_original_file(original_path, compressed_path, backup_dir, logger=None):
    """Replace original file with compressed version after creating backup"""
    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        original_path = Path(original_path)
        compressed_path = Path(compressed_path)

        if not compressed_path.exists():
            raise FileNotFoundError(f"Compressed file not found: {compressed_path}")

        # Create backup
        backup_path = create_backup(original_path, backup_dir)
        logger.info(f"Original file backed up to: {backup_path}")

        # Replace original with compressed version
        shutil.move(str(compressed_path), str(original_path))
        logger.info(f"Original file replaced with compressed version: {original_path}")

        return backup_path

    except Exception as e:
        error_msg = f"Failed to replace original file {original_path}: {e}"
        logger.error(error_msg)
        raise RuntimeError(error_msg)


def find_all_pdfs(directory, recursive=True):
    """Find all PDF files in directory and subdirectories"""
    directory = Path(directory)
    if recursive:
        # Find PDFs recursively in all subdirectories
        pdf_files = list(directory.rglob("*.pdf"))
    else:
        # Find PDFs only in the current directory
        pdf_files = list(directory.glob("*.pdf"))

    return pdf_files


def get_relative_path(file_path, base_path):
    """Get relative path from base directory"""
    try:
        return Path(file_path).relative_to(Path(base_path))
    except ValueError:
        return Path(file_path).name


def compress_all_pdfs_in_directory(directory, color_image_dpi=150, quality="ebook", replace_originals=False,
                                   recursive=True):
    """Compress all PDFs in a directory and optionally its subdirectories"""
    directory = Path(directory)

    # Setup logging
    log_dir = directory / "compression_logs"
    logger = setup_logging(log_dir)

    try:
        logger.info("=" * 50)
        logger.info("PDF COMPRESSION SESSION STARTED")
        logger.info("=" * 50)
        logger.info(f"Source directory: {directory}")
        logger.info(f"Recursive processing: {recursive}")
        logger.info(f"Color Image DPI setting: {color_image_dpi}")
        logger.info(f"Quality setting: {quality}")
        logger.info(f"Replace originals: {replace_originals}")

        # Create output directory
        if replace_originals:
            output_base_dir = directory / "temp_compressed"
            backup_base_dir = directory / "original_backups"
        else:
            output_base_dir = directory.parent / f"{directory.name}_compressed_pdfs"
            backup_base_dir = None

        output_base_dir.mkdir(exist_ok=True)
        if backup_base_dir:
            backup_base_dir.mkdir(exist_ok=True)

        logger.info(f"Output base directory: {output_base_dir}")
        if backup_base_dir:
            logger.info(f"Backup base directory: {backup_base_dir}")

        # Find PDF files
        pdf_files = find_all_pdfs(directory, recursive)
        if not pdf_files:
            message = f"No PDF files found in the selected folder{' and its subdirectories' if recursive else ''}."
            logger.warning(message)
            messagebox.showinfo("Info", message)
            return

        logger.info(f"Found {len(pdf_files)} PDF files to compress")

        # Group files by directory for better logging
        files_by_dir = {}
        for pdf_file in pdf_files:
            parent_dir = pdf_file.parent
            if parent_dir not in files_by_dir:
                files_by_dir[parent_dir] = []
            files_by_dir[parent_dir].append(pdf_file)

        logger.info(f"Files distributed across {len(files_by_dir)} directories:")
        for dir_path, files in files_by_dir.items():
            rel_dir = get_relative_path(dir_path, directory)
            logger.info(f"  {rel_dir}: {len(files)} files")

        # Process each PDF
        successful_compressions = 0
        failed_compressions = 0
        total_original_size = 0
        total_compressed_size = 0

        for i, pdf in enumerate(pdf_files, 1):
            try:
                # Get relative path for better logging
                rel_path = get_relative_path(pdf, directory)
                logger.info(f"\n--- Processing file {i}/{len(pdf_files)}: {rel_path} ---")

                # Create output path maintaining directory structure
                if replace_originals:
                    # For replacement, maintain the same relative structure in temp directory
                    relative_pdf_path = pdf.relative_to(directory)
                    output_path = output_base_dir / relative_pdf_path
                    backup_dir = backup_base_dir / relative_pdf_path.parent
                else:
                    # For separate output, maintain directory structure
                    relative_pdf_path = pdf.relative_to(directory)
                    output_path = output_base_dir / relative_pdf_path
                    backup_dir = None

                # Create output directory if it doesn't exist
                output_path.parent.mkdir(parents=True, exist_ok=True)
                if backup_dir:
                    backup_dir.mkdir(parents=True, exist_ok=True)

                # Compress PDF
                result = compress_pdf(str(pdf), str(output_path), color_image_dpi, quality, logger)

                if result['success']:
                    total_original_size += result['original_size']
                    total_compressed_size += result['compressed_size']

                    # Replace original if requested
                    if replace_originals:
                        try:
                            replace_original_file(pdf, output_path, backup_dir, logger)
                        except Exception as e:
                            logger.error(f"Failed to replace original file: {e}")
                            failed_compressions += 1
                            continue

                    successful_compressions += 1
                    logger.info(f"✓ Successfully processed: {rel_path}")

            except Exception as e:
                failed_compressions += 1
                rel_path = get_relative_path(pdf, directory)
                logger.error(f"✗ Failed to process {rel_path}: {e}")
                messagebox.showerror("Error", f"Failed to compress '{rel_path}': {e}")

        # Clean up temp directory if replacing originals
        if replace_originals and output_base_dir.exists():
            try:
                shutil.rmtree(output_base_dir)
                logger.info("Temporary directory cleaned up")
            except Exception as e:
                logger.warning(f"Failed to clean up temporary directory: {e}")

        # Summary
        logger.info("\n" + "=" * 50)
        logger.info("COMPRESSION SESSION SUMMARY")
        logger.info("=" * 50)
        logger.info(f"Total files processed: {len(pdf_files)}")
        logger.info(f"Successful compressions: {successful_compressions}")
        logger.info(f"Failed compressions: {failed_compressions}")

        if successful_compressions > 0:
            overall_compression = (
                                          1 - total_compressed_size / total_original_size) * 100 if total_original_size > 0 else 0
            logger.info(f"Total original size: {total_original_size:.2f} MB")
            logger.info(f"Total compressed size: {total_compressed_size:.2f} MB")
            logger.info(f"Overall compression ratio: {overall_compression:.1f}%")

        # Show completion message
        if replace_originals:
            message = f"Compression completed!\n\nProcessed: {successful_compressions}/{len(pdf_files)} files"
            if backup_base_dir and successful_compressions > 0:
                message += f"\nOriginal files backed up to: {backup_base_dir}"
        else:
            message = f"Compression completed!\n\nProcessed: {successful_compressions}/{len(pdf_files)} files\nCompressed files saved to: {output_base_dir}"

        messagebox.showinfo("Compression Complete", message)
        logger.info("PDF compression session completed")

    except Exception as e:
        error_msg = f"Unexpected error during batch compression: {e}"
        logger.error(error_msg)
        messagebox.showerror("Unexpected Error", error_msg)
        raise


class PDFCompressorGUI:
    """Main GUI window for PDF compression with all parameters"""

    def __init__(self, root):
        self.root = root
        self.root.title("PDF Compressor - Advanced Settings")
        self.root.geometry("750x700")  # Adjusted width and height for scrollable interface
        self.root.resizable(True, True)

        # Set minimum window size to ensure usability
        self.root.minsize(600, 500)

        # Variables for settings
        self.selected_folders = []  # List to store multiple folders
        self.color_image_dpi = tk.IntVar(value=150)
        self.quality = tk.StringVar(value="screen")
        self.recursive = tk.BooleanVar(value=True)
        self.replace_mode = tk.StringVar(value="no_replace")  # no_replace, replace_with_backup, replace_without_backup

        # Progress tracking variables - THIS WAS MISSING
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_details = tk.StringVar(value="")

        # Status
        self.status_text = tk.StringVar(value="Ready to compress PDFs")

        self.setup_ui()

    def setup_ui(self):
        # Create main canvas with scrollbar for the entire interface
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)

        # Main frame with padding (now inside scrollable_frame)
        main_frame = tk.Frame(scrollable_frame, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = tk.Label(main_frame, text="PDF Compressor",
                               font=("Arial", 18, "bold"), fg="darkblue")
        title_label.pack(pady=(0, 20))

        # Folder selection frame
        folder_frame = tk.LabelFrame(main_frame, text="Folder Selection", padx=10, pady=10)
        folder_frame.pack(fill=tk.X, pady=10)

        # Add folder controls
        add_folder_frame = tk.Frame(folder_frame)
        add_folder_frame.pack(fill=tk.X, pady=5)

        add_button = tk.Button(add_folder_frame, text="+ Add Folder",
                               command=self.add_folder, bg="lightgreen", font=("Arial", 10, "bold"))
        add_button.pack(side=tk.LEFT)

        # Add Excel import button
        excel_button = tk.Button(add_folder_frame, text="📊 Import from Excel",
                                 command=self.import_from_excel, bg="lightblue", font=("Arial", 10, "bold"))
        excel_button.pack(side=tk.LEFT, padx=(10, 0))

        clear_button = tk.Button(add_folder_frame, text="Clear All",
                                 command=self.clear_folders, bg="lightcoral")
        clear_button.pack(side=tk.LEFT, padx=(10, 0))

        # Selected folders listbox with scrollbar
        listbox_frame = tk.Frame(folder_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.folders_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set,
                                          height=5, font=("Arial", 9))
        self.folders_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.folders_listbox.yview)

        # Add double-click to remove
        self.folders_listbox.bind("<Double-Button-1>", self.remove_selected_folder)

        remove_button = tk.Button(folder_frame, text="Remove Selected",
                                  command=self.remove_selected_folder, bg="orange")
        remove_button.pack(pady=5)

        # Compression settings frame
        settings_frame = tk.LabelFrame(main_frame, text="Compression Settings", padx=10, pady=10)
        settings_frame.pack(fill=tk.X, pady=10)

        # Color Image DPI setting
        dpi_frame = tk.Frame(settings_frame)
        dpi_frame.pack(fill=tk.X, pady=5)

        tk.Label(dpi_frame, text="Color Image DPI:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        dpi_control_frame = tk.Frame(dpi_frame)
        dpi_control_frame.pack(fill=tk.X, pady=2)

        dpi_scale = tk.Scale(dpi_control_frame, from_=72, to=600,
                             variable=self.color_image_dpi, orient=tk.HORIZONTAL,
                             length=300, command=self.update_dpi_label)
        dpi_scale.pack(side=tk.LEFT)

        self.dpi_label = tk.Label(dpi_control_frame, text="150 DPI",
                                  font=("Arial", 10), fg="darkgreen")
        self.dpi_label.pack(side=tk.LEFT, padx=(10, 0))

        # DPI presets
        dpi_presets_frame = tk.Frame(dpi_frame)
        dpi_presets_frame.pack(fill=tk.X, pady=5)

        tk.Button(dpi_presets_frame, text="Low (72)",
                  command=lambda: self.set_dpi(72), bg="lightcoral").pack(side=tk.LEFT, padx=2)
        tk.Button(dpi_presets_frame, text="Medium (150)",
                  command=lambda: self.set_dpi(150), bg="lightyellow").pack(side=tk.LEFT, padx=2)
        tk.Button(dpi_presets_frame, text="High (300)",
                  command=lambda: self.set_dpi(300), bg="lightgreen").pack(side=tk.LEFT, padx=2)

        # Processing options frame
        options_frame = tk.LabelFrame(main_frame, text="Processing Options", padx=10, pady=10)
        options_frame.pack(fill=tk.X, pady=10)

        # Recursive processing
        recursive_check = tk.Checkbutton(options_frame, text="Process subdirectories recursively",
                                         variable=self.recursive, font=("Arial", 10))
        recursive_check.pack(anchor=tk.W, pady=2)

        # Replace options
        tk.Label(options_frame, text="File replacement:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(10, 5))

        replace_options = [
            ("Keep originals (create new compressed files)", "no_replace"),
            ("Replace originals with backup", "replace_with_backup"),
            ("Replace originals without backup", "replace_without_backup")
        ]

        for text, value in replace_options:
            tk.Radiobutton(options_frame, text=text, variable=self.replace_mode,
                           value=value, font=("Arial", 9)).pack(anchor=tk.W, pady=1)

        # Progress frame - THIS WAS MISSING
        progress_frame = tk.LabelFrame(main_frame, text="Progress", padx=10, pady=10)
        progress_frame.pack(fill=tk.X, pady=10)

        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                            maximum=100, length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # Progress details label
        progress_details_label = tk.Label(progress_frame, textvariable=self.progress_details,
                                          font=("Arial", 9), fg="darkblue", wraplength=650)
        progress_details_label.pack(anchor=tk.W)

        # Action buttons frame
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=15)

        # Start compression button
        start_button = tk.Button(buttons_frame, text="Start Compression",
                                 command=self.start_compression,
                                 bg="green", fg="white", font=("Arial", 12, "bold"),
                                 height=2, width=15)
        start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        # Clear/Reset button
        reset_button = tk.Button(buttons_frame, text="Reset Settings",
                                 command=self.reset_settings,
                                 bg="orange", fg="white", font=("Arial", 12, "bold"),
                                 height=2, width=15)
        reset_button.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(5, 0))

        # Status frame
        status_frame = tk.LabelFrame(main_frame, text="Status", padx=10, pady=8)
        status_frame.pack(fill=tk.X, pady=8)

        status_label = tk.Label(status_frame, textvariable=self.status_text,
                                font=("Arial", 10), fg="blue", wraplength=650)
        status_label.pack(anchor=tk.W)

        # Info frame
        info_frame = tk.LabelFrame(main_frame, text="Information", padx=10, pady=5)
        info_frame.pack(fill=tk.X, pady=5)

        info_text = tk.Text(info_frame, height=3, wrap=tk.WORD, font=("Arial", 9))
        info_text.pack(fill=tk.X)
        info_text.insert(tk.END,
                         "• Double-click on folder list to remove a folder\n"
                         "• Use Color Image DPI to control compression level (lower = smaller files)\n"
                         "• Screen quality preset is used for all compressions (optimized for web/email)\n"
                         "• Import from Excel: First column should contain folder paths")
        info_text.config(state=tk.DISABLED)
    def update_dpi_label(self, value):
        """Update DPI label when scale changes"""
        dpi_val = int(float(value))
        self.dpi_label.config(text=f"{dpi_val} DPI")

    def set_dpi(self, dpi_value):
        """Set DPI to specific value"""
        self.color_image_dpi.set(dpi_value)
        self.update_dpi_label(dpi_value)

    def add_folder(self):
        """Add a folder to the selection list"""
        folder = filedialog.askdirectory(title="Select folder with PDFs to compress")
        if folder and folder not in self.selected_folders:
            self.selected_folders.append(folder)
            self.folders_listbox.insert(tk.END, folder)
            self.status_text.set(f"Added folder: {Path(folder).name} ({len(self.selected_folders)} total)")
        elif folder in self.selected_folders:
            messagebox.showinfo("Info", "This folder is already in the list!")

    def import_from_excel(self):
        """Import folder paths from Excel file (first column)"""
        if not EXCEL_SUPPORT:
            messagebox.showerror("Error",
                                 "Excel support not available. Please install:\n"
                                 "pip install pandas openpyxl\n"
                                 "or\n"
                                 "pip install openpyxl")
            return

        excel_file = filedialog.askopenfilename(
            title="Select Excel file with folder paths",
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("All files", "*.*")
            ]
        )

        if not excel_file:
            return

        try:
            # Try reading with pandas first
            if 'pandas' in sys.modules or 'pd' in globals():
                paths = self.read_excel_with_pandas(excel_file)
            else:
                paths = self.read_excel_with_openpyxl(excel_file)

            # Filter valid paths and add them
            added_count = 0
            skipped_count = 0
            invalid_count = 0

            for path_str in paths:
                if not path_str or str(path_str).strip() == '' or str(path_str) == 'nan':
                    continue

                path_str = str(path_str).strip()
                path_obj = Path(path_str)

                # Check if path exists
                if not path_obj.exists():
                    invalid_count += 1
                    continue

                # Check if path is a directory
                if not path_obj.is_dir():
                    invalid_count += 1
                    continue

                # Check if already in list
                if path_str in self.selected_folders:
                    skipped_count += 1
                    continue

                # Add to list
                self.selected_folders.append(path_str)
                self.folders_listbox.insert(tk.END, path_str)
                added_count += 1

            # Show results
            result_msg = f"Excel import completed!\n\n"
            result_msg += f"Added: {added_count} folders\n"
            if skipped_count > 0:
                result_msg += f"Skipped (duplicates): {skipped_count}\n"
            if invalid_count > 0:
                result_msg += f"Invalid/missing paths: {invalid_count}\n"
            result_msg += f"Total folders: {len(self.selected_folders)}"

            messagebox.showinfo("Import Results", result_msg)
            self.status_text.set(f"Imported {added_count} folders from Excel ({len(self.selected_folders)} total)")

        except Exception as e:
            messagebox.showerror("Excel Import Error",
                                 f"Failed to read Excel file:\n{str(e)}\n\n"
                                 "Make sure:\n"
                                 "• File is a valid Excel file (.xlsx or .xls)\n"
                                 "• First column contains folder paths\n"
                                 "• File is not open in another program")

    def read_excel_with_pandas(self, excel_file):
        """Read Excel file using pandas"""
        try:
            # Read first column only
            df = pd.read_excel(excel_file, usecols=[0], header=None)
            return df.iloc[:, 0].dropna().tolist()
        except Exception as e:
            raise Exception(f"Pandas read error: {str(e)}")

    def read_excel_with_openpyxl(self, excel_file):
        """Read Excel file using openpyxl (fallback method)"""
        try:
            import openpyxl
            workbook = openpyxl.load_workbook(excel_file, data_only=True)
            worksheet = workbook.active

            paths = []
            for row in worksheet.iter_rows(min_col=1, max_col=1, values_only=True):
                if row[0] is not None:
                    paths.append(str(row[0]).strip())

            workbook.close()
            return paths
        except Exception as e:
            raise Exception(f"OpenPyXL read error: {str(e)}")

    def remove_selected_folder(self, event=None):
        """Remove selected folder from the list"""
        selection = self.folders_listbox.curselection()
        if selection:
            index = selection[0]
            removed_folder = self.selected_folders.pop(index)
            self.folders_listbox.delete(index)
            self.status_text.set(f"Removed folder: {Path(removed_folder).name}")

    def clear_folders(self):
        """Clear all selected folders"""
        if self.selected_folders:
            if messagebox.askyesno("Confirm", "Clear all selected folders?"):
                self.selected_folders.clear()
                self.folders_listbox.delete(0, tk.END)
                self.status_text.set("All folders cleared")

    def reset_settings(self):
        """Reset all settings to defaults"""
        self.selected_folders.clear()
        self.folders_listbox.delete(0, tk.END)
        self.color_image_dpi.set(150)
        self.quality.set("screen")
        self.recursive.set(True)
        self.replace_mode.set("no_replace")
        self.status_text.set("Settings reset to defaults")
        self.progress_details.set("")
        self.progress_var.set(0)
        self.update_dpi_label(150)

    def start_compression(self):
        """Start the compression process"""
        if not self.selected_folders:
            messagebox.showerror("Error", "Please add at least one folder!")
            return

        # Verify all folders exist
        invalid_folders = [f for f in self.selected_folders if not Path(f).exists()]
        if invalid_folders:
            messagebox.showerror("Error", f"These folders do not exist:\n{chr(10).join(invalid_folders)}")
            return

        # Get settings
        color_image_dpi = self.color_image_dpi.get()
        quality = self.quality.get()
        recursive = self.recursive.get()
        replace_mode = self.replace_mode.get()

        # Convert replace mode to boolean values
        replace_originals = replace_mode in ["replace_with_backup", "replace_without_backup"]
        create_backup = replace_mode == "replace_with_backup"

        # Show confirmation
        settings_summary = f"""Compression Settings:

Selected Folders: {len(self.selected_folders)} folder(s)
{chr(10).join([f"  • {Path(f).name}" for f in self.selected_folders])}

Color Image DPI: {color_image_dpi}
Process subdirectories: {'Yes' if recursive else 'No'}
File handling: {
        'Keep originals (create new files)' if replace_mode == 'no_replace'
        else 'Replace originals with backup' if replace_mode == 'replace_with_backup'
        else 'Replace originals without backup'
        }

Do you want to start compression with these settings?"""

        if not messagebox.askyesno("Confirm Compression", settings_summary):
            return

        # Reset progress
        self.progress_var.set(0)
        self.progress_details.set("")

        # Update status
        self.status_text.set("Starting compression...")
        self.root.update()

        try:
            print("Starting PDF compression...")
            print("Check the console and log files for detailed progress...")

            # Process each folder
            total_processed = 0
            total_successful = 0
            total_failed = 0

            for i, folder in enumerate(self.selected_folders, 1):
                folder_progress = (i - 1) / len(self.selected_folders) * 100
                self.progress_var.set(folder_progress)

                self.status_text.set(f"Processing folder {i}/{len(self.selected_folders)}: {Path(folder).name}")
                self.progress_details.set(f"Folder: {folder}")
                self.root.update()

                # Process folder with progress callback
                result = self.compress_folder_with_progress(
                    folder, color_image_dpi, quality, replace_originals, recursive, create_backup, i,
                    len(self.selected_folders)
                )

                total_processed += 1
                total_successful += result.get('successful', 0)
                total_failed += result.get('failed', 0)

            # Final update
            self.progress_var.set(100)
            self.status_text.set(f"Compression completed! Processed {total_processed} folder(s)")
            self.progress_details.set(f"Total: {total_successful} successful, {total_failed} failed")

            # Show final summary
            summary_msg = f"Compression completed!\n\n"
            summary_msg += f"Folders processed: {total_processed}\n"
            summary_msg += f"Files successful: {total_successful}\n"
            if total_failed > 0:
                summary_msg += f"Files failed: {total_failed}\n"
            summary_msg += f"\nCheck log files for detailed information."

            messagebox.showinfo("Compression Complete", summary_msg)

        except Exception as e:
            error_msg = f"Compression failed: {str(e)}"
            logging.error(f"Application error: {e}")
            messagebox.showerror("Compression Error", error_msg)
            self.status_text.set("Compression failed!")
            self.progress_details.set("Error occurred")

    def compress_folder_with_progress(self, folder, color_image_dpi, quality, replace_originals, recursive,
                                      create_backup, folder_index, total_folders):
        """Compress folder with progress updates"""
        try:
            # Get all PDF files first to calculate progress
            pdf_files = find_all_pdfs(folder, recursive)
            if not pdf_files:
                self.progress_details.set(f"No PDF files found in {Path(folder).name}")
                return {'successful': 0, 'failed': 0}

            successful = 0
            failed = 0

            # Setup logging for this folder
            directory = Path(folder)
            log_dir = directory / "compression_logs"
            logger = setup_logging(log_dir)

            logger.info(f"Processing folder: {folder}")
            logger.info(f"Found {len(pdf_files)} PDF files")

            # Create output directories
            if replace_originals:
                output_base_dir = directory / "temp_compressed"
                backup_base_dir = directory / "original_backups" if create_backup else None
            else:
                output_base_dir = directory.parent / f"{directory.name}_compressed_pdfs"
                backup_base_dir = None

            output_base_dir.mkdir(exist_ok=True)
            if backup_base_dir:
                backup_base_dir.mkdir(exist_ok=True)

            # Process each PDF
            for i, pdf_file in enumerate(pdf_files, 1):
                try:
                    # Calculate overall progress
                    folder_progress = ((folder_index - 1) + (i / len(pdf_files))) / total_folders * 100
                    self.progress_var.set(folder_progress)

                    rel_path = get_relative_path(pdf_file, directory)
                    self.progress_details.set(f"Processing: {rel_path} ({i}/{len(pdf_files)})")
                    self.root.update()

                    # Create output path
                    if replace_originals:
                        relative_pdf_path = pdf_file.relative_to(directory)
                        output_path = output_base_dir / relative_pdf_path
                        backup_dir = backup_base_dir / relative_pdf_path.parent if backup_base_dir else None
                    else:
                        relative_pdf_path = pdf_file.relative_to(directory)
                        output_path = output_base_dir / relative_pdf_path
                        backup_dir = None

                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    if backup_dir:
                        backup_dir.mkdir(parents=True, exist_ok=True)

                    # Compress PDF
                    result = compress_pdf(str(pdf_file), str(output_path), color_image_dpi, quality, logger)

                    if result['success']:
                        # Replace original if requested
                        if replace_originals:
                            if create_backup:
                                replace_original_file(pdf_file, output_path, backup_dir, logger)
                            else:
                                shutil.move(str(output_path), str(pdf_file))
                                logger.info(f"Original file replaced without backup: {pdf_file}")

                        successful += 1
                        logger.info(f"✓ Successfully processed: {rel_path}")
                    else:
                        failed += 1
                        logger.error(f"✗ Failed to process: {rel_path}")

                except Exception as e:
                    failed += 1
                    rel_path = get_relative_path(pdf_file, directory)
                    logger.error(f"✗ Failed to process {rel_path}: {e}")

            # Clean up temp directory if replacing originals
            if replace_originals and output_base_dir.exists():
                try:
                    shutil.rmtree(output_base_dir)
                    logger.info("Temporary directory cleaned up")
                except Exception as e:
                    logger.warning(f"Failed to clean up temporary directory: {e}")

            logger.info(f"Folder processing complete: {successful} successful, {failed} failed")
            return {'successful': successful, 'failed': failed}

        except Exception as e:
            logging.error(f"Error processing folder {folder}: {e}")
            return {'successful': 0, 'failed': 1}


def main():
    """Main function with GUI"""
    root = tk.Tk()
    app = PDFCompressorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()