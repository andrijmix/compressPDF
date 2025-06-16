import os
import sys
import subprocess
import logging
import shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
import time
import multiprocessing

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


def setup_logging():
    """Setup logging to both console and file in the application directory"""
    # Get the directory where the script/executable is located
    if getattr(sys, 'frozen', False):
        # If running as compiled executable
        app_dir = Path(sys.executable).parent
    else:
        # If running as script
        app_dir = Path(__file__).parent

    # Create logs directory in the application directory
    log_dir = app_dir / "logs"
    log_dir.mkdir(exist_ok=True)

    # Create log filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"pdf_compression_{timestamp}.log"

    # Remove any existing handlers to avoid duplicates
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

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
    logger.info(f"Application directory: {app_dir}")
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


def compress_single_pdf_task(args):
    """Worker function for processing a single PDF file in thread pool"""
    pdf_path, output_path, backup_dir, color_image_dpi, quality, replace_originals, create_backup, logger, thread_id = args

    try:
        # Log which thread is processing this file for debugging
        rel_path = get_relative_path(pdf_path, pdf_path.parent.parent if pdf_path.parent.parent else pdf_path.parent)
        logger.info(f"🧵 Thread-{thread_id}: Starting {rel_path}")

        # Compress PDF
        result = compress_pdf(str(pdf_path), str(output_path), color_image_dpi, quality, logger)

        if result['success']:
            # Replace original if requested
            if replace_originals:
                try:
                    if create_backup:
                        replace_original_file(pdf_path, output_path, backup_dir, logger)
                    else:
                        # Replace without backup
                        shutil.move(str(output_path), str(pdf_path))
                        logger.info(f"Original file replaced without backup: {pdf_path}")
                except Exception as e:
                    logger.error(f"Failed to replace original file: {e}")
                    return {'status': 'failed', 'pdf_path': pdf_path, 'error': str(e), 'thread_id': thread_id}

            logger.info(f"🧵 Thread-{thread_id}: Completed {rel_path}")
            return {
                'status': 'success',
                'pdf_path': pdf_path,
                'original_size': result['original_size'],
                'compressed_size': result['compressed_size'],
                'compression_ratio': result['compression_ratio'],
                'thread_id': thread_id
            }
        else:
            return {'status': 'failed', 'pdf_path': pdf_path, 'error': 'Compression failed', 'thread_id': thread_id}

    except Exception as e:
        logger.error(f"🧵 Thread-{thread_id}: Error processing {pdf_path}: {e}")
        return {'status': 'failed', 'pdf_path': pdf_path, 'error': str(e), 'thread_id': thread_id}


def compress_all_pdfs_in_directory_threaded(directory, color_image_dpi=150, quality="ebook",
                                            replace_originals=False, recursive=True, create_backup=True,
                                            min_file_size_mb=1.0, max_threads=4, logger=None):
    """Compress all PDFs in a directory using multithreading"""
    directory = Path(directory)

    # Use existing logger if provided, otherwise get default logger
    if logger is None:
        logger = logging.getLogger(__name__)

    try:
        logger.info("=" * 50)
        logger.info(f"PROCESSING DIRECTORY WITH THREADING: {directory}")
        logger.info("=" * 50)
        logger.info(f"Source directory: {directory}")
        logger.info(f"Recursive processing: {recursive}")
        logger.info(f"Color Image DPI setting: {color_image_dpi}")
        logger.info(f"Quality setting: {quality}")
        logger.info(f"Replace originals: {replace_originals}")
        logger.info(f"Create backup: {create_backup}")
        logger.info(f"Minimum file size: {min_file_size_mb:.1f} MB")
        logger.info(f"Max threads: {max_threads}")

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
            return {'successful': 0, 'failed': 0, 'skipped': 0, 'message': message}

        logger.info(f"Found {len(pdf_files)} PDF files to process")

        # Filter files by size
        files_to_process = []
        skipped_files = []

        for pdf_file in pdf_files:
            try:
                file_size_mb = get_file_size(str(pdf_file))
                if file_size_mb >= min_file_size_mb:
                    files_to_process.append(pdf_file)
                else:
                    skipped_files.append((pdf_file, file_size_mb))
                    rel_path = get_relative_path(pdf_file, directory)
                    logger.info(
                        f"⏭️ Skipping small file: {rel_path} ({file_size_mb:.2f} MB < {min_file_size_mb:.1f} MB)")
            except Exception as e:
                logger.warning(f"Could not check size of {pdf_file}: {e}")
                files_to_process.append(pdf_file)  # Process it anyway if we can't check size

        logger.info(f"Files to process: {len(files_to_process)}")
        logger.info(f"Files skipped (too small): {len(skipped_files)}")

        if not files_to_process:
            message = f"No PDF files meet the minimum size requirement ({min_file_size_mb:.1f} MB)."
            logger.warning(message)
            return {'successful': 0, 'failed': 0, 'skipped': len(skipped_files), 'message': message}

        # Prepare tasks for thread pool
        tasks = []
        for i, pdf_file in enumerate(files_to_process):
            # Create output path maintaining directory structure
            if replace_originals:
                relative_pdf_path = pdf_file.relative_to(directory)
                output_path = output_base_dir / relative_pdf_path
                backup_dir = backup_base_dir / relative_pdf_path.parent if backup_base_dir else None
            else:
                relative_pdf_path = pdf_file.relative_to(directory)
                output_path = output_base_dir / relative_pdf_path
                backup_dir = None

            # Create output directory if it doesn't exist
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if backup_dir:
                backup_dir.mkdir(parents=True, exist_ok=True)

            # Create task tuple with thread ID for tracking
            thread_id = (i % max_threads) + 1  # Assign thread IDs cyclically
            task = (
            pdf_file, output_path, backup_dir, color_image_dpi, quality, replace_originals, create_backup, logger,
            thread_id)
            tasks.append(task)

        # Process files using thread pool
        successful_compressions = 0
        failed_compressions = 0
        total_original_size = 0
        total_compressed_size = 0

        logger.info(f"Starting parallel processing with {max_threads} threads...")
        logger.info(f"Tasks distribution: {len(tasks)} files across {max_threads} threads")
        start_time = time.time()

        # Track thread usage
        thread_usage = {}

        with ThreadPoolExecutor(max_workers=max_threads, thread_name_prefix="PDFCompressor") as executor:
            # Submit all tasks
            future_to_task = {executor.submit(compress_single_pdf_task, task): task for task in tasks}

            # Process completed tasks
            for i, future in enumerate(as_completed(future_to_task), 1):
                task = future_to_task[future]
                pdf_file = task[0]
                rel_path = get_relative_path(pdf_file, directory)

                try:
                    result = future.result()
                    thread_id = result.get('thread_id', 'Unknown')

                    # Track thread usage
                    if thread_id not in thread_usage:
                        thread_usage[thread_id] = 0
                    thread_usage[thread_id] += 1

                    if result['status'] == 'success':
                        successful_compressions += 1
                        total_original_size += result['original_size']
                        total_compressed_size += result['compressed_size']
                        logger.info(
                            f"✓ [{i}/{len(tasks)}] T{thread_id}: {rel_path} ({result['compression_ratio']:.1f}% reduction)")
                    else:
                        failed_compressions += 1
                        logger.error(
                            f"✗ [{i}/{len(tasks)}] T{thread_id}: {rel_path} - {result.get('error', 'Unknown error')}")

                except Exception as e:
                    failed_compressions += 1
                    logger.error(f"✗ [{i}/{len(tasks)}] Exception processing {rel_path}: {e}")

        processing_time = time.time() - start_time

        # Log thread usage statistics
        logger.info(f"Thread usage statistics:")
        for thread_id, count in sorted(thread_usage.items()):
            logger.info(f"  Thread-{thread_id}: processed {count} files")

        logger.info(f"Parallel processing completed in {processing_time:.1f} seconds")

        # Clean up temp directory if replacing originals
        if replace_originals and output_base_dir.exists():
            try:
                shutil.rmtree(output_base_dir)
                logger.info("Temporary directory cleaned up")
            except Exception as e:
                logger.warning(f"Failed to clean up temporary directory: {e}")

        # Summary for this directory
        logger.info(f"\n--- THREADED DIRECTORY SUMMARY: {directory.name} ---")
        logger.info(f"Processing time: {processing_time:.1f} seconds")
        logger.info(f"Total files found: {len(pdf_files)}")
        logger.info(f"Files processed: {len(files_to_process)}")
        logger.info(f"Files skipped (too small): {len(skipped_files)}")
        logger.info(f"Successful compressions: {successful_compressions}")
        logger.info(f"Failed compressions: {failed_compressions}")

        if successful_compressions > 0:
            files_per_second = successful_compressions / processing_time if processing_time > 0 else 0
            logger.info(f"Processing speed: {files_per_second:.1f} files/second")

            overall_compression = (
                                              1 - total_compressed_size / total_original_size) * 100 if total_original_size > 0 else 0
            logger.info(f"Total original size: {total_original_size:.2f} MB")
            logger.info(f"Total compressed size: {total_compressed_size:.2f} MB")
            logger.info(f"Overall compression ratio: {overall_compression:.1f}%")

        return {
            'successful': successful_compressions,
            'failed': failed_compressions,
            'skipped': len(skipped_files),
            'total_original_size': total_original_size,
            'total_compressed_size': total_compressed_size,
            'processing_time': processing_time
        }

    except Exception as e:
        error_msg = f"Unexpected error during threaded batch compression in {directory}: {e}"
        logger.error(error_msg)
        raise


class PDFCompressorGUI:
    """Main GUI window for PDF compression with all parameters"""

    def __init__(self, root):
        self.root = root
        self.root.title("PDF Compressor - Advanced Settings")
        self.root.geometry("750x800")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)

        # Initialize logger at startup
        self.logger = setup_logging()

        # Get CPU count for threading
        self.cpu_count = multiprocessing.cpu_count()
        default_threads = min(max(2, self.cpu_count - 1), 16)  # Leave 1 CPU free, max 16 threads

        # Variables for settings
        self.selected_folders = []
        self.color_image_dpi = tk.IntVar(value=150)
        self.quality = tk.StringVar(value="screen")
        self.recursive = tk.BooleanVar(value=True)
        self.replace_mode = tk.StringVar(value="no_replace")
        self.min_file_size = tk.DoubleVar(value=1.0)
        self.max_threads = tk.IntVar(value=default_threads)

        # Threading control
        self.processing_thread = None
        self.stop_processing = threading.Event()

        # Progress tracking variables
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_details = tk.StringVar(value="")

        # Status
        self.status_text = tk.StringVar(value="Ready to compress PDFs")

        # Create UI
        self.setup_ui()

    def setup_ui(self):
        # Create main canvas with scrollbar
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Main frame
        main_frame = tk.Frame(scrollable_frame, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = tk.Label(main_frame, text="PDF Compressor",
                               font=("Arial", 18, "bold"), fg="darkblue")
        title_label.pack(pady=(0, 20))

        # Folder selection frame
        self.create_folder_selection_frame(main_frame)

        # Compression settings frame
        self.create_compression_settings_frame(main_frame)

        # Processing options frame
        self.create_processing_options_frame(main_frame)

        # Progress frame
        self.create_progress_frame(main_frame)

        # Action buttons frame
        self.create_action_buttons_frame(main_frame)

        # Status frame
        self.create_status_frame(main_frame)

        # Info frame
        self.create_info_frame(main_frame)

    def create_folder_selection_frame(self, parent):
        """Create folder selection UI components"""
        folder_frame = tk.LabelFrame(parent, text="Folder Selection", padx=10, pady=10)
        folder_frame.pack(fill=tk.X, pady=10)

        # Add folder controls
        add_folder_frame = tk.Frame(folder_frame)
        add_folder_frame.pack(fill=tk.X, pady=5)

        add_button = tk.Button(add_folder_frame, text="+ Add Folder",
                               command=self.add_folder, bg="lightgreen", font=("Arial", 10, "bold"))
        add_button.pack(side=tk.LEFT)

        excel_button = tk.Button(add_folder_frame, text="📊 Import from Excel",
                                 command=self.import_from_excel, bg="lightblue", font=("Arial", 10, "bold"))
        excel_button.pack(side=tk.LEFT, padx=(10, 0))

        clear_button = tk.Button(add_folder_frame, text="Clear All",
                                 command=self.clear_folders, bg="lightcoral")
        clear_button.pack(side=tk.LEFT, padx=(10, 0))

        # Selected folders listbox
        listbox_frame = tk.Frame(folder_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        listbox_scrollbar = tk.Scrollbar(listbox_frame)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.folders_listbox = tk.Listbox(listbox_frame, yscrollcommand=listbox_scrollbar.set,
                                          height=5, font=("Arial", 9))
        self.folders_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        listbox_scrollbar.config(command=self.folders_listbox.yview)

        self.folders_listbox.bind("<Double-Button-1>", self.remove_selected_folder)

        remove_button = tk.Button(folder_frame, text="Remove Selected",
                                  command=self.remove_selected_folder, bg="orange")
        remove_button.pack(pady=5)

    def create_compression_settings_frame(self, parent):
        """Create compression settings UI components"""
        settings_frame = tk.LabelFrame(parent, text="Compression Settings", padx=10, pady=10)
        settings_frame.pack(fill=tk.X, pady=10)

        # DPI setting
        self.create_dpi_setting(settings_frame)

        # File size setting
        self.create_file_size_setting(settings_frame)

        # Threading setting
        self.create_threading_setting(settings_frame)

    def create_dpi_setting(self, parent):
        """Create DPI setting controls"""
        dpi_frame = tk.Frame(parent)
        dpi_frame.pack(fill=tk.X, pady=5)

        tk.Label(dpi_frame, text="Color Image DPI:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        dpi_control_frame = tk.Frame(dpi_frame)
        dpi_control_frame.pack(fill=tk.X, pady=2)

        # Create label first
        self.dpi_label = tk.Label(dpi_control_frame, text="150 DPI",
                                  font=("Arial", 10), fg="darkgreen")

        dpi_scale = tk.Scale(dpi_control_frame, from_=72, to=600,
                             variable=self.color_image_dpi, orient=tk.HORIZONTAL,
                             length=300, command=self.update_dpi_label)
        dpi_scale.pack(side=tk.LEFT)

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

    def create_file_size_setting(self, parent):
        """Create file size setting controls"""
        size_frame = tk.Frame(parent)
        size_frame.pack(fill=tk.X, pady=5)

        tk.Label(size_frame, text="Skip files smaller than:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        size_control_frame = tk.Frame(size_frame)
        size_control_frame.pack(fill=tk.X, pady=2)

        # Create label first
        self.size_label = tk.Label(size_control_frame, text="1.0 MB",
                                   font=("Arial", 10), fg="darkgreen")

        size_scale = tk.Scale(size_control_frame, from_=0.1, to=10.0,
                              variable=self.min_file_size, orient=tk.HORIZONTAL,
                              length=300, resolution=0.1, command=self.update_size_label)
        size_scale.pack(side=tk.LEFT)

        self.size_label.pack(side=tk.LEFT, padx=(10, 0))

        # Size presets
        size_presets_frame = tk.Frame(size_frame)
        size_presets_frame.pack(fill=tk.X, pady=5)

        tk.Button(size_presets_frame, text="Skip None (0.1 MB)",
                  command=lambda: self.set_min_size(0.1), bg="lightcoral").pack(side=tk.LEFT, padx=2)
        tk.Button(size_presets_frame, text="Small Files (1 MB)",
                  command=lambda: self.set_min_size(1.0), bg="lightyellow").pack(side=tk.LEFT, padx=2)
        tk.Button(size_presets_frame, text="Medium Files (5 MB)",
                  command=lambda: self.set_min_size(5.0), bg="lightgreen").pack(side=tk.LEFT, padx=2)
        tk.Button(size_presets_frame, text="Large Files (10 MB)",
                  command=lambda: self.set_min_size(10.0), bg="lightblue").pack(side=tk.LEFT, padx=2)

    def create_threading_setting(self, parent):
        """Create threading setting controls"""
        threading_frame = tk.Frame(parent)
        threading_frame.pack(fill=tk.X, pady=5)

        # Threading info label
        cpu_info = f"CPU cores detected: {self.cpu_count}"
        tk.Label(threading_frame, text=f"Parallel threads: ({cpu_info})", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        thread_control_frame = tk.Frame(threading_frame)
        thread_control_frame.pack(fill=tk.X, pady=2)

        # Create label first
        self.thread_label = tk.Label(thread_control_frame, text=f"{self.max_threads.get()} threads",
                                     font=("Arial", 10), fg="darkgreen")

        # Dynamic max threads based on CPU count
        max_threads_limit = min(self.cpu_count * 2, 32)  # Allow up to 2x CPU cores, max 32

        thread_scale = tk.Scale(thread_control_frame, from_=1, to=max_threads_limit,
                                variable=self.max_threads, orient=tk.HORIZONTAL,
                                length=300, command=self.update_thread_label)
        thread_scale.pack(side=tk.LEFT)

        self.thread_label.pack(side=tk.LEFT, padx=(10, 0))

        # Thread presets
        thread_presets_frame = tk.Frame(threading_frame)
        thread_presets_frame.pack(fill=tk.X, pady=5)

        # Dynamic presets based on CPU count
        preset_single = 1
        preset_balanced = max(2, self.cpu_count // 2)
        preset_fast = max(4, self.cpu_count - 1)
        preset_max = min(self.cpu_count * 2, 16)

        tk.Button(thread_presets_frame, text=f"Single ({preset_single})",
                  command=lambda: self.set_threads(preset_single), bg="lightcoral").pack(side=tk.LEFT, padx=2)
        tk.Button(thread_presets_frame, text=f"Balanced ({preset_balanced})",
                  command=lambda: self.set_threads(preset_balanced), bg="lightyellow").pack(side=tk.LEFT, padx=2)
        tk.Button(thread_presets_frame, text=f"Fast ({preset_fast})",
                  command=lambda: self.set_threads(preset_fast), bg="lightgreen").pack(side=tk.LEFT, padx=2)
        tk.Button(thread_presets_frame, text=f"Maximum ({preset_max})",
                  command=lambda: self.set_threads(preset_max), bg="lightblue").pack(side=tk.LEFT, padx=2)

    def create_processing_options_frame(self, parent):
        """Create processing options UI components"""
        options_frame = tk.LabelFrame(parent, text="Processing Options", padx=10, pady=10)
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

    def create_progress_frame(self, parent):
        """Create progress UI components"""
        progress_frame = tk.LabelFrame(parent, text="Progress", padx=10, pady=10)
        progress_frame.pack(fill=tk.X, pady=10)

        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                            maximum=100, length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # Progress details label
        progress_details_label = tk.Label(progress_frame, textvariable=self.progress_details,
                                          font=("Arial", 9), fg="darkblue", wraplength=650)
        progress_details_label.pack(anchor=tk.W)

    def create_action_buttons_frame(self, parent):
        """Create action buttons UI components"""
        buttons_frame = tk.Frame(parent)
        buttons_frame.pack(fill=tk.X, pady=15)

        # Start compression button
        self.start_button = tk.Button(buttons_frame, text="Start Compression",
                                      command=self.start_compression,
                                      bg="green", fg="white", font=("Arial", 12, "bold"),
                                      height=2, width=15)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        # Reset button
        reset_button = tk.Button(buttons_frame, text="Reset Settings",
                                 command=self.reset_settings,
                                 bg="orange", fg="white", font=("Arial", 12, "bold"),
                                 height=2, width=15)
        reset_button.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(5, 0))

    def create_status_frame(self, parent):
        """Create status UI components"""
        status_frame = tk.LabelFrame(parent, text="Status", padx=10, pady=8)
        status_frame.pack(fill=tk.X, pady=8)

        status_label = tk.Label(status_frame, textvariable=self.status_text,
                                font=("Arial", 10), fg="blue", wraplength=650)
        status_label.pack(anchor=tk.W)

    def create_info_frame(self, parent):
        """Create info UI components"""
        info_frame = tk.LabelFrame(parent, text="Information", padx=10, pady=5)
        info_frame.pack(fill=tk.X, pady=5)

        info_text = tk.Text(info_frame, height=4, wrap=tk.WORD, font=("Arial", 9))
        info_text.pack(fill=tk.X)
        info_text.insert(tk.END,
                         "• Double-click on folder list to remove a folder\n"
                         "• Use Color Image DPI to control compression level (lower = smaller files)\n"
                         "• Set minimum file size to skip small files that don't need compression\n"
                         f"• Threads auto-detected: {self.cpu_count} CPU cores, max {min(self.cpu_count * 2, 32)} threads\n"
                         "• Import from Excel: First column should contain folder paths\n"
                         "• All logs are saved in the 'logs' folder next to the application")
        info_text.config(state=tk.DISABLED)

    # Event handlers and utility methods
    def update_dpi_label(self, value):
        """Update DPI label when scale changes"""
        dpi_val = int(float(value))
        self.dpi_label.config(text=f"{dpi_val} DPI")

    def update_size_label(self, value):
        """Update size label when scale changes"""
        size_val = float(value)
        self.size_label.config(text=f"{size_val:.1f} MB")

    def update_thread_label(self, value):
        """Update thread label when scale changes"""
        thread_val = int(float(value))
        self.thread_label.config(text=f"{thread_val} threads")

    def set_dpi(self, dpi_value):
        """Set DPI to specific value"""
        self.color_image_dpi.set(dpi_value)
        self.update_dpi_label(dpi_value)

    def set_min_size(self, size_value):
        """Set minimum file size to specific value"""
        self.min_file_size.set(size_value)
        self.update_size_label(size_value)

    def set_threads(self, thread_value):
        """Set number of threads to specific value"""
        self.max_threads.set(thread_value)
        self.update_thread_label(thread_value)

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
            if 'pd' in globals():
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

                if not path_obj.exists() or not path_obj.is_dir():
                    invalid_count += 1
                    continue

                if path_str in self.selected_folders:
                    skipped_count += 1
                    continue

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
        self.min_file_size.set(1.0)
        self.max_threads.set(min(max(2, self.cpu_count - 1), 16))
        self.status_text.set("Settings reset to defaults")
        self.progress_details.set("")
        self.progress_var.set(0)
        self.update_dpi_label(150)
        self.update_size_label(1.0)
        self.update_thread_label(self.max_threads.get())

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
        min_file_size = self.min_file_size.get()
        max_threads = self.max_threads.get()

        # Convert replace mode to boolean values
        replace_originals = replace_mode in ["replace_with_backup", "replace_without_backup"]
        create_backup = replace_mode == "replace_with_backup"

        # Show confirmation
        settings_summary = f"""Compression Settings:

Selected Folders: {len(self.selected_folders)} folder(s)
{chr(10).join([f"  • {Path(f).name}" for f in self.selected_folders])}

Color Image DPI: {color_image_dpi}
Minimum file size: {min_file_size:.1f} MB
Parallel threads: {max_threads}
Process subdirectories: {'Yes' if recursive else 'No'}
File handling: {
        'Keep originals (create new files)' if replace_mode == 'no_replace'
        else 'Replace originals with backup' if replace_mode == 'replace_with_backup'
        else 'Replace originals without backup'
        }

Do you want to start compression with these settings?"""

        if not messagebox.askyesno("Confirm Compression", settings_summary):
            return

        # Reset progress and start processing
        self.progress_var.set(0)
        self.progress_details.set("")
        self.stop_processing.clear()
        self.status_text.set("Starting compression...")
        self.start_button.config(state='disabled')
        self.root.update()

        # Start processing in separate thread
        self.processing_thread = threading.Thread(
            target=self.run_compression_threaded,
            args=(color_image_dpi, quality, recursive, replace_originals, create_backup, min_file_size, max_threads)
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def run_compression_threaded(self, color_image_dpi, quality, recursive, replace_originals, create_backup,
                                 min_file_size, max_threads):
        """Run compression in separate thread"""
        try:
            self.logger.info("\n" + "=" * 60)
            self.logger.info("BULK PDF COMPRESSION SESSION STARTED")
            self.logger.info("=" * 60)
            self.logger.info(f"Total folders to process: {len(self.selected_folders)}")
            self.logger.info(f"Settings: DPI={color_image_dpi}, Quality={quality}, Recursive={recursive}")
            self.logger.info(f"Minimum file size: {min_file_size:.1f} MB")
            self.logger.info(f"Parallel threads: {max_threads}")
            self.logger.info(f"Replace mode: {replace_originals}")

            # Process each folder
            total_processed = 0
            total_successful = 0
            total_failed = 0
            total_skipped = 0
            total_original_size = 0
            total_compressed_size = 0
            start_time = time.time()

            for i, folder in enumerate(self.selected_folders, 1):
                if self.stop_processing.is_set():
                    break

                folder_progress = (i - 1) / len(self.selected_folders) * 100
                self.root.after(0, lambda p=folder_progress: self.progress_var.set(p))
                self.root.after(0, lambda i=i, folder=folder: self.status_text.set(
                    f"Processing folder {i}/{len(self.selected_folders)}: {Path(folder).name}"))
                self.root.after(0, lambda folder=folder: self.progress_details.set(f"Folder: {folder}"))

                # Process folder with threading
                result = compress_all_pdfs_in_directory_threaded(
                    folder, color_image_dpi, quality, replace_originals, recursive, create_backup, min_file_size,
                    max_threads, self.logger
                )

                total_processed += 1
                total_successful += result.get('successful', 0)
                total_failed += result.get('failed', 0)
                total_skipped += result.get('skipped', 0)
                total_original_size += result.get('total_original_size', 0)
                total_compressed_size += result.get('total_compressed_size', 0)

            # Calculate processing time
            end_time = time.time()
            processing_time = end_time - start_time

            # Final update
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_text.set(
                f"Compression completed! Processed {total_processed} folder(s)"))
            self.root.after(0, lambda: self.progress_details.set(
                f"Total: {total_successful} successful, {total_failed} failed, {total_skipped} skipped"))

            # Final session summary
            self.logger.info("\n" + "=" * 60)
            self.logger.info("BULK COMPRESSION SESSION SUMMARY")
            self.logger.info("=" * 60)
            self.logger.info(f"Processing time: {processing_time:.1f} seconds")
            self.logger.info(f"Folders processed: {total_processed}")
            self.logger.info(f"Total files successful: {total_successful}")
            self.logger.info(f"Total files failed: {total_failed}")
            self.logger.info(f"Total files skipped: {total_skipped}")

            if total_successful > 0:
                overall_compression = (
                                                  1 - total_compressed_size / total_original_size) * 100 if total_original_size > 0 else 0
                files_per_second = total_successful / processing_time if processing_time > 0 else 0
                self.logger.info(f"Total original size: {total_original_size:.2f} MB")
                self.logger.info(f"Total compressed size: {total_compressed_size:.2f} MB")
                self.logger.info(f"Overall compression ratio: {overall_compression:.1f}%")
                self.logger.info(f"Processing speed: {files_per_second:.1f} files/second")

            self.logger.info("Session completed successfully")

            # Show final summary
            summary_msg = f"Compression completed!\n\n"
            summary_msg += f"Processing time: {processing_time:.1f} seconds\n"
            summary_msg += f"Folders processed: {total_processed}\n"
            summary_msg += f"Files successful: {total_successful}\n"
            if total_failed > 0:
                summary_msg += f"Files failed: {total_failed}\n"
            if total_skipped > 0:
                summary_msg += f"Files skipped (too small): {total_skipped}\n"
            if total_successful > 0:
                overall_compression = (
                                                  1 - total_compressed_size / total_original_size) * 100 if total_original_size > 0 else 0
                files_per_second = total_successful / processing_time if processing_time > 0 else 0
                summary_msg += f"Total space saved: {overall_compression:.1f}%\n"
                summary_msg += f"Speed: {files_per_second:.1f} files/second\n"
            summary_msg += f"\nDetailed logs saved in the 'logs' folder."

            self.root.after(0, lambda: messagebox.showinfo("Compression Complete", summary_msg))

        except Exception as e:
            error_msg = f"Compression failed: {str(e)}"
            self.logger.error(f"Application error: {e}")
            self.root.after(0, lambda: messagebox.showerror("Compression Error", error_msg))
            self.root.after(0, lambda: self.status_text.set("Compression failed!"))
            self.root.after(0, lambda: self.progress_details.set("Error occurred"))
        finally:
            # Re-enable controls
            self.root.after(0, lambda: self.start_button.config(state='normal'))


def main():
    """Main function with GUI"""
    root = tk.Tk()
    app = PDFCompressorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()