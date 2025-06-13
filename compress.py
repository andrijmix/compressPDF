import os
import sys
import subprocess
import logging
import shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


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


def compress_pdf(input_file, output_file, dpi=150, quality="ebook", logger=None):
    """Compress a single PDF file"""
    if logger is None:
        logger = logging.getLogger(__name__)

    gs_path = get_gs_executable_path()

    try:
        logger.info(f"Starting compression: {input_file}")
        logger.info(f"Output file: {output_file}")
        logger.info(f"Settings - DPI: {dpi}, Quality: {quality}")

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
            f"-r{dpi}",
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


def compress_all_pdfs_in_directory(directory, dpi=150, quality="ebook", replace_originals=False, recursive=True):
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
        logger.info(f"DPI setting: {dpi}")
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
                result = compress_pdf(str(pdf), str(output_path), dpi, quality, logger)

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


def main():
    """Main function with GUI"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    try:
        # Get folder
        folder = filedialog.askdirectory(title="Select folder with PDFs to compress")
        if not folder:
            return

        # Ask about recursive processing
        recursive = messagebox.askyesno(
            "Process Subdirectories?",
            "Do you want to process PDFs in subdirectories as well?\n\n"
            "Yes: Process all PDFs in the selected folder and all its subdirectories\n"
            "No: Process only PDFs in the selected folder"
        )

        # Get DPI setting
        dpi = simpledialog.askinteger(
            "DPI Setting",
            "Enter DPI (72=low quality, 150=medium, 300=high):",
            initialvalue=150,
            minvalue=72,
            maxvalue=600
        )
        if dpi is None:
            return

        # Ask about replacing originals
        replace_originals = messagebox.askyesno(
            "Replace Original Files?",
            "Do you want to replace the original files with compressed versions?\n\n"
            "Yes: Original files will be backed up and replaced\n"
            "No: Compressed files will be saved to a new folder"
        )


        quality = "screen"  # Default quality setting

        print("Please wait while the PDFs are being compressed...")
        print("Check the console and log files for detailed progress...")

        compress_all_pdfs_in_directory(folder, dpi, quality, replace_originals, recursive)

    except Exception as e:
        logging.error(f"Application error: {e}")
        messagebox.showerror("Unexpected Error", str(e))


if __name__ == "__main__":
    main()