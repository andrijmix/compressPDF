import PyPDF2
import io
from PIL import Image
import fitz  # PyMuPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading


class PDFCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Light Compressor")
        self.root.geometry("700x700")
        self.root.resizable(True, True)

        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.compression_status = tk.StringVar(value="Ready to compress")

        # Compression settings
        self.jpeg_quality = tk.IntVar(value=85)
        self.scale_factor = tk.DoubleVar(value=1.0)
        self.compression_method = tk.StringVar(value="JPEG")
        self.convert_to_grayscale = tk.BooleanVar(value=False)
        self.remove_metadata = tk.BooleanVar(value=True)

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="PDF Light Compressor",
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Input file selection
        ttk.Label(main_frame, text="Input PDF File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file, width=50).grid(row=1, column=1,
                                                                           sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(main_frame, text="Browse",
                   command=self.browse_input_file).grid(row=1, column=2, padx=(5, 0))

        # Output file selection
        ttk.Label(main_frame, text="Output PDF File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file, width=50).grid(row=2, column=1,
                                                                            sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(main_frame, text="Browse",
                   command=self.browse_output_file).grid(row=2, column=2, padx=(5, 0))

        # Compression settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Compression Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        settings_frame.columnconfigure(1, weight=1)

        # JPEG Quality
        ttk.Label(settings_frame, text="JPEG Quality:").grid(row=0, column=0, sticky=tk.W, pady=2)
        quality_frame = ttk.Frame(settings_frame)
        quality_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        quality_frame.columnconfigure(0, weight=1)

        self.quality_scale = ttk.Scale(quality_frame, from_=10, to=100,
                                       variable=self.jpeg_quality, orient=tk.HORIZONTAL)
        self.quality_scale.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        self.quality_label = ttk.Label(quality_frame, text="85%")
        self.quality_label.grid(row=0, column=1)
        self.quality_scale.config(command=self.update_quality_label)

        # Scale Factor
        ttk.Label(settings_frame, text="Scale Factor:").grid(row=1, column=0, sticky=tk.W, pady=2)
        scale_frame = ttk.Frame(settings_frame)
        scale_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        scale_frame.columnconfigure(0, weight=1)

        self.scale_scale = ttk.Scale(scale_frame, from_=0.25, to=2.0,
                                     variable=self.scale_factor, orient=tk.HORIZONTAL)
        self.scale_scale.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        self.scale_label = ttk.Label(scale_frame, text="1.0x")
        self.scale_label.grid(row=0, column=1)
        self.scale_scale.config(command=self.update_scale_label)

        # Compression Method
        ttk.Label(settings_frame, text="Format:").grid(row=2, column=0, sticky=tk.W, pady=2)
        method_combo = ttk.Combobox(settings_frame, textvariable=self.compression_method,
                                    values=["JPEG", "PNG"], state="readonly", width=20)
        method_combo.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=2)

        # Checkboxes
        ttk.Checkbutton(settings_frame, text="Convert to Grayscale",
                        variable=self.convert_to_grayscale).grid(row=3, column=0, columnspan=2,
                                                                 sticky=tk.W, pady=2)

        ttk.Checkbutton(settings_frame, text="Remove Metadata",
                        variable=self.remove_metadata).grid(row=4, column=0, columnspan=2,
                                                            sticky=tk.W, pady=2)

        # Preset buttons
        preset_frame = ttk.Frame(settings_frame)
        preset_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(preset_frame, text="Maximum Quality",
                   command=self.set_max_quality).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="Balanced",
                   command=self.set_balanced).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="Maximum Compression",
                   command=self.set_max_compression).pack(side=tk.LEFT, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)

        # Status label
        status_label = ttk.Label(main_frame, textvariable=self.compression_status)
        status_label.grid(row=5, column=0, columnspan=3, pady=5)

        # Compress button
        self.compress_btn = ttk.Button(main_frame, text="Compress PDF",
                                       command=self.start_compression)
        self.compress_btn.grid(row=6, column=0, columnspan=3, pady=20)

        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Compression Results", padding="10")
        results_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        results_frame.columnconfigure(1, weight=1)

        # File size labels
        self.original_size_label = ttk.Label(results_frame, text="Original size: --")
        self.original_size_label.grid(row=0, column=0, sticky=tk.W, pady=2)

        self.compressed_size_label = ttk.Label(results_frame, text="Compressed size: --")
        self.compressed_size_label.grid(row=1, column=0, sticky=tk.W, pady=2)

        self.compression_ratio_label = ttk.Label(results_frame, text="Compression ratio: --")
        self.compression_ratio_label.grid(row=2, column=0, sticky=tk.W, pady=2)

    def update_quality_label(self, value):
        self.quality_label.config(text=f"{int(float(value))}%")

    def update_scale_label(self, value):
        self.scale_label.config(text=f"{float(value):.2f}x")

    def set_max_quality(self):
        self.jpeg_quality.set(95)
        self.scale_factor.set(1.0)
        self.compression_method.set("PNG")
        self.convert_to_grayscale.set(False)
        self.remove_metadata.set(False)
        self.update_quality_label(95)
        self.update_scale_label(1.0)

    def set_balanced(self):
        self.jpeg_quality.set(75)
        self.scale_factor.set(0.85)
        self.compression_method.set("JPEG")
        self.convert_to_grayscale.set(False)
        self.remove_metadata.set(True)
        self.update_quality_label(75)
        self.update_scale_label(0.85)

    def set_max_compression(self):
        self.jpeg_quality.set(40)
        self.scale_factor.set(0.6)
        self.compression_method.set("JPEG")
        self.convert_to_grayscale.set(True)
        self.remove_metadata.set(True)
        self.update_quality_label(40)
        self.update_scale_label(0.6)

    def browse_input_file(self):
        try:
            filename = filedialog.askopenfilename(
                title="Select PDF file to compress",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if filename:
                self.input_file.set(filename)
                # Auto-generate output filename
                if not self.output_file.get():
                    base_name = os.path.splitext(filename)[0]
                    output_name = f"{base_name}_compressed.pdf"
                    self.output_file.set(output_name)
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting input file: {str(e)}")

    def browse_output_file(self):
        try:
            filename = filedialog.asksaveasfilename(
                title="Save compressed PDF as",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if filename:
                self.output_file.set(filename)
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting output file: {str(e)}")

    def start_compression(self):
        if not self.input_file.get() or not self.output_file.get():
            messagebox.showwarning("Warning", "Please select both input and output files")
            return

        # Disable button and start progress
        self.compress_btn.config(state='disabled')
        self.progress.start(10)
        self.compression_status.set("Compressing PDF...")

        # Start compression in separate thread
        thread = threading.Thread(target=self.compress_pdf_thread)
        thread.daemon = True
        thread.start()

    def compress_pdf_thread(self):
        try:
            original_size = self.get_file_size(self.input_file.get())

            # Perform compression
            self.compress_with_settings(self.input_file.get(), self.output_file.get())

            compressed_size = self.get_file_size(self.output_file.get())
            compression_ratio = self.get_compression_ratio(original_size, compressed_size)

            # Update UI in main thread
            self.root.after(0, self.compression_completed, original_size, compressed_size, compression_ratio)

        except FileNotFoundError as e:
            self.root.after(0, self.compression_failed, f"File not found: {str(e)}")
        except PermissionError as e:
            self.root.after(0, self.compression_failed, f"Permission denied: {str(e)}")
        except Exception as e:
            self.root.after(0, self.compression_failed, f"Compression failed: {str(e)}")

    def compression_completed(self, original_size, compressed_size, compression_ratio):
        # Stop progress and enable button
        self.progress.stop()
        self.compress_btn.config(state='normal')
        self.compression_status.set("Compression completed successfully!")

        # Update results
        self.original_size_label.config(text=f"Original size: {original_size:.2f} MB")
        self.compressed_size_label.config(text=f"Compressed size: {compressed_size:.2f} MB")
        self.compression_ratio_label.config(text=f"Compression ratio: {compression_ratio:.1f}%")

        messagebox.showinfo("Success",
                            f"PDF compressed successfully!\n"
                            f"Original: {original_size:.2f} MB\n"
                            f"Compressed: {compressed_size:.2f} MB\n"
                            f"Saved: {compression_ratio:.1f}%")

    def compression_failed(self, error_message):
        # Stop progress and enable button
        self.progress.stop()
        self.compress_btn.config(state='normal')
        self.compression_status.set("Compression failed!")

        messagebox.showerror("Error", error_message)

    def compress_with_settings(self, input_pdf, output_pdf):
        """Compress PDF with user-defined settings"""
        doc = None
        new_doc = None

        try:
            doc = fitz.open(input_pdf)
            new_doc = fitz.open()

            # Get settings
            quality = self.jpeg_quality.get()
            scale = self.scale_factor.get()
            image_format = self.compression_method.get()
            grayscale = self.convert_to_grayscale.get()
            remove_meta = self.remove_metadata.get()

            for page_num in range(len(doc)):
                page = doc[page_num]

                # Apply scale factor
                mat = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=mat)

                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))

                # Convert to grayscale if requested
                if grayscale:
                    img = img.convert('L')
                elif img.mode != 'RGB':
                    img = img.convert('RGB')

                # Compress image
                img_bytes = io.BytesIO()
                if image_format == "JPEG":
                    img.save(img_bytes, format='JPEG', quality=quality, optimize=True)
                else:  # PNG
                    img.save(img_bytes, format='PNG', optimize=True)
                img_bytes.seek(0)

                new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=img_bytes.getvalue())

            # Save with compression options
            if remove_meta:
                new_doc.set_metadata({})

            new_doc.save(output_pdf, deflate=True, clean=True, garbage=4)

        except Exception as e:
            raise Exception(f"Error during PDF processing: {str(e)}")

        finally:
            if doc:
                doc.close()
            if new_doc:
                new_doc.close()

    def compress_with_settings(self, input_pdf, output_pdf):
        """Compress PDF with user-defined settings"""
        doc = None
        new_doc = None

        try:
            doc = fitz.open(input_pdf)
            new_doc = fitz.open()

            # Get settings
            quality = self.jpeg_quality.get()
            scale = self.scale_factor.get()
            image_format = self.compression_method.get()
            grayscale = self.convert_to_grayscale.get()
            remove_meta = self.remove_metadata.get()

            for page_num in range(len(doc)):
                page = doc[page_num]

                # Apply scale factor
                mat = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=mat)

                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))

                # Convert to grayscale if requested
                if grayscale:
                    img = img.convert('L')
                elif img.mode != 'RGB':
                    img = img.convert('RGB')

                # Compress image
                img_bytes = io.BytesIO()
                if image_format == "JPEG":
                    img.save(img_bytes, format='JPEG', quality=quality, optimize=True)
                else:  # PNG
                    img.save(img_bytes, format='PNG', optimize=True)
                img_bytes.seek(0)

                new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=img_bytes.getvalue())

            # Save with compression options
            if remove_meta:
                new_doc.set_metadata({})

            new_doc.save(output_pdf, deflate=True, clean=True, garbage=4)

        except Exception as e:
            raise Exception(f"Error during PDF processing: {str(e)}")

        finally:
            if doc:
                doc.close()
            if new_doc:
                new_doc.close()

    def compress_light(self, input_pdf, output_pdf):
        """Legacy light compression method"""
        doc = None
        new_doc = None

        try:
            doc = fitz.open(input_pdf)
            new_doc = fitz.open()

            for page_num in range(len(doc)):
                page = doc[page_num]

                # Original resolution
                pix = page.get_pixmap()

                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))

                if img.mode != 'RGB':
                    img = img.convert('RGB')

                # High quality compression
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='JPEG', quality=85, optimize=True)
                img_bytes.seek(0)

                new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=img_bytes.getvalue())

            new_doc.save(output_pdf, deflate=True, clean=True, garbage=2)

        except Exception as e:
            raise Exception(f"Error during PDF processing: {str(e)}")

        finally:
            if doc:
                doc.close()
            if new_doc:
                new_doc.close()

    def get_file_size(self, file_path):
        """Get file size in MB"""
        try:
            return os.path.getsize(file_path) / (1024 * 1024)
        except OSError as e:
            raise Exception(f"Cannot get file size: {str(e)}")

    def get_compression_ratio(self, original_size, compressed_size):
        """Calculate compression ratio"""
        if original_size == 0:
            return 0
        return (1 - compressed_size / original_size) * 100


def main():
    try:
        root = tk.Tk()
        app = PDFCompressorApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Application error: {str(e)}")
        messagebox.showerror("Critical Error", f"Application failed to start: {str(e)}")


if __name__ == "__main__":
    main()