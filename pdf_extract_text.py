import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import pandas as pd


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Address Block Extractor")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        self.pdf_path = None
        self.directory_path = None
        self.rect_start = None
        self.rect_end = None
        self.canvas = None
        self.rect_id = None
        self.image = None
        self.extracted_data = []
        self.is_directory_mode = False

        self.setup_gui()

    def setup_gui(self):
        # Configure styles
        style = ttk.Style()
        style.configure('TButton', font=('Arial', 12), padding=10)
        style.configure('TLabel', font=('Arial', 12))
        style.configure('TEntry', font=('Arial', 12))

        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # PDF selection
        ttk.Button(main_frame, text="Select Single PDF", command=self.select_pdf).grid(row=0, column=0, pady=10,
                                                                                       sticky=tk.W)
        ttk.Button(main_frame, text="Select Directory of PDFs", command=self.select_directory).grid(row=0, column=1,
                                                                                                    pady=10,
                                                                                                    sticky=tk.W)

        # Canvas for PDF display
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.canvas = tk.Canvas(canvas_frame, width=600, height=800, background="white")
        self.canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

        # Page selection (only shown for single PDF mode)
        self.page_label = ttk.Label(main_frame, text="Extract every nth page:")
        self.page_label.grid(row=2, column=0, pady=5, sticky=tk.W)
        self.page_entry = ttk.Entry(main_frame)
        self.page_entry.grid(row=3, column=0, pady=5, sticky=tk.W)
        self.page_entry.insert(0, "1")

        # Extract and Save
        ttk.Button(main_frame, text="Extract and Save", command=self.extract_and_save).grid(row=4, column=0,
                                                                                            columnspan=2, pady=20,
                                                                                            sticky=tk.S)

        # Make sure the grid expands properly
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=0)

    def select_pdf(self):
        # Clear the previously extracted data
        self.extracted_data.clear()
        self.is_directory_mode = False

        # Reset the canvas and rectangle selection
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.canvas.delete("all")

        # Ensure page selection options are visible
        self.page_label.grid()
        self.page_entry.grid()

        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.display_first_page()

    def select_directory(self):
        # Clear the previously extracted data
        self.extracted_data.clear()
        self.is_directory_mode = True

        # Reset the canvas and rectangle selection
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.canvas.delete("all")

        # Hide the page selection options
        self.page_label.grid_remove()
        self.page_entry.grid_remove()

        self.directory_path = filedialog.askdirectory()
        if self.directory_path:
            messagebox.showinfo("Info", "Please select a sample PDF to define the extraction area.")
            self.select_sample_pdf_for_directory()

    def select_sample_pdf_for_directory(self):
        # This function is to select a sample PDF and define the extraction area
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            self.display_first_page()

    def display_first_page(self):
        if not self.pdf_path:
            return

        doc = fitz.open(self.pdf_path)
        page = doc[0]  # Always display the first page
        pix = page.get_pixmap()

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.image = ImageTk.PhotoImage(img)

        # Resize the canvas and window to fit the PDF page size
        self.canvas.config(width=pix.width, height=pix.height)
        self.canvas.create_image(0, 0, anchor="nw", image=self.image)

        # Ensure the window adjusts to the new content
        self.root.update_idletasks()

    def on_button_press(self, event):
        self.rect_start = (event.x, event.y)
        if self.rect_id:
            self.canvas.delete(self.rect_id)

    def on_mouse_drag(self, event):
        if self.rect_start:
            if self.rect_id:
                self.canvas.delete(self.rect_id)
            self.rect_id = self.canvas.create_rectangle(
                self.rect_start[0], self.rect_start[1], event.x, event.y, outline="red"
            )

    def on_button_release(self, event):
        self.rect_end = (event.x, event.y)

    def extract_and_save(self):
        if not self.pdf_path or not self.rect_start or not self.rect_end:
            messagebox.showerror("Error", "Please select a PDF and draw a rectangle.")
            return

        rect = fitz.Rect(
            min(self.rect_start[0], self.rect_end[0]),
            min(self.rect_start[1], self.rect_end[1]),
            max(self.rect_start[0], self.rect_end[0]),
            max(self.rect_start[1], self.rect_end[1]),
        )

        if self.is_directory_mode:
            self.extract_from_directory(rect)
        else:
            self.extract_from_single_pdf(rect)

    def extract_from_single_pdf(self, rect):
        try:
            n = int(self.page_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid number for n pages.")
            return

        doc = fitz.open(self.pdf_path)

        for i in range(0, len(doc), n):
            page = doc[i]
            text = page.get_text("text", clip=rect).strip()
            if text:
                self.extracted_data.append(text.splitlines())

        if self.extracted_data:
            df = pd.DataFrame(self.extracted_data)
            save_path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
            )
            if save_path:
                df.to_csv(save_path, index=False, header=False)
                messagebox.showinfo("Success", "Data saved successfully.")
                self.extracted_data.clear()
        else:
            messagebox.showinfo("Info", "No text extracted.")

    def extract_from_directory(self, rect):
        # Extract from the first page of each PDF in the directory
        for root, _, files in os.walk(self.directory_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(root, file)
                    doc = fitz.open(pdf_path)
                    page = doc[0]  # Always extract from the first page

                    text = page.get_text("text", clip=rect).strip()
                    if text:
                        self.extracted_data.append([pdf_path] + text.splitlines())

        if self.extracted_data:
            # Determine the maximum number of columns needed
            max_columns = max(len(row) for row in self.extracted_data)

            # Generate appropriate headers
            headers = ["File Path"] + [f"Line {i + 1}" for i in range(max_columns - 1)]

            # Create the DataFrame
            df = pd.DataFrame(self.extracted_data)

            # Ensure all rows have the same number of columns by filling missing values with empty strings
            df = df.reindex(columns=range(max_columns)).fillna("")

            # Save to CSV
            save_path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
            )
            if save_path:
                df.to_csv(save_path, index=False, header=headers)
                messagebox.showinfo("Success", "Data saved successfully.")
                self.extracted_data.clear()
        else:
            messagebox.showinfo("Info", "No text extracted.")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
