import sys

print(sys.executable)
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
import io
import re
import os
import sys
import webbrowser
import urllib.parse
import threading



def get_tesseract_path():
    # Running as a PyInstaller executable
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
        return os.path.join(base_path, "_internal", "Tesseract-OCR", "tesseract.exe")

    # Running from source code
    return os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "Tesseract-OCR",
        "tesseract.exe"
    )

pytesseract.pytesseract.tesseract_cmd = get_tesseract_path()

class UniversalExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Directory & Table Tool")
        self.root.geometry("750x780")

        # Create Tabbed Interface
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')

        self.tab_extract = ttk.Frame(self.notebook)
        self.tab_classify = ttk.Frame(self.notebook)
        self.tab_filter = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_extract, text="1. Extract PDF")
        self.notebook.add(self.tab_classify, text="2. Classify Data")
        self.notebook.add(self.tab_filter, text="3. Filter & Export")

        # Shared Filter Variables
        self.filter_in_var = tk.StringVar()
        self.filter_col_var = tk.StringVar()
        self.filter_kw_var = tk.StringVar(value="Solar, Energy, Panel")

        # Classifier State
        self.df = None
        self.current_index = 0
        self.output_file = ""

        # Simplified Custom Columns Variable
        self.custom_cols_var = tk.StringVar(value="Name of the Company, CIN, Equity ISIN, Name of the Group, Sector")

        self.setup_extract_tab()
        self.setup_classify_tab()
        self.setup_filter_tab()

    # ==========================================
    #             1. SIMPLIFIED EXTRACTION TAB
    # ==========================================
    def setup_extract_tab(self):
        # --- 1. File Settings ---
        file_frame = ttk.LabelFrame(self.tab_extract, text="1. File Settings", padding=10)
        file_frame.pack(padx=10, pady=5, fill="x")

        ttk.Label(file_frame, text="PDF File:").grid(row=0, column=0, sticky="w", pady=2)
        self.pdf_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.pdf_path_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(file_frame, text="Save Excel As:").grid(row=1, column=0, sticky="w", pady=2)
        self.extract_out_var = tk.StringVar(value="Extracted_PDF_Data.xlsx")
        ttk.Entry(file_frame, textvariable=self.extract_out_var, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_extract_out).grid(row=1, column=2)

        page_frame = ttk.Frame(file_frame)
        page_frame.grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(page_frame, text="Start Page:").pack(side="left")
        self.start_page_var = tk.IntVar(value=1)
        ttk.Entry(page_frame, textvariable=self.start_page_var, width=5).pack(side="left", padx=5)
        ttk.Label(page_frame, text="End Page:").pack(side="left")
        self.end_page_var = tk.IntVar(value=10)
        ttk.Entry(page_frame, textvariable=self.end_page_var, width=5).pack(side="left", padx=5)

        # --- 2. Smart Settings ---
        settings_frame = ttk.LabelFrame(self.tab_extract, text="2. Smart Extraction Settings", padding=10)
        settings_frame.pack(padx=10, pady=5, fill="x")

        # Document Format Dropdown
        ttk.Label(settings_frame, text="PDF Layout Type:").grid(row=0, column=0, sticky="w", pady=5)
        self.layout_mode_var = tk.StringVar(value="Structured Table / Grid")
        layout_dropdown = ttk.Combobox(settings_frame, textvariable=self.layout_mode_var, state="readonly", width=35)
        layout_dropdown['values'] = ("Structured Table / Grid", "Unstructured Directory (Blocks)")
        layout_dropdown.grid(row=0, column=1, sticky="w", padx=5)
        layout_dropdown.bind("<<ComboboxSelected>>", self.on_layout_mode_change)

        # Fast Text vs. OCR Dropdown
        ttk.Label(settings_frame, text="Extraction Method:").grid(row=1, column=0, sticky="w", pady=5)
        self.method_var = tk.StringVar(value="Fast Digital (Selectable Text)")
        method_dropdown = ttk.Combobox(settings_frame, textvariable=self.method_var, state="readonly", width=35)
        method_dropdown['values'] = ("Fast Digital (Selectable Text)", "OCR Scan (Scanned PDF / Image)")
        method_dropdown.grid(row=1, column=1, sticky="w", padx=5)
        method_dropdown.bind("<<ComboboxSelected>>", self.toggle_tesseract_warning)

        # Collapsible Advanced Custom Columns Checkbox
        self.show_adv_var = tk.BooleanVar(value=False)
        self.adv_chk = ttk.Checkbutton(settings_frame, text="Specify Columns to Extract (Advanced)",
                                       variable=self.show_adv_var, command=self.toggle_advanced_frame)
        self.adv_chk.grid(row=2, column=1, sticky="w", pady=5)

        # Hidden Tesseract path
        self.tesseract_path_var = tk.StringVar(value=get_tesseract_path())

        # Tesseract Warning Message Label
        self.tesseract_warn_label = ttk.Label(settings_frame, text="", foreground="orange", font=("Arial", 9, "italic"))
        self.tesseract_warn_label.grid(row=3, column=1, sticky="w")

        # --- 3. Advanced Columns Frame (Hidden by Default) ---
        self.adv_frame = ttk.LabelFrame(self.tab_extract, text="Advanced Column Customization", padding=10)

        ttk.Label(self.adv_frame, text="Columns to Keep / Rename (Comma Separated):", font=("Arial", 9, "bold")).pack(
            anchor="w", pady=2)
        ttk.Entry(self.adv_frame, textvariable=self.custom_cols_var, width=75).pack(fill="x", pady=2)

        help_text = (
            "• Table Mode: Only columns matching these header names will be extracted from the PDF.\n"
            "• Directory Mode: Renames the standard fields (Company, Address, Contact, Description) to these names."
        )
        ttk.Label(self.adv_frame, text=help_text, justify="left", foreground="gray", font=("Arial", 9, "italic")).pack(
            anchor="w", pady=5)

        # --- 4. Execution Control ---
        run_frame = ttk.Frame(self.tab_extract)
        run_frame.pack(pady=15)

        self.extract_status_var = tk.StringVar(value="Ready.")
        ttk.Label(run_frame, textvariable=self.extract_status_var, font=("Arial", 10, "bold")).pack(pady=5)

        self.btn_extract = ttk.Button(run_frame, text="Run Extractor", command=self.start_extraction_thread)
        self.btn_extract.pack(pady=5)

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path_var.set(path)

    def browse_extract_out(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.extract_out_var.set(path)

    def on_layout_mode_change(self, event=None):
        mode = self.layout_mode_var.get()
        if "Structured Table" in mode:
            self.custom_cols_var.set("Name of the Company, CIN, Equity ISIN, Name of the Group, Sector")
        else:
            self.custom_cols_var.set("Company Name, Location/Address, Phone Extracted, Description/Products")

    def toggle_advanced_frame(self):
        if self.show_adv_var.get():
            self.adv_frame.pack(padx=10, pady=5, fill="x")
        else:
            self.adv_frame.pack_forget()

    def toggle_tesseract_warning(self, event=None):
        if "OCR Scan" in self.method_var.get():
            self.tesseract_warn_label.config(text=f"* Requires local OCR engine setup.")
        else:
            self.tesseract_warn_label.config(text="")

    def get_keywords(self, var):
        return [k.strip().upper() for k in var.get().split(',') if k.strip()]

    def filter_table_by_columns(self, all_rows, desired_cols):
        if not all_rows or not desired_cols:
            return pd.DataFrame(all_rows)

        # Clean rows of empty rows
        cleaned_rows = [r for r in all_rows if r and any(x is not None and str(x).strip() != "" for x in r)]
        if not cleaned_rows:
            return pd.DataFrame(all_rows)

        header_row_idx = 0
        matched_indices = {}
        for d_col in desired_cols:
            d_col_clean = d_col.strip().lower()
            if not d_col_clean:
                continue
            best_idx = None
            for idx, candidate in enumerate(cleaned_rows[header_row_idx]):
                if candidate is None:
                    continue
                cand_clean = str(candidate).strip().lower()
                if d_col_clean in cand_clean or cand_clean in d_col_clean:
                    best_idx = idx
                    break
            if best_idx is not None:
                matched_indices[d_col.strip()] = best_idx

        if not matched_indices:
            # Fall back to returning all found columns if no headings match
            return pd.DataFrame(cleaned_rows)

        new_rows = []
        headers = list(matched_indices.keys())
        new_rows.append(headers)

        for row in cleaned_rows[header_row_idx + 1:]:
            new_row = []
            for col_name in headers:
                idx = matched_indices[col_name]
                if idx < len(row):
                    new_row.append(row[idx])
                else:
                    new_row.append("")
            new_rows.append(new_row)

        df = pd.DataFrame(new_rows[1:], columns=new_rows[0])
        return df

    def start_extraction_thread(self):
        if not self.pdf_path_var.get():
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        if "OCR Scan" in self.method_var.get():
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path_var.get()
            if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
                messagebox.showerror("Error",
                                     f"Tesseract OCR missing. For scanned documents, place Tesseract-OCR folder next to the app, or use Fast Digital mode.")
                return

        self.btn_extract.config(state="disabled")
        self.extract_status_var.set("Initializing automated engine...")
        threading.Thread(target=self.run_extraction, daemon=True).start()

    def run_extraction(self):
        try:
            pdf_path = self.pdf_path_var.get()
            out_path = self.extract_out_var.get()
            start_page = self.start_page_var.get()
            end_page = self.end_page_var.get()
            layout_mode = self.layout_mode_var.get()
            is_ocr = "OCR Scan" in self.method_var.get()

            doc = fitz.open(pdf_path)
            total_pages = doc.page_count
            start_page = max(1, start_page)
            end_page = min(total_pages, end_page)

            # ====================================================
            # CASE A: STRUCTURED TABLE / GRID EXTRACTION (Auto-Detect)
            # ====================================================
            if "Structured Table / Grid" in layout_mode:
                self.extract_status_var.set("Scanning document layout for tables...")
                all_table_rows = []

                for page_num in range(start_page - 1, end_page):
                    self.extract_status_var.set(f"Extracting page grid {page_num + 1} of {end_page}...")
                    page = doc.load_page(page_num)

                    if is_ocr:
                        pix = page.get_pixmap(dpi=300)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        raw_text = pytesseract.image_to_string(img)
                        lines = [line.strip().split() for line in raw_text.split('\n') if line.strip()]
                        all_table_rows.extend(lines)
                    else:
                        tables = page.find_tables()
                        for table in tables:
                            rows = table.extract()
                            if rows:
                                all_table_rows.extend(rows)

                if not all_table_rows:
                    self.extract_status_var.set("No grid lines found. Attempting digital word layout fallback...")
                    for page_num in range(start_page - 1, end_page):
                        page = doc.load_page(page_num)
                        blocks = page.get_text("blocks")
                        blocks.sort(key=lambda b: (b[1], b[0]))
                        for b in blocks:
                            clean_text = b[4].strip()
                            if clean_text:
                                all_table_rows.append([clean_text])

                # Check if custom columns list is enabled
                if self.show_adv_var.get() and self.custom_cols_var.get().strip():
                    desired_cols = [c.strip() for c in self.custom_cols_var.get().split(",") if c.strip()]
                    df = self.filter_table_by_columns(all_table_rows, desired_cols)
                else:
                    df = pd.DataFrame(all_table_rows)
                    df.dropna(how='all', inplace=True)

                df.to_excel(out_path, index=False, header=not df.columns.astype(str).str.isdigit().all())
                self.extract_status_var.set(f"Success! Extracted {len(df)} lines.")
                messagebox.showinfo("Success", f"Extraction complete! Saved to {os.path.basename(out_path)}")

            # ====================================================
            # CASE B: UNSTRUCTURED DIRECTORY EXTRACTION (Blocks)
            # ====================================================
            else:
                # Built-in robust directory block keyword match parameters
                start_kws = ["HALL", "BOOTH", "EXHIBITOR", "COMPANY", "STAND", "SR.NO", "SR. NO."]
                addr_kws = ["STALL", "STAND", "ADDRESS", "ROAD", "STREET", "PLOT", "LOCATION", "CITY", "STATE", "ZIP"]
                contact_kws = ["TEL", "MOBILE", "E-MAIL", "EMAIL", "WEBSITE", "WWW", "CONTACT PERSON", "CONTACT"]
                prod_kws = ["PRODUCTS ON DISPLAY", "PRODUCTS", "DESCRIPTION", "PROFILE", "SERVICES", "SECTOR"]

                companies = []
                current_company = {}
                company_name_buffer, address_buffer, products_buffer = [], [], []
                phone = ""
                capture_state = "SEARCHING_COMPANY"

                for page_num in range(start_page - 1, end_page):
                    self.extract_status_var.set(f"Parsing block page {page_num + 1} of {end_page}...")
                    page = doc.load_page(page_num)

                    if is_ocr:
                        pix = page.get_pixmap(dpi=300)
                        img = Image.open(io.BytesIO(pix.tobytes("png")))
                        raw_text = pytesseract.image_to_string(img)
                    else:
                        raw_text = page.get_text("text")

                    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]

                    for line in lines:
                        upper_line = line.upper()
                        is_new_company = False

                        has_start = any(k in upper_line for k in start_kws)
                        has_addr = any(k in upper_line for k in addr_kws)
                        has_contact = any(k in upper_line for k in contact_kws)
                        has_prod = any(k in upper_line for k in prod_kws)

                        if has_start and not has_addr:
                            is_new_company = True
                        elif (
                                len(line) > 5 and line.isupper() and not has_prod and not has_contact and not has_addr and not has_start):
                            if capture_state not in ["SEARCHING_COMPANY", "READING_COMPANY_NAME"]:
                                is_new_company = True

                        if is_new_company:
                            if company_name_buffer:
                                current_company["Company Name"] = " ".join(company_name_buffer)
                                current_company["Location/Address"] = " ".join(address_buffer)
                                current_company["Phone Extracted"] = phone
                                current_company["Description/Products"] = " ".join(products_buffer)
                                companies.append(current_company)

                            current_company = {}
                            matched_k = next((k for k in start_kws if k in upper_line), None)
                            if matched_k:
                                part = re.split(matched_k, line, flags=re.IGNORECASE)[0].strip()
                                company_name_buffer = [part] if part else []
                            else:
                                company_name_buffer = [line]

                            address_buffer, products_buffer = [], []
                            phone = ""
                            capture_state = "READING_COMPANY_NAME"
                            continue

                        # State Machine Parsing Heuristic
                        if capture_state == "READING_COMPANY_NAME":
                            if has_addr:
                                matched_k = next((k for k in addr_kws if k in upper_line), None)
                                if matched_k:
                                    name_part = re.split(matched_k, line, flags=re.IGNORECASE)[0].strip()
                                    name_part = re.sub(r'[^a-zA-Z0-9\s]', '', name_part).strip()
                                    if len(name_part) > 1 and name_part.isupper():
                                        company_name_buffer.append(name_part)
                                capture_state = "READING_ADDRESS"
                            elif line.isupper() and len(line) > 2 and not has_start:
                                company_name_buffer.append(line.strip())
                            else:
                                capture_state = "READING_ADDRESS"
                                address_buffer.append(line)

                        elif capture_state == "READING_ADDRESS":
                            if has_addr:
                                continue
                            elif has_contact:
                                capture_state = "SEARCHING_PRODUCTS"
                                nums = re.findall(r'\+?\d[\d\s\-]{7,}\d', line)
                                if nums:
                                    phone = nums[0].strip()
                            else:
                                address_buffer.append(line)

                        elif capture_state == "SEARCHING_PRODUCTS":
                            if has_contact:
                                nums = re.findall(r'\+?\d[\d\s\-]{7,}\d', line)
                                if nums and not phone:
                                    phone = nums[0].strip()
                            elif has_prod:
                                parts = line.split(":", 1)
                                if len(parts) > 1:
                                    products_buffer.append(parts[1].strip())
                                capture_state = "READING_PRODUCTS"

                        elif capture_state == "READING_PRODUCTS":
                            products_buffer.append(line)

                # Save last company in iteration
                if company_name_buffer:
                    current_company["Company Name"] = " ".join(company_name_buffer)
                    current_company["Location/Address"] = " ".join(address_buffer)
                    current_company["Phone Extracted"] = phone
                    current_company["Description/Products"] = " ".join(products_buffer)
                    companies.append(current_company)

                self.extract_status_var.set("Saving directory to Excel...")
                df = pd.DataFrame(companies)

                # Map to custom column names if customized
                if self.show_adv_var.get() and self.custom_cols_var.get().strip():
                    desired_cols = [c.strip() for c in self.custom_cols_var.get().split(",") if c.strip()]
                    rename_dict = {}
                    standard_cols = ["Company Name", "Location/Address", "Phone Extracted", "Description/Products"]
                    for i, col in enumerate(standard_cols):
                        if i < len(desired_cols):
                            rename_dict[col] = desired_cols[i]
                    df.rename(columns=rename_dict, inplace=True)
                    # Keep only the requested columns
                    keep_cols = [desired_cols[i] for i in range(min(len(standard_cols), len(desired_cols)))]
                    if keep_cols:
                        df = df[keep_cols]

                df.to_excel(out_path, index=False)
                self.extract_status_var.set(f"Success! Extracted {len(df)} companies.")
                messagebox.showinfo("Success",
                                    f"Extraction complete! Saved {len(df)} directory items to {os.path.basename(out_path)}")

        except Exception as e:
            self.extract_status_var.set("Error occurred.")
            messagebox.showerror("System Error", str(e))
        finally:
            self.btn_extract.config(state="normal")

    # ==========================================
    #            2. CLASSIFICATION TAB
    # ==========================================
    def setup_classify_tab(self):
        file_frame = ttk.LabelFrame(self.tab_classify, text="1. Load Data", padding=10)
        file_frame.pack(padx=10, pady=5, fill="x")

        ttk.Label(file_frame, text="Input Excel:").grid(row=0, column=0, sticky="w")
        self.class_in_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.class_in_var, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=lambda: self.class_in_var.set(filedialog.askopenfilename())).grid(
            row=0, column=2)

        ttk.Label(file_frame, text="Save/Master Excel:").grid(row=1, column=0, sticky="w", pady=5)
        self.class_out_var = tk.StringVar(value="Master_Classified_Data.xlsx")
        ttk.Entry(file_frame, textvariable=self.class_out_var, width=40).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=lambda: self.class_out_var.set(
            filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]))).grid(row=1,
                                                                                                                 column=2)

        # Dynamic Category Setup
        cat_frame = ttk.LabelFrame(self.tab_classify, text="2. Setup Categories", padding=10)
        cat_frame.pack(padx=10, pady=5, fill="x")

        ttk.Label(cat_frame, text="Categories (Comma Separated):").grid(row=0, column=0, sticky="w", pady=2)
        self.custom_categories_var = tk.StringVar(value="Option A, Option B, Both, Neither, Skip")
        ttk.Entry(cat_frame, textvariable=self.custom_categories_var, width=60).grid(row=1, column=0, padx=5, pady=2)

        ttk.Button(self.tab_classify, text="Load Data & Start Classifying", command=self.load_classifier_data).pack(
            pady=10)

        # Classification Area
        self.work_frame = ttk.Frame(self.tab_classify)

        self.progress_label = ttk.Label(self.work_frame, text="", font=("Arial", 10))
        self.progress_label.pack(pady=5)

        self.company_label = ttk.Label(self.work_frame, text="Company Name", font=("Arial", 14, "bold"), wraplength=600,
                                       justify="center")
        self.company_label.pack(pady=5)

        self.location_label = ttk.Label(self.work_frame, text="Location", font=("Arial", 10, "italic"), wraplength=600,
                                        justify="center")
        self.location_label.pack(pady=5)

        self.category_var = tk.StringVar()
        self.radio_frame = ttk.Frame(self.work_frame)
        self.radio_frame.pack(pady=10)

        btn_frame = ttk.Frame(self.work_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Search Google", command=self.search_web).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save & Next ➔", command=self.save_and_next).pack(side=tk.LEFT, padx=5)

    def build_radio_buttons(self):
        for widget in self.radio_frame.winfo_children():
            widget.destroy()

        categories = [c.strip() for c in self.custom_categories_var.get().split(",") if c.strip()]
        if not categories:
            categories = ["Skip"]

        self.category_var.set(categories[-1])

        for cat in categories:
            ttk.Radiobutton(self.radio_frame, text=cat, variable=self.category_var, value=cat).pack(anchor=tk.W)

    def load_classifier_data(self):
        in_file = self.class_in_var.get()
        out_file = self.class_out_var.get()

        if not os.path.exists(in_file):
            messagebox.showerror("Error", "Input file not found!")
            return

        df_new = pd.read_excel(in_file)
        if "Category" not in df_new.columns:
            df_new["Category"] = ""

        if os.path.exists(out_file):
            df_old = pd.read_excel(out_file)
            self.df = pd.concat([df_old, df_new], ignore_index=True)
            self.df.to_excel(out_file, index=False)
        else:
            self.df = df_new

        self.output_file = out_file
        self.current_index = 0
        for idx, row in self.df.iterrows():
            if pd.isna(row.get('Category')) or str(row.get('Category')).strip() == "":
                self.current_index = idx
                break

        self.build_radio_buttons()
        self.work_frame.pack(fill="both", expand=True)
        self.load_company()

    def load_company(self):
        if self.current_index >= len(self.df):
            messagebox.showinfo("Done", "All companies classified!")
            self.work_frame.pack_forget()
            return

        name_col = next((col for col in self.df.columns if "name" in col.lower() or "company" in col.lower()), None)
        loc_col = next((col for col in self.df.columns if "location" in col.lower() or "address" in col.lower()), None)

        if not name_col:
            messagebox.showerror("Error", "Could not find a 'Company Name' column in the Excel file.")
            return

        company_name = str(self.df.at[self.current_index, name_col]).strip()
        location = "No Location Data"
        if loc_col and pd.notna(self.df.at[self.current_index, loc_col]):
            location = str(self.df.at[self.current_index, loc_col]).replace("\n", ", ")

        self.progress_label.config(text=f"Record {self.current_index + 1} of {len(self.df)}")
        self.company_label.config(text=company_name)
        self.location_label.config(text=location)

        cats = [c.strip() for c in self.custom_categories_var.get().split(",")]
        self.category_var.set(cats[-1] if cats else "Skip")

        self.root.attributes("-topmost", True)
        self.root.after(1000, lambda: self.root.attributes("-topmost", False))

        self.search_web(company_name, location if loc_col else "")

    def search_web(self, name=None, loc=None):
        if not name:
            name_col = next((col for col in self.df.columns if "name" in col.lower() or "company" in col.lower()), None)
            loc_col = next((col for col in self.df.columns if "location" in col.lower() or "address" in col.lower()),
                           None)
            name = str(self.df.at[self.current_index, name_col]).strip()
            loc = str(self.df.at[self.current_index, loc_col]).strip() if loc_col and pd.notna(
                self.df.at[self.current_index, loc_col]) else ""

        loc_clean = re.sub(r'[^\w\s]', '', loc.replace("\n", " ")) if loc else ""
        query = f'"{name}" {loc_clean}' if loc_clean and loc_clean != "No Location Data" else f'"{name}"'
        webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote(query)}")

    def save_and_next(self):
        self.df.at[self.current_index, 'Category'] = self.category_var.get()
        self.df.to_excel(self.output_file, index=False)
        self.current_index += 1
        self.load_company()

    # ==========================================
    #            3. FILTER & EXPORT TAB
    # ==========================================
    def setup_filter_tab(self):
        frame = ttk.LabelFrame(self.tab_filter, text="3. Filter & Export Settings", padding=15)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        # File Selection Row
        f_row = ttk.Frame(frame)
        f_row.pack(fill="x", pady=5)
        ttk.Label(f_row, text="1. Source File:", width=15).pack(side="left")
        ttk.Entry(f_row, textvariable=self.filter_in_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(f_row, text="Browse", command=self.load_columns_for_filter).pack(side="left")

        # Column Criteria Selection Row
        crit_row = ttk.Frame(frame)
        crit_row.pack(fill="x", pady=5)
        ttk.Label(crit_row, text="2. Search Column:", width=15).pack(side="left")
        self.col_dropdown = ttk.Combobox(crit_row, textvariable=self.filter_col_var, state="readonly")
        self.col_dropdown.pack(side="left", fill="x", expand=True, padx=5)

        # Keyword Configuration Row
        kw_row = ttk.Frame(frame)
        kw_row.pack(fill="x", pady=5)
        ttk.Label(kw_row, text="3. Keywords:", width=15).pack(side="left")
        ttk.Entry(kw_row, textvariable=self.filter_kw_var).pack(side="left", fill="x", expand=True, padx=5)

        # Dynamic Columns Checklist with Scrollbars
        ttk.Label(frame, text="4. Columns to Export (Select below):").pack(anchor="w", pady=(10, 0))
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True, pady=5)

        self.export_cols_list = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, exportselection=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.export_cols_list.yview)
        self.export_cols_list.config(yscrollcommand=scrollbar.set)

        self.export_cols_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Selection Helpers
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Select All",
                   command=lambda: self.export_cols_list.select_set(0, tk.END)).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear",
                   command=lambda: self.export_cols_list.selection_clear(0, tk.END)).pack(side="left")

        # Executing Button
        ttk.Button(frame, text="Run Filter & Save New Excel", command=self.run_filter_export).pack(fill="x", pady=10)

    def load_columns_for_filter(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.filter_in_var.set(path)
            try:
                df = pd.read_excel(path)
                cols = list(df.columns)
                self.col_dropdown['values'] = cols
                if cols:
                    self.col_dropdown.current(0)
                self.export_cols_list.delete(0, tk.END)
                for col in cols:
                    self.export_cols_list.insert(tk.END, col)
            except Exception as e:
                messagebox.showerror("Error", f"Could not read columns: {e}")

    def run_filter_export(self):
        if not self.filter_in_var.get() or not os.path.exists(self.filter_in_var.get()):
            messagebox.showerror("Error", "Please select a valid source Excel file.")
            return

        selected_indices = self.export_cols_list.curselection()
        if not selected_indices:
            messagebox.showwarning("Oops", "Please select at least one column to export using the list above.")
            return

        search_col = self.filter_col_var.get()
        if not search_col:
            messagebox.showerror("Error", "Please select a column to search.")
            return

        try:
            df = pd.read_excel(self.filter_in_var.get())
            keywords = [k.strip().lower() for k in self.filter_kw_var.get().split(",") if k.strip()]

            if search_col not in df.columns:
                messagebox.showerror("Error", f"'{search_col}' not found in Excel headers.")
                return

            export_cols = [self.export_cols_list.get(i) for i in selected_indices]

            # Matching function handling NaN cells
            def has_keyword(val):
                if pd.isna(val):
                    return False
                val_str = str(val).lower()
                return any(kw in val_str for kw in keywords)

            mask = df[search_col].apply(has_keyword)
            filtered_df = df[mask][export_cols]

            if filtered_df.empty:
                messagebox.showinfo("Result", "No matches found for those keywords.")
                return

            out_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")],
                                                    initialfile="Filtered_Export.xlsx")
            if out_path:
                filtered_df.to_excel(out_path, index=False)
                messagebox.showinfo("Success",
                                    f"Done! Saved {len(filtered_df)} matching rows to {os.path.basename(out_path)}")
        except Exception as e:
            messagebox.showerror("System Error", f"An error occurred: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalExtractorApp(root)
    root.mainloop()
