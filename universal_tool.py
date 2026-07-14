import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from pytesseract import Output
import pandas as pd
import io
import re
import os
import webbrowser
import urllib.parse
import threading


def get_tesseract_path():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, 'Tesseract-OCR', 'tesseract.exe')
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Tesseract-OCR', 'tesseract.exe')


class UniversalExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal Directory Tool (Extractor, Classifier & Filter)")
        self.root.geometry("750x650")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')

        self.tab_extract = ttk.Frame(self.notebook)
        self.tab_classify = ttk.Frame(self.notebook)
        self.tab_filter = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_extract, text="1. Extract PDF")
        self.notebook.add(self.tab_classify, text="2. Classify Data")
        self.notebook.add(self.tab_filter, text="3. Filter & Export")

        self.filter_in_var = tk.StringVar()
        self.filter_col_var = tk.StringVar()
        self.filter_kw_var = tk.StringVar(value="Solar, Energy, Panel")

        self.df = None
        self.current_index = 0
        self.output_file = ""

        self.setup_extract_tab()
        self.setup_classify_tab()
        self.setup_filter_tab()

    # ==========================================
    #             1. EXTRACTION TAB
    # ==========================================
    def setup_extract_tab(self):
        file_frame = ttk.LabelFrame(self.tab_extract, text="1. File Settings", padding=10)
        file_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(file_frame, text="PDF File:").grid(row=0, column=0, sticky="w", pady=5)
        self.pdf_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.pdf_path_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(file_frame, text="Save Excel As:").grid(row=1, column=0, sticky="w", pady=5)
        self.extract_out_var = tk.StringVar(value="Extracted_Data.xlsx")
        ttk.Entry(file_frame, textvariable=self.extract_out_var, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_extract_out).grid(row=1, column=2)

        page_frame = ttk.Frame(file_frame)
        page_frame.grid(row=2, column=1, sticky="w", pady=10)
        ttk.Label(page_frame, text="Start Page:").pack(side="left")
        self.start_page_var = tk.IntVar(value=1)
        ttk.Entry(page_frame, textvariable=self.start_page_var, width=5).pack(side="left", padx=5)
        ttk.Label(page_frame, text="End Page:").pack(side="left")
        self.end_page_var = tk.IntVar(value=10)
        ttk.Entry(page_frame, textvariable=self.end_page_var, width=5).pack(side="left", padx=5)

        self.tesseract_path_var = tk.StringVar(value=get_tesseract_path())

        auto_frame = ttk.LabelFrame(self.tab_extract, text="2. Smart Auto-Detect Engine", padding=10)
        auto_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(auto_frame,
                  text="Universal Engine will dynamically detect the layout type:\n• Grid Tables (e.g., Financial Sheets, Lists)\n• Borderless Tables (e.g., Tabular Exhibitor Lists)\n• Spatial Profiles (e.g., Vertical cards with Hall/Stall boxes)",
                  justify="left").pack(anchor="w")

        run_frame = ttk.Frame(self.tab_extract)
        run_frame.pack(pady=20)

        self.extract_status_var = tk.StringVar(value="Ready.")
        ttk.Label(run_frame, textvariable=self.extract_status_var, font=("Arial", 10, "bold")).pack(pady=5)

        self.btn_extract = ttk.Button(run_frame, text="START UNIVERSAL EXTRACTION",
                                      command=self.start_extraction_thread)
        self.btn_extract.pack(pady=5)

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path_var.set(path)

    def browse_extract_out(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.extract_out_var.set(path)

    # ==========================================
    #               TEXT HELPERS
    # ==========================================
    def get_words_from_page(self, page):
        words = page.get_text("words")
        if len(words) >= 10:
            return [(w[0], w[1], w[2], w[3], w[4]) for w in words]

        tess_path = self.tesseract_path_var.get()
        if not os.path.exists(tess_path):
            raise FileNotFoundError(f"Tesseract OCR engine not found at:\n{tess_path}")

        pytesseract.pytesseract.tesseract_cmd = tess_path
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT)
        df = pd.DataFrame(ocr_data)
        df = df[(df.conf != '-1') & (df.text.str.strip() != '')]

        ocr_words = []
        for _, row in df.iterrows():
            x0 = float(row['left'])
            y0 = float(row['top'])
            x1 = x0 + float(row['width'])
            y1 = y0 + float(row['height'])
            text = str(row['text'])
            ocr_words.append((x0, y0, x1, y1, text))
        return ocr_words

    def get_layout_preserved_text(self, words_data, gap_threshold=15):
        if not words_data: return ""
        max_x = max(w[2] for w in words_data) if words_data else 0
        current_gap_threshold = gap_threshold * (300 / 72) if max_x > 1200 else gap_threshold

        lines = {}
        for w in words_data:
            x0, y0, x1, y1, text = w
            matched_y = None
            for y in lines:
                if abs(y0 - y) < 8:
                    matched_y = y
                    break
            if matched_y is None:
                lines[y0] = []
                matched_y = y0
            lines[matched_y].append((x0, text, x1))

        col_starts = []
        for y in lines:
            line_words = sorted(lines[y], key=lambda x: x[0])
            if not line_words: continue
            col_starts.append(line_words[0][0])
            prev_x1 = line_words[0][2]
            for w in line_words[1:]:
                x0, text, x1 = w
                if (x0 - prev_x1) > current_gap_threshold:
                    col_starts.append(x0)
                prev_x1 = x1

        if not col_starts: return ""
        col_starts = sorted(col_starts)
        clusters = []
        for x in col_starts:
            if not clusters:
                clusters.append([x])
            else:
                avg_last = sum(clusters[-1]) / len(clusters[-1])
                tolerance = 20 * (300 / 72) if max_x > 1200 else 20
                if x - avg_last < tolerance:
                    clusters[-1].append(x)
                else:
                    clusters.append([x])

        column_coordinates = [min(c) for c in clusters]

        def get_column_index(x0):
            best_idx = 0
            for idx, col_x in enumerate(column_coordinates):
                tol = 10 * (300 / 72) if max_x > 1200 else 10
                if x0 >= col_x - tol:
                    best_idx = idx
                else:
                    break
            return best_idx

        sorted_y = sorted(lines.keys())
        final_text = []
        for y in sorted_y:
            line_words = sorted(lines[y], key=lambda x: x[0])
            column_words = {i: [] for i in range(len(column_coordinates))}
            for word_data in line_words:
                col_idx = get_column_index(word_data[0])
                column_words[col_idx].append(word_data[1])

            line_parts = []
            active_cols = [idx for idx, w in column_words.items() if w]
            if active_cols:
                max_col_idx = max(active_cols)
                for idx in range(max_col_idx + 1):
                    words = column_words[idx]
                    line_parts.append(" ".join(words) if words else "")
                final_text.append("\t".join(line_parts))
            else:
                final_text.append("")

        return "\n".join(final_text)

    # ==========================================
    #          UNIVERSAL PARSING METHODS
    # ==========================================
    def parse_borderless_table(self, text):
        records = []
        current_record = []

        for line in text.split('\n'):
            cols = [c.strip() for c in line.split('\t')]
            if not any(cols): continue

            first_col = cols[0]
            is_new_record = bool(re.match(r'^\d+$', first_col) or re.match(r'^\d+[\.\s]', first_col))
            if not current_record and len(cols) >= 3:
                is_new_record = True

            if is_new_record:
                if current_record: records.append(current_record)
                current_record = cols
            else:
                if current_record:
                    for i in range(len(cols)):
                        val = cols[i]
                        if not val: continue
                        if i < len(current_record):
                            current_record[i] = current_record[i] + " " + val if current_record[i] else val
                        else:
                            current_record.append(val)

        if current_record:
            records.append(current_record)

        dict_records = []
        for r in records:
            dict_records.append({f"Column {i + 1}": col_val for i, col_val in enumerate(r)})
        return dict_records

    def parse_with_spatial_model(self, page):
        """Specifically parses Profile Directory Cards (e.g. Aahar Spices format)"""
        words = page.get_text("words")
        if not words: return []
        words.sort(key=lambda w: w[1])

        lines = []
        current_line_words = []

        if words:
            current_y = words[0][1]
            for w in words:
                x0, y0, x1, y1, text = w[:5]
                if abs(y0 - current_y) < 6:
                    current_line_words.append((x0, text))
                else:
                    current_line_words.sort(key=lambda x: x[0])
                    lines.append(" ".join([item[1] for item in current_line_words]))
                    current_line_words = [(x0, text)]
                    current_y = y0

            if current_line_words:
                current_line_words.sort(key=lambda x: x[0])
                lines.append(" ".join([item[1] for item in current_line_words]))

        records = []
        current_record = {"Company Name": "", "Hall": "", "Stall": "", "Address": "", "Contact Person": "",
                          "Tel./Mobile": "", "E-mail": "", "Website": "", "Products on Display": ""}
        address_buffer = []
        last_key = None

        for line in lines:
            line_str = line.strip()
            if not line_str: continue

            # New card trigger
            is_new_record_trigger = False
            if re.search(r'HALL:\s*([^\s].*?)(?=\s*(?:STALL:|BOOTH:|$))', line_str, re.IGNORECASE) and current_record[
                "Hall"]:
                is_new_record_trigger = True
            elif "Contact Person" in line_str and current_record["Contact Person"]:
                is_new_record_trigger = True

            if is_new_record_trigger and current_record["Company Name"]:
                if address_buffer: current_record["Address"] = " ".join(address_buffer)
                records.append(current_record)
                current_record = {"Company Name": "", "Hall": "", "Stall": "", "Address": "", "Contact Person": "",
                                  "Tel./Mobile": "", "E-mail": "", "Website": "", "Products on Display": ""}
                address_buffer = []
                last_key = None

            temp_line = line_str

            # Hall/Stall Extraction
            h_match = re.search(r'HALL:\s*([^\s].*?)(?=\s*(?:STALL:|BOOTH:|$))', temp_line, re.IGNORECASE)
            if h_match:
                current_record["Hall"] = h_match.group(1).strip()
                temp_line = re.sub(r'HALL:\s*.*?(?=\s*(?:STALL:|BOOTH:|$))', '', temp_line, flags=re.IGNORECASE).strip()

            s_match = re.search(r'(?:STALL|BOOTH):\s*([^\s].*)', temp_line, re.IGNORECASE)
            if s_match:
                current_record["Stall"] = s_match.group(1).strip()
                temp_line = re.sub(r'(?:STALL|BOOTH):\s*.*', '', temp_line, flags=re.IGNORECASE).strip()

            temp_line = temp_line.strip()
            if not temp_line: continue

            # Key-Value extraction
            if ':' in temp_line:
                parts = temp_line.split(':', 1)
                key_raw, val_raw = parts[0].strip(), parts[1].strip()

                if 1 < len(key_raw) < 30 and not key_raw.lower().startswith("http"):
                    if "contact" in key_raw.lower() or "person" in key_raw.lower():
                        current_record["Contact Person"] = val_raw
                        last_key = "Contact Person"
                    elif "tel" in key_raw.lower() or "mobile" in key_raw.lower() or "phone" in key_raw.lower():
                        current_record["Tel./Mobile"] = val_raw
                        last_key = "Tel./Mobile"
                    elif "email" in key_raw.lower() or "e-mail" in key_raw.lower():
                        current_record["E-mail"] = val_raw
                        last_key = "E-mail"
                    elif "website" in key_raw.lower() or "web" in key_raw.lower():
                        current_record["Website"] = val_raw
                        last_key = "Website"
                    elif "product" in key_raw.lower() or "display" in key_raw.lower():
                        current_record["Products on Display"] = val_raw
                        last_key = "Products on Display"
                    else:
                        self.route_text(current_record, address_buffer, last_key, temp_line)
                else:
                    self.route_text(current_record, address_buffer, last_key, temp_line)
            else:
                self.route_text(current_record, address_buffer, last_key, temp_line)

        if current_record["Company Name"]:
            if address_buffer: current_record["Address"] = " ".join(address_buffer)
            records.append(current_record)

        return records

    def route_text(self, record, address_buffer, last_key, text):
        single_line_keys = ["Contact Person", "Tel./Mobile", "E-mail", "Website"]
        if last_key == "Products on Display":
            record["Products on Display"] += " " + text
        elif last_key in single_line_keys:
            address_buffer.append(text)
        else:
            if not record["Company Name"]:
                record["Company Name"] = text
            else:
                address_buffer.append(text)

    def parse_generic_paragraphs(self, text):
        records = []
        blocks = re.split(r'\n\s*\n', text)
        for block in blocks:
            lines = [l.strip() for l in block.split('\n') if l.strip()]
            if len(lines) < 2: continue

            record = {"Company / Title": lines[0], "Email": "", "Phone": "", "Website": "", "Details": ""}
            details = []
            for line in lines[1:]:
                email = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', line)
                if email and not record["Email"]:
                    record["Email"] = email.group(0)
                    line = line.replace(email.group(0), "").strip()

                web = re.search(r'(www\.[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|http[s]?://[^\s]+)', line)
                if web and not record["Website"]:
                    record["Website"] = web.group(0)
                    line = line.replace(web.group(0), "").strip()

                phone = re.search(r'(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{4}', line)
                if phone and not record["Phone"]:
                    record["Phone"] = phone.group(0)
                    line = line.replace(phone.group(0), "").strip()

                if line: details.append(line)
            record["Details"] = "\n".join(details)
            records.append(record)
        return records

    # ==========================================
    #             MAIN EXTRACTION THREAD
    # ==========================================
    def start_extraction_thread(self):
        if not self.pdf_path_var.get():
            messagebox.showerror("Error", "Please select a PDF file.")
            return

        self.btn_extract.config(state="disabled")
        self.extract_status_var.set("Initializing Universal Analysis...")
        threading.Thread(target=self.run_extraction, daemon=True).start()

    def run_extraction(self):
        try:
            doc = fitz.open(self.pdf_path_var.get())
            all_records = []

            for page_num in range(self.start_page_var.get() - 1, self.end_page_var.get()):
                self.extract_status_var.set(f"Analyzing Layout on Page {page_num + 1} of {self.end_page_var.get()}...")
                page = doc.load_page(page_num)
                page_text = page.get_text("text")

                # ==========================================
                # DYNAMIC ROUTING LOGIC
                # ==========================================

                # 1. Is it a Specialized Card Directory? (Like the Spices PDF)
                card_keywords = ["hall:", "stall:", "contact person", "products on display"]
                keyword_matches = sum(1 for kw in card_keywords if kw.lower() in page_text.lower())
                colon_count = page_text.count(":")

                if keyword_matches >= 2 or (colon_count >= 5 and "http" not in page_text.lower()):
                    records = self.parse_with_spatial_model(page)
                    all_records.extend(records)
                    continue

                # 2. Is it a Standard Grid Table? (Like the Bajaj Finance PDF)
                tables = page.find_tables()
                grid_found = False
                if tables and tables.tables:
                    for table in tables.tables:
                        extracted_data = table.extract()
                        if len(extracted_data) > 1:
                            grid_found = True
                            for row in extracted_data:
                                cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                                if any(cleaned_row):
                                    row_dict = {f"Col {i + 1}": col_val for i, col_val in enumerate(cleaned_row)}
                                    all_records.append(row_dict)
                if grid_found:
                    continue

                # 3. Is it a Borderless Table? (High tab density)
                words_data = self.get_words_from_page(page)
                if words_data:
                    layout_text = self.get_layout_preserved_text(words_data)
                    lines = [l for l in layout_text.split('\n') if l.strip()]
                    if lines:
                        tab_density = sum(l.count('\t') for l in lines) / len(lines)
                        if tab_density >= 0.5:
                            records = self.parse_borderless_table(layout_text)
                            all_records.extend(records)
                            continue

                # 4. Fallback: Generic Profiles/Paragraphs
                records = self.parse_generic_paragraphs(page_text)
                all_records.extend(records)

            if not all_records:
                self.extract_status_var.set("No records found.")
                messagebox.showwarning("No Data", "The extraction completed, but 0 records were found on these pages.")
                return

            self.extract_status_var.set("Saving to Excel...")
            # Because different PDFs have different structures, Pandas automatically aligns
            # all dynamic headers (e.g., 'Col 1', 'Company Name', 'Hall') properly!
            df = pd.DataFrame(all_records)
            df.to_excel(self.extract_out_var.get(), index=False)

            self.extract_status_var.set(f"Success! Saved {len(df)} records.")
            messagebox.showinfo("Extraction Complete", f"Saved {len(df)} records to {self.extract_out_var.get()}")

        except Exception as e:
            self.extract_status_var.set("Error occurred.")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_extract.config(state="normal")

    # ==========================================
    #            2. CLASSIFY TAB (UNCHANGED)
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

        cat_frame = ttk.LabelFrame(self.tab_classify, text="2. Setup Categories", padding=10)
        cat_frame.pack(padx=10, pady=5, fill="x")

        ttk.Label(cat_frame, text="Categories (Comma Separated):").grid(row=0, column=0, sticky="w", pady=2)
        self.custom_categories_var = tk.StringVar(value="Option A, Option B, Both, Neither, Skip")
        ttk.Entry(cat_frame, textvariable=self.custom_categories_var, width=60).grid(row=1, column=0, padx=5, pady=2)

        ttk.Button(self.tab_classify, text="Load Data & Start Classifying", command=self.load_classifier_data).pack(
            pady=10)

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
        for widget in self.radio_frame.winfo_children(): widget.destroy()
        categories = [c.strip() for c in self.custom_categories_var.get().split(",") if c.strip()]
        if not categories: categories = ["Skip"]
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
        if "Category" not in df_new.columns: df_new["Category"] = ""

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

        name_col = next((col for col in self.df.columns if
                         "name" in col.lower() or "company" in col.lower() or "col 1" in col.lower()), None)
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
            name_col = next((col for col in self.df.columns if
                             "name" in col.lower() or "company" in col.lower() or "col 1" in col.lower()), None)
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
    #            3. FILTER TAB (UNCHANGED)
    # ==========================================
    def setup_filter_tab(self):
        frame = ttk.LabelFrame(self.tab_filter, text="3. Filter & Export Settings", padding=15)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        f_row = ttk.Frame(frame)
        f_row.pack(fill="x", pady=5)
        ttk.Label(f_row, text="1. Source File:", width=15).pack(side="left")
        ttk.Entry(f_row, textvariable=self.filter_in_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(f_row, text="Browse", command=self.load_columns_for_filter).pack(side="left")

        crit_row = ttk.Frame(frame)
        crit_row.pack(fill="x", pady=5)
        ttk.Label(crit_row, text="2. Search Column:", width=15).pack(side="left")
        self.col_dropdown = ttk.Combobox(crit_row, textvariable=self.filter_col_var, state="readonly")
        self.col_dropdown.pack(side="left", fill="x", expand=True, padx=5)

        kw_row = ttk.Frame(frame)
        kw_row.pack(fill="x", pady=5)
        ttk.Label(kw_row, text="3. Keywords:", width=15).pack(side="left")
        ttk.Entry(kw_row, textvariable=self.filter_kw_var).pack(side="left", fill="x", expand=True, padx=5)

        ttk.Label(frame, text="4. Columns to Export (Select below):").pack(anchor="w", pady=(10, 0))
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True, pady=5)

        self.export_cols_list = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, exportselection=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.export_cols_list.yview)
        self.export_cols_list.config(yscrollcommand=scrollbar.set)
        self.export_cols_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Select All", command=lambda: self.export_cols_list.select_set(0, tk.END)).pack(
            side="left", padx=5)
        ttk.Button(btn_frame, text="Clear", command=lambda: self.export_cols_list.selection_clear(0, tk.END)).pack(
            side="left")
        ttk.Button(frame, text="Run Filter & Save New Excel", command=self.run_filter_export).pack(fill="x", pady=10)

    def load_columns_for_filter(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.filter_in_var.set(path)
            try:
                df = pd.read_excel(path)
                cols = list(df.columns)
                self.col_dropdown['values'] = cols
                if cols: self.col_dropdown.current(0)
                self.export_cols_list.delete(0, tk.END)
                for col in cols: self.export_cols_list.insert(tk.END, col)
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

            def has_keyword(val):
                if pd.isna(val): return False
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
