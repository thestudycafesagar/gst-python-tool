import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import os
import sys

# --- Configuration ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class GSTR3BConverterPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("GSTR-3B Pro Converter")
        self.geometry("700x580")
        self.resizable(False, False)

        # Variables
        self.selected_files = []
        self.merge_mode = ctk.BooleanVar(value=False) # Checkbox variable
        
        # UI Layout
        self.create_widgets()

    def create_widgets(self):
        # Header
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(pady=20)
        
        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="GSTR-3B Batch Processor", 
            font=("Segoe UI", 24)
        )
        self.title_label.pack()
        
        self.subtitle = ctk.CTkLabel(
            self.header_frame,
            text="Saves to: /GST Downloaded/ folder",
            text_color="#10B981"
        )
        self.subtitle.pack()

        # File Selection Area
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(pady=10, padx=40, fill="x")

        self.select_btn = ctk.CTkButton(
            self.file_frame, 
            text="Select PDF Files", 
            command=self.select_files,
            width=200,
            height=40,
            font=("Segoe UI", 14)
        )
        self.select_btn.pack(pady=15)

        self.file_count_label = ctk.CTkLabel(
            self.file_frame,
            text="No files selected",
            text_color="gray"
        )
        self.file_count_label.pack(pady=(0, 10))

        self.btn_demo = ctk.CTkButton(
            self.file_frame, text="▶ View Demo", command=self.open_demo_link,
            fg_color="#DC2626", hover_color="#B91C1C", height=28,
            font=("Segoe UI", 12, "bold"), width=140
        )
        self.btn_demo.pack(pady=(0, 15))

        # --- Options Area (New Checkbox) ---
        self.options_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.options_frame.pack(pady=5)

        self.merge_check = ctk.CTkCheckBox(
            self.options_frame, 
            text="Bunch Converter (Merge all into one Excel)", 
            variable=self.merge_mode,
            font=("Segoe UI", 12),
            checkbox_height=24,
            checkbox_width=24,
            border_width=2
        )
        self.merge_check.pack()
        
        self.hint_label = ctk.CTkLabel(
            self.options_frame,
            text="(Ideal for combining 12 months or 4 quarters into one sheet)",
            font=("Segoe UI", 10),
            text_color="gray"
        )
        self.hint_label.pack()
        self.convert_btn = ctk.CTkButton(
            self,
            text="Convert to Excel Now",
            command=self.process_files,
            state="disabled",
            height=50,
            width=250,
            font=("Segoe UI", 16, "bold"),
            fg_color="#059669",
            hover_color="#047857"
        )
        self.convert_btn.pack(pady=20)

        self.open_folder_btn = ctk.CTkButton(
            self,
            text="📂 Open Output Folder",
            command=self.open_output_folder,
            height=40,
            width=200,
            font=("Segoe UI", 14),
            fg_color="#64748B",
            hover_color="#475569"
        )
        self.open_folder_btn.pack(pady=(0, 10))
        self.open_folder_btn.pack_forget() # Initially hidden

        # Log/Status Area
        self.textbox = ctk.CTkTextbox(self, height=120, width=600)
        self.textbox.pack(pady=10)
        self.textbox.insert("0.0", "Status log will appear here...\n")
        self.textbox.configure(state="disabled")

    def open_demo_link(self):
        import webbrowser
        webbrowser.open_new_tab("https://www.youtube.com/watch?v=XXXXXXXXXX")

    def open_output_folder(self):
        target = os.path.join(os.getcwd(), "GST Downloaded", "GSTR 3B to Excel")
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showinfo("Info", "Output folder not found.")

    def select_files(self):
        filetypes = (("PDF files", "*.pdf"), ("All files", "*.*"))
        filenames = filedialog.askopenfilenames(title="Select GSTR-3B PDFs", filetypes=filetypes)
        
        if filenames:
            self.selected_files = filenames
            count = len(filenames)
            self.file_count_label.configure(text=f"{count} file(s) selected", text_color="white")
            self.convert_btn.configure(state="normal")
            self.log(f"Selected {count} files.")

    def extract_data(self, pdf_path):
        """
        Extracts specific tables and returns them.
        """
        data = {
            "meta": {"GSTIN": "Unknown", "Year": "Unknown", "Month": "Unknown"},
            "supplies": [], 
            "itc": [],      
            "nil": []       
        }

        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text: full_text += text + "\n"
                    
                    tables = page.extract_tables()
                    for table in tables:
                        if not table: continue
                        clean_table = [[str(c).replace('\n', ' ') if c else "0.00" for c in row] for row in table]
                        header = " ".join(clean_table[0]).lower()

                        # 1. Identify Table 3.1 (Supplies)
                        is_supplies = "nature of supplies" in header and "integrated" in header and len(clean_table[0]) >= 5
                        if is_supplies:
                            if any("outward" in str(row).lower() for row in clean_table):
                                for row in clean_table:
                                    if "nature of supplies" in str(row[0]).lower(): continue
                                    if len(row) < 5: continue
                                    data["supplies"].append(row)

                        # 2. Identify Table 4 (ITC)
                        elif "details" in header and "integrated" in header:
                             if any("itc" in str(row).lower() for row in clean_table):
                                for row in clean_table:
                                    if "details" in str(row[0]).lower(): continue 
                                    if "itc available" in str(row[0]).lower(): continue
                                    if len(row) < 5: continue
                                    data["itc"].append(row)

                        # 3. Identify Table 5 (Nil Rated)
                        elif "nature of supplies" in header and "inter-state" in header:
                            for row in clean_table:
                                if "nature of supplies" in str(row[0]).lower(): continue
                                data["nil"].append(row)

                # Metadata Extraction
                gstin_match = re.search(r"GSTIN.*?(\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})", full_text)
                year_match = re.search(r"Year\s*(\d{4}-\d{2})", full_text)
                period_match = re.search(r"Period\s*([A-Za-z]{3}-[A-Za-z]{3}|[A-Za-z]+)", full_text)
                
                if gstin_match: data["meta"]["GSTIN"] = gstin_match.group(1)
                if year_match: data["meta"]["Year"] = year_match.group(1)
                if period_match: data["meta"]["Month"] = period_match.group(1)
                
        except Exception as e:
            self.log(f"Error parsing PDF: {e}")
            
        return data

    def prepare_rows(self, data):
        """
        Flattens the data structure: prepends Meta info to every row.
        Returns: {"supplies": [[meta, row], ...], ...}
        """
        meta = data["meta"]
        clean_float = lambda val: float(str(val).replace(",", "")) if str(val).replace(",", "").replace(".", "").isdigit() else 0.00
        
        prepared = {"supplies": [], "itc": [], "nil": []}

        # Prepare Supplies
        for row in data["supplies"]:
            if len(row) >= 5:
                # Structure: [GSTIN, Year, Month, Nature, Value, IGST, CGST, SGST, Cess]
                nature = row[0]
                val = clean_float(row[1]) if len(row) > 1 else 0
                igst = clean_float(row[2]) if len(row) > 2 else 0
                cgst = clean_float(row[3]) if len(row) > 3 else 0
                sgst = clean_float(row[4]) if len(row) > 4 else 0
                cess = clean_float(row[5]) if len(row) > 5 else 0
                prepared["supplies"].append([meta["GSTIN"], meta["Year"], meta["Month"], nature, val, igst, cgst, sgst, cess])

        # Prepare ITC
        for row in data["itc"]:
            if len(row) >= 4:
                details = row[0]
                igst = clean_float(row[1])
                cgst = clean_float(row[2])
                sgst = clean_float(row[3])
                cess = clean_float(row[4]) if len(row) > 4 else 0
                prepared["itc"].append([meta["GSTIN"], meta["Year"], meta["Month"], details, igst, cgst, sgst, cess])

        # Prepare Nil
        for row in data["nil"]:
            if len(row) >= 3:
                nature = row[0]
                inter = clean_float(row[1])
                intra = clean_float(row[2])
                prepared["nil"].append([meta["GSTIN"], meta["Year"], meta["Month"], nature, inter, intra])

        return prepared

    def write_excel(self, combined_data, output_path):
        wb = openpyxl.Workbook()
        
        # Styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        def setup_sheet(sheet_name, headers, rows):
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.create_sheet(sheet_name)
            
            # Write Headers
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = 20
            
            # Write Rows
            for r_idx, row_data in enumerate(rows, 2):
                for c_idx, val in enumerate(row_data, 1):
                    c = ws.cell(row=r_idx, column=c_idx, value=val)
                    c.border = border

        # 1. Supplies Sheet
        headers_sup = ["GSTIN", "Year", "Month", "Nature of Supplies", "Taxable Value", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"]
        setup_sheet("GSTR3B_Supplies", headers_sup, combined_data["supplies"])

        # 2. ITC Sheet
        headers_itc = ["GSTIN", "Year", "Month", "Details", "Integrated Tax", "Central Tax", "State/UT Tax", "Cess"]
        setup_sheet("GSTR3B_ITC", headers_itc, combined_data["itc"])

        # 3. Nil Sheet
        headers_nil = ["GSTIN", "Year", "Month", "Nature of Supplies", "Inter-State Supplies", "Intra-State Supplies"]
        setup_sheet("GSTR3B_Nil", headers_nil, combined_data["nil"])

        # Remove default sheet
        if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])
        
        wb.save(output_path)

    def process_files(self):
        if not self.selected_files:
            return

        self.convert_btn.configure(state="disabled")
        
        # --- Output Folder Setup ---
        output_folder = os.path.join(os.getcwd(), "GST Downloaded", "GSTR 3B to Excel")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)
        
        self.open_folder_btn.pack_forget()

        # --- LOGIC BRANCHING ---
        is_merge = self.merge_mode.get()
        
        if is_merge:
            # === BUNCH CONVERTER MODE ===
            self.log(f"Starting Bunch Merge for {len(self.selected_files)} files...")
            
            master_data = {"supplies": [], "itc": [], "nil": []}
            first_meta = None

            for idx, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.log(f"Reading ({idx+1}/{len(self.selected_files)}): {filename}...")
                
                # Extract
                raw_data = self.extract_data(file_path)
                if not first_meta: first_meta = raw_data["meta"]
                
                # Prepare (Flatten)
                flat_rows = self.prepare_rows(raw_data)
                
                # Merge
                master_data["supplies"].extend(flat_rows["supplies"])
                master_data["itc"].extend(flat_rows["itc"])
                master_data["nil"].extend(flat_rows["nil"])

            # Save Consolidated File
            out_name = "GSTR3B_Consolidated.xlsx"
            if first_meta and first_meta["Year"] != "Unknown":
                out_name = f"GSTR3B_Consolidated_{first_meta['Year']}.xlsx"
            
            out_path = os.path.join(output_folder, out_name)
            self.write_excel(master_data, out_path)
            
            self.log("----------------")
            self.log(f"MERGE SUCCESSFUL!")
            self.log(f"Saved: {out_name}")
            messagebox.showinfo("Success", f"Merged {len(self.selected_files)} files into:\n{out_name}")
            self.open_folder_btn.pack(pady=(0, 10))

        else:
            # === INDIVIDUAL MODE ===
            self.log(f"Starting Individual Conversion...")
            
            for idx, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.log(f"Processing ({idx+1}/{len(self.selected_files)}): {filename}")
                
                # Extract
                raw_data = self.extract_data(file_path)
                
                # Prepare
                flat_rows = self.prepare_rows(raw_data)
                
                # Naming
                safe_month = raw_data["meta"]["Month"].replace(" ", "_")
                safe_year = raw_data["meta"]["Year"].replace("-", "_")
                if safe_month == "Unknown":
                    out_name = filename.replace(".pdf", ".xlsx")
                else:
                    out_name = f"GSTR3B_{safe_month}_{safe_year}.xlsx"
                
                out_path = os.path.join(output_folder, out_name)
                
                # Write
                self.write_excel(flat_rows, out_path)
                self.log(f" -> Saved: {out_name}")

            self.log("----------------")
            self.log("Batch Complete.")
            self.open_folder_btn.pack(pady=(0, 10))
            messagebox.showinfo("Done", f"Processed {len(self.selected_files)} files.")

        self.convert_btn.configure(state="normal")
        self.selected_files = []
        self.file_count_label.configure(text="No files selected", text_color="gray")

if __name__ == "__main__":
    app = GSTR3BConverterPro()
    app.mainloop()