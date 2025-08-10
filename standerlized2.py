import pandas as pd
import re
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import sys
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

class MOSFETColumnMapper:
    def __init__(self):
        # Standard column names (target names)
        self.standard_columns = [
            "Manufacturer",
            "Type number", 
            "Package version",
            "Package name",
            "Product status",
            "Configuration",
            "Channel type",
            "Nr of transistors",
            "VDS [max] (V)",
            "RDSon [max] @ VGS = 10 V (mΩ)",
            "RDSon [max] @ VGS = 5 V (mΩ)",
            "RDSon [typ] @ VGS = 10 V (mΩ)",
            "RDSon [typ] @ VGS = 5 V (mΩ)",
            "RDSon [max] @ Tj = 175 °C (mΩ)",
            "Tj [max] (°C)",
            "ID [max] (A)",
            "ID [max] @ T = 100 °C (A)",
            "IDM [max] (A)",
            "QGD [typ] (nC)",
            "QG(tot) [typ] @ VGS = 10 V (nC)",
            "Ptot [max] (W)",
            "Qr [typ] (nC)",
            "VGSth [typ] (V)",
            "Automotive qualified",
            "Ciss [typ] (pF)",
            "Coss [typ] (pF)",
            "Rth(j-mb) [max] (K/W)",
            "Release date",
            "Datasheet",
            "OPN"
        ]
        
        # Alternative search patterns for each standard column
        self.search_patterns = {
            "Manufacturer": [
                "manufacturer", "mfr", "brand", "company", "supplier", "vendor", "make"
            ],
            "Type number": [
                "type", "part number", "part no", "p/n", "pn", "model", "type no", "type_number", 
                "part_number", "partnumber", "typenum", "type_no"
            ],
            "Package version": [
                "package version", "pkg version", "package ver", "pkg ver", "pack version", 
                "package_version", "pkg_version", "pack_ver"
            ],
            "Package name": [
                "package", "pkg", "package name", "package_name", "pkg_name", "pack", 
                "packaging", "case", "outline"
            ],
            "Product status": [
                "status", "product status", "prod status", "product_status", "prod_status", 
                "lifecycle", "active", "obsolete", "nrnd"
            ],
            "Configuration": [
                "config", "configuration", "conf", "setup", "arrangement", "topology"
            ],
            "Channel type": [
                "channel", "channel type", "ch type", "channel_type", "ch_type", "polarity", 
                "n-channel", "p-channel", "nch", "pch"
            ],
            "Nr of transistors": [
                "transistors", "number of transistors", "nr transistors", "no transistors", 
                "transistor count", "nr_transistors", "num_transistors", "transistor_count"
            ],
            "VDS [max] (V)": [
                "vds", "v ds", "vds max", "vds_max", "vdsmax", "drain source voltage", 
                "breakdown voltage", "bvdss", "v_ds", "vds(max)"
            ],
            "RDSon [max] @ VGS = 10 V (mΩ)": [
                "rdson", "rds on", "rds(on)", "rdson 10v", "rdson@10v", "rdson_10v", 
                "rds_on", "on resistance", "on_resistance", "rdson max 10v"
            ],
            "RDSon [max] @ VGS = 5 V (mΩ)": [
                "rdson 5v", "rdson@5v", "rdson_5v", "rds on 5v", "rdson max 5v", 
                "on resistance 5v"
            ],
            "RDSon [typ] @ VGS = 10 V (mΩ)": [
                "rdson typ", "rdson typical", "rdson typ 10v", "rdson_typ_10v", 
                "typical rdson 10v"
            ],
            "RDSon [typ] @ VGS = 5 V (mΩ)": [
                "rdson typ 5v", "rdson typical 5v", "rdson_typ_5v", "typical rdson 5v"
            ],
            "RDSon [max] @ Tj = 175 °C (mΩ)": [
                "rdson 175c", "rdson@175c", "rdson_175c", "rdson high temp", 
                "rdson max 175c", "high temperature rdson"
            ],
            "Tj [max] (°C)": [
                "tj", "junction temp", "junction temperature", "tj max", "tj_max", 
                "max junction temp", "operating temp", "temp max"
            ],
            "ID [max] (A)": [
                "id", "i d", "drain current", "id max", "id_max", "idmax", "continuous drain current", 
                "i_d", "current max"
            ],
            "ID [max] @ T = 100 °C (A)": [
                "id 100c", "id@100c", "id_100c", "drain current 100c", "id max 100c"
            ],
            "IDM [max] (A)": [
                "idm", "i dm", "pulsed drain current", "idm max", "idm_max", "pulse current", 
                "peak drain current"
            ],
            "QGD [typ] (nC)": [
                "qgd", "q gd", "gate drain charge", "qgd typ", "qgd_typ", "miller charge", 
                "gate-drain charge"
            ],
            "QG(tot) [typ] @ VGS = 10 V (nC)": [
                "qg", "q g", "gate charge", "qg tot", "qg total", "qg_tot", "total gate charge", 
                "qg 10v", "gate charge 10v"
            ],
            "Ptot [max] (W)": [
                "ptot", "p tot", "power", "ptot max", "ptot_max", "total power", "max power", 
                "power dissipation", "pd"
            ],
            "Qr [typ] (nC)": [
                "qr", "q r", "reverse recovery charge", "qr typ", "qr_typ", "recovery charge"
            ],
            "VGSth [typ] (V)": [
                "vgsth", "vgs th", "threshold", "gate threshold", "vgsth typ", "vgsth_typ", 
                "threshold voltage", "vth", "v_gsth"
            ],
            "Automotive qualified": [
                "automotive", "auto", "aec", "qualified", "automotive qualified", 
                "automotive_qualified", "aec-q101", "auto qualified"
            ],
            "Ciss [typ] (pF)": [
                "ciss", "c iss", "input capacitance", "ciss typ", "ciss_typ", "gate capacitance"
            ],
            "Coss [typ] (pF)": [
                "coss", "c oss", "output capacitance", "coss typ", "coss_typ", "drain capacitance"
            ],
            "Rth(j-mb) [max] (K/W)": [
                "rth", "thermal resistance", "rth j-mb", "rth(j-mb)", "rth_j_mb", "junction to mounting base", 
                "thermal", "rth max"
            ],
            "Release date": [
                "date", "release date", "release_date", "launch date", "introduction date", 
                "release", "launched"
            ],
            "Datasheet": [
                "datasheet", "data sheet", "spec sheet", "specification", "specs", "documentation", 
                "pdf", "technical data"
            ],
            "OPN": [
                "opn", "order part number", "ordering part number", "order code", "part code", 
                "ordering code", "sales code"
            ]
        }

    def clean_column_name(self, col_name):
        """Clean and normalize column name for better matching"""
        if pd.isna(col_name):
            return ""
        
        # Convert to string and lowercase
        clean_name = str(col_name).lower().strip()
        
        # Remove extra whitespace
        clean_name = re.sub(r'\s+', ' ', clean_name)
        
        # Remove common prefixes/suffixes that might interfere
        clean_name = re.sub(r'^(column|col|field)\s*[:\-_]?\s*', '', clean_name)
        
        return clean_name

    def find_best_match(self, column_name, threshold=70):
        """Find the best matching standard column for a given column name"""
        clean_col = self.clean_column_name(column_name)
        
        if not clean_col:
            return None, 0
        
        best_match = None
        best_score = 0
        
        for standard_col in self.standard_columns:
            # Direct match with standard column
            score = fuzz.ratio(clean_col, standard_col.lower())
            if score > best_score:
                best_score = score
                best_match = standard_col
            
            # Match with alternative patterns
            patterns = self.search_patterns.get(standard_col, [])
            for pattern in patterns:
                score = fuzz.ratio(clean_col, pattern.lower())
                if score > best_score:
                    best_score = score
                    best_match = standard_col
                
                # Also try partial ratio for substring matches
                partial_score = fuzz.partial_ratio(clean_col, pattern.lower())
                if partial_score > best_score and partial_score >= threshold + 10:  # Higher threshold for partial matches
                    best_score = partial_score
                    best_match = standard_col
        
        return best_match if best_score >= threshold else None, best_score

    def map_columns(self, df, threshold=70):
        """Map all columns in the dataframe to standard column names"""
        column_mapping = {}
        unmapped_columns = []
        mapping_scores = {}
        
        for col in df.columns:
            best_match, score = self.find_best_match(col, threshold)
            if best_match:
                column_mapping[col] = best_match
                mapping_scores[col] = score
            else:
                unmapped_columns.append(col)
        
        return column_mapping, unmapped_columns, mapping_scores

    def process_excel_file(self, file_path, sheet_name=None, threshold=70, output_file=None):
        """Process an Excel file and map its columns"""
        try:
            # Read Excel file
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            print(f"Loaded Excel file: {file_path}")
            print(f"Original columns: {len(df.columns)}")
            print("-" * 50)
            
            # Map columns
            column_mapping, unmapped_columns, mapping_scores = self.map_columns(df, threshold)
            
            # Display results
            print("COLUMN MAPPING RESULTS:")
            print("=" * 50)
            
            if column_mapping:
                print(f"\nMAPPED COLUMNS ({len(column_mapping)}):")
                for original, mapped in column_mapping.items():
                    score = mapping_scores[original]
                    print(f"  '{original}' -> '{mapped}' (confidence: {score:.1f}%)")
            
            if unmapped_columns:
                print(f"\nUNMAPPED COLUMNS ({len(unmapped_columns)}):")
                for col in unmapped_columns:
                    print(f"  '{col}'")
            
            # Create mapped dataframe
            mapped_df = df.rename(columns=column_mapping)
            
            # Save to new file if output path provided
            if output_file:
                mapped_df.to_excel(output_file, index=False)
                print(f"\nMapped dataframe saved to: {output_file}")
            
            return mapped_df, column_mapping, unmapped_columns, mapping_scores
            
        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return None, None, None, None

class InteractiveColumnMapper:
    def __init__(self):
        self.mapper = MOSFETColumnMapper()
        self.root = None
        
    def select_input_file(self):
        """Open file dialog to select input Excel file"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return file_path
    
    def get_sheet_names(self, file_path):
        """Get all sheet names from Excel file"""
        try:
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return []
    
    def select_sheet_dialog(self, sheet_names):
        """Show dialog to select sheet"""
        if len(sheet_names) == 1:
            return sheet_names[0]
        
        root = tk.Tk()
        root.title("Select Sheet")
        root.geometry("300x200")
        
        selected_sheet = [None]  # Use list to modify from inner function
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_sheet[0] = sheet_names[selection[0]]
                root.quit()
        
        def on_cancel():
            root.quit()
        
        tk.Label(root, text="Select a sheet to process:", font=("Arial", 10)).pack(pady=10)
        
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set)
        for sheet in sheet_names:
            listbox.insert(tk.END, sheet)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        listbox.selection_set(0)  # Select first item by default
        
        scrollbar.config(command=listbox.yview)
        
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Select", command=on_select).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        root.mainloop()
        root.destroy()
        
        return selected_sheet[0]
    
    def get_threshold_dialog(self):
        """Dialog to get matching threshold"""
        root = tk.Tk()
        root.withdraw()
        
        threshold = simpledialog.askinteger(
            "Matching Threshold",
            "Enter matching threshold (0-100):\nHigher values = stricter matching\nRecommended: 70",
            initialvalue=70,
            minvalue=0,
            maxvalue=100
        )
        root.destroy()
        return threshold if threshold is not None else 70
    
    def select_output_location(self, default_name):
        """Dialog to select output file location and name"""
        root = tk.Tk()
        root.withdraw()
        
        output_path = filedialog.asksaveasfilename(
            title="Save Mapped Excel File As",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return output_path
    
    def show_preview_dialog(self, column_mapping, unmapped_columns, mapping_scores):
        """Show preview of column mappings before processing"""
        root = tk.Tk()
        root.title("Column Mapping Preview")
        root.geometry("800x600")
        
        proceed = [False]  # Use list to modify from inner function
        
        def on_proceed():
            proceed[0] = True
            root.quit()
        
        def on_cancel():
            root.quit()
        
        # Create notebook for tabs
        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Mapped columns tab
        if column_mapping:
            mapped_frame = ttk.Frame(notebook)
            notebook.add(mapped_frame, text=f"Mapped Columns ({len(column_mapping)})")
            
            # Create treeview for mapped columns
            columns = ("Original", "Mapped To", "Confidence")
            mapped_tree = ttk.Treeview(mapped_frame, columns=columns, show="headings", height=15)
            
            for col in columns:
                mapped_tree.heading(col, text=col)
                mapped_tree.column(col, width=250)
            
            for original, mapped in column_mapping.items():
                score = mapping_scores[original]
                mapped_tree.insert("", tk.END, values=(original, mapped, f"{score:.1f}%"))
            
            # Add scrollbar
            mapped_scrollbar = ttk.Scrollbar(mapped_frame, orient=tk.VERTICAL, command=mapped_tree.yview)
            mapped_tree.configure(yscrollcommand=mapped_scrollbar.set)
            
            mapped_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            mapped_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Unmapped columns tab
        if unmapped_columns:
            unmapped_frame = ttk.Frame(notebook)
            notebook.add(unmapped_frame, text=f"Unmapped Columns ({len(unmapped_columns)})")
            
            unmapped_listbox = tk.Listbox(unmapped_frame, font=("Arial", 10))
            for col in unmapped_columns:
                unmapped_listbox.insert(tk.END, col)
            
            unmapped_scroll = tk.Scrollbar(unmapped_frame, orient=tk.VERTICAL, command=unmapped_listbox.yview)
            unmapped_listbox.configure(yscrollcommand=unmapped_scroll.set)
            
            unmapped_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            unmapped_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Button frame
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Proceed with Mapping", command=on_proceed, 
                 bg="green", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        # Summary label
        summary_text = f"Ready to map {len(column_mapping)} columns"
        if unmapped_columns:
            summary_text += f" ({len(unmapped_columns)} will remain unmapped)"
        
        tk.Label(root, text=summary_text, font=("Arial", 10)).pack(pady=5)
        
        root.mainloop()
        root.destroy()
        
        return proceed[0]
    
    def run_interactive(self):
        """Run the interactive version"""
        print("=== Interactive MOSFET Column Mapper ===\n")
        
        # Step 1: Select input file
        print("Step 1: Select input Excel file...")
        input_file = self.select_input_file()
        
        if not input_file:
            print("No file selected. Exiting.")
            return
        
        print(f"Selected: {os.path.basename(input_file)}")
        
        # Step 2: Get sheet names and let user select
        print("\nStep 2: Reading Excel file...")
        sheet_names = self.get_sheet_names(input_file)
        
        if not sheet_names:
            print("Could not read Excel file or no sheets found.")
            return
        
        selected_sheet = self.select_sheet_dialog(sheet_names)
        if not selected_sheet:
            print("No sheet selected. Exiting.")
            return
        
        print(f"Selected sheet: {selected_sheet}")
        
        # Step 3: Get threshold
        print("\nStep 3: Set matching threshold...")
        threshold = self.get_threshold_dialog()
        print(f"Threshold set to: {threshold}%")
        
        # Step 4: Load and analyze the file
        print("\nStep 4: Analyzing columns...")
        try:
            df = pd.read_excel(input_file, sheet_name=selected_sheet)
            column_mapping, unmapped_columns, mapping_scores = self.mapper.map_columns(df, threshold)
            
            print(f"Found {len(df.columns)} columns total")
            print(f"Mapped: {len(column_mapping)} columns")
            print(f"Unmapped: {len(unmapped_columns)} columns")
            
        except Exception as e:
            print(f"Error processing file: {str(e)}")
            messagebox.showerror("Error", f"Error processing file: {str(e)}")
            return
        
        # Step 5: Show preview and get confirmation
        print("\nStep 5: Showing mapping preview...")
        if not self.show_preview_dialog(column_mapping, unmapped_columns, mapping_scores):
            print("Operation cancelled by user.")
            return
        
        # Step 6: Select output location
        print("\nStep 6: Select output location...")
        input_name = Path(input_file).stem
        default_output = f"{input_name}_mapped.xlsx"
        
        output_file = self.select_output_location(default_output)
        if not output_file:
            print("No output location selected. Exiting.")
            return
        
        print(f"Output file: {os.path.basename(output_file)}")
        
        # Step 7: Process and save
        print("\nStep 7: Processing and saving...")
        try:
            mapped_df = df.rename(columns=column_mapping)
            mapped_df.to_excel(output_file, index=False)
            
            print(f"\n✅ SUCCESS!")
            print(f"Mapped file saved as: {output_file}")
            print(f"Columns mapped: {len(column_mapping)}")
            print(f"Threshold used: {threshold}%")
            
            messagebox.showinfo("Success", 
                              f"File processed successfully!\n\n"
                              f"Mapped {len(column_mapping)} columns\n"
                              f"Saved to: {os.path.basename(output_file)}")
            
        except Exception as e:
            print(f"Error saving file: {str(e)}")
            messagebox.showerror("Error", f"Error saving file: {str(e)}")

def main():
    """Main function - choose between command line and interactive mode"""
    if len(sys.argv) > 1:
        # Command line mode (original functionality)
        file_path = sys.argv[1]
        sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
        threshold = int(sys.argv[3]) if len(sys.argv) > 3 else 70
        output_file = sys.argv[4] if len(sys.argv) > 4 else None
        
        mapper = MOSFETColumnMapper()
        result = mapper.process_excel_file(file_path, sheet_name, threshold, output_file)
        
        if result[0] is not None:
            print(f"\nProcessing completed successfully!")
            print(f"Threshold used: {threshold}%")
        else:
            print("Processing failed!")
    else:
        # Interactive mode
        try:
            interactive_mapper = InteractiveColumnMapper()
            interactive_mapper.run_interactive()
        except Exception as e:
            print(f"Error running interactive mode: {str(e)}")
            print("Make sure tkinter is installed: pip install tk")

if __name__ == "__main__":
    main()