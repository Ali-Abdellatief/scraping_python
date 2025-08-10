import pandas as pd
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import scrolledtext
import glob

class ExcelCombiner:
    def __init__(self):
        self.selected_files = []
        self.combined_df = None
        self.source_column_name = "Source_File"
        
    def select_multiple_files(self):
        """Open file dialog to select multiple Excel files"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files to Combine",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return list(file_paths) if file_paths else []
    
    def select_folder_files(self):
        """Select all Excel files from a folder"""
        root = tk.Tk()
        root.withdraw()
        
        folder_path = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        root.destroy()
        
        if not folder_path:
            return []
        
        # Find all Excel files in the folder
        excel_files = []
        for pattern in ['*.xlsx', '*.xls']:
            excel_files.extend(glob.glob(os.path.join(folder_path, pattern)))
        
        return excel_files
    
    def get_file_selection_method(self):
        """Dialog to choose file selection method"""
        root = tk.Tk()
        root.title("File Selection Method")
        root.geometry("400x200")
        root.resizable(False, False)
        
        selection_method = [None]  # Use list to modify from inner function
        
        def select_individual():
            selection_method[0] = "individual"
            root.quit()
        
        def select_folder():
            selection_method[0] = "folder"
            root.quit()
        
        def cancel():
            root.quit()
        
        # Title
        tk.Label(root, text="How would you like to select Excel files?", 
                font=("Arial", 12, "bold")).pack(pady=20)
        
        # Buttons frame
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Select Individual Files", 
                 command=select_individual, width=20, height=2,
                 bg="#4CAF50", fg="white", font=("Arial", 10)).pack(pady=5)
        
        tk.Button(button_frame, text="Select All Files from Folder", 
                 command=select_folder, width=20, height=2,
                 bg="#2196F3", fg="white", font=("Arial", 10)).pack(pady=5)
        
        tk.Button(button_frame, text="Cancel", 
                 command=cancel, width=20, height=1).pack(pady=10)
        
        root.mainloop()
        root.destroy()
        
        return selection_method[0]
    
    def preview_files_dialog(self, file_paths):
        """Show preview of selected files and their basic info"""
        if not file_paths:
            return False
            
        root = tk.Tk()
        root.title("File Preview")
        root.geometry("800x500")
        
        proceed = [False]
        
        def on_proceed():
            proceed[0] = True
            root.quit()
        
        def on_cancel():
            root.quit()
        
        def remove_file():
            selection = file_tree.selection()
            if selection:
                item = selection[0]
                file_path = file_tree.item(item)['values'][0]  # Full path stored in first column
                if file_path in file_paths:
                    file_paths.remove(file_path)
                file_tree.delete(item)
                update_summary()
        
        def update_summary():
            summary_label.config(text=f"Total files selected: {len(file_paths)}")
        
        # Main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        tk.Label(main_frame, text="Selected Files Preview", 
                font=("Arial", 14, "bold")).pack(pady=(0, 10))
        
        # File tree
        columns = ("File Name", "Sheet Count", "Rows", "Columns", "Size")
        file_tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=12)
        
        for col in columns:
            file_tree.heading(col, text=col)
        
        # Set column widths
        file_tree.column("File Name", width=200)
        file_tree.column("Sheet Count", width=80)
        file_tree.column("Rows", width=80)
        file_tree.column("Columns", width=80)
        file_tree.column("Size", width=80)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=file_tree.yview)
        file_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate tree with file info
        for file_path in file_paths:
            try:
                # Get basic file info
                file_name = os.path.basename(file_path)
                file_size = f"{os.path.getsize(file_path) / 1024:.1f} KB"
                
                # Try to get Excel info
                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_count = len(excel_file.sheet_names)
                    
                    # Read first sheet to get dimensions
                    df_sample = pd.read_excel(file_path, nrows=0)  # Just get column info
                    df_full = pd.read_excel(file_path)  # Get row count
                    
                    rows = len(df_full)
                    columns = len(df_sample.columns)
                    
                except Exception:
                    sheet_count = "Error"
                    rows = "Error"
                    columns = "Error"
                
                # Store full path in values for removal functionality
                file_tree.insert("", tk.END, values=(file_path, sheet_count, rows, columns, file_size))
                
            except Exception as e:
                file_tree.insert("", tk.END, values=(file_path, "Error", "Error", "Error", "Error"))
        
        # Summary label
        summary_label = tk.Label(main_frame, text=f"Total files selected: {len(file_paths)}", 
                                font=("Arial", 10))
        summary_label.pack(pady=(10, 0))
        
        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Remove Selected File", 
                 command=remove_file, bg="#f44336", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Proceed with Combination", 
                 command=on_proceed, bg="#4CAF50", fg="white", 
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        root.mainloop()
        root.destroy()
        
        return proceed[0] and len(file_paths) > 0
    
    def get_combination_options(self):
        """Dialog to get combination options"""
        root = tk.Tk()
        root.title("Combination Options")
        root.geometry("500x500")  # Increased height for new option
        
        options = {}
        proceed = [False]
        
        def on_proceed():
            # Get all option values
            options['add_source_column'] = add_source_var.get()
            options['source_column_name'] = source_name_entry.get().strip() or "Source_File"
            options['remove_duplicates'] = remove_duplicates_var.get()
            options['combine_method'] = combine_method_var.get()
            options['handle_missing'] = missing_method_var.get()
            options['add_custom_columns'] = add_custom_columns_var.get()  # NEW OPTION
            proceed[0] = True
            root.quit()
        
        def on_cancel():
            root.quit()
        
        # Title
        tk.Label(root, text="Excel Combination Options", 
                font=("Arial", 14, "bold")).pack(pady=10)
        
        # Main options frame
        options_frame = tk.Frame(root)
        options_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Source column option
        source_frame = tk.LabelFrame(options_frame, text="Source Tracking", padx=10, pady=10)
        source_frame.pack(fill=tk.X, pady=5)
        
        add_source_var = tk.BooleanVar(value=True)
        tk.Checkbutton(source_frame, text="Add source file column", 
                      variable=add_source_var).pack(anchor=tk.W)
        
        tk.Label(source_frame, text="Column name:").pack(anchor=tk.W, pady=(5,0))
        source_name_entry = tk.Entry(source_frame, width=30)
        source_name_entry.insert(0, "Source_File")
        source_name_entry.pack(anchor=tk.W, pady=(0,5))
        
        # NEW: Custom columns option
        custom_columns_frame = tk.LabelFrame(options_frame, text="Custom Columns", padx=10, pady=10)
        custom_columns_frame.pack(fill=tk.X, pady=5)
        
        add_custom_columns_var = tk.BooleanVar(value=False)
        custom_checkbox = tk.Checkbutton(custom_columns_frame, 
                                       text="Add BG, MAG, and Category columns (empty)", 
                                       variable=add_custom_columns_var)
        custom_checkbox.pack(anchor=tk.W)
        
        # Info label for custom columns
        info_label = tk.Label(custom_columns_frame, 
                             text="This will add three empty columns: 'BG', 'MAG', and 'Category'",
                             font=("Arial", 8), fg="gray")
        info_label.pack(anchor=tk.W, pady=(2,0))
        
        # Duplicate handling
        duplicate_frame = tk.LabelFrame(options_frame, text="Duplicate Handling", padx=10, pady=10)
        duplicate_frame.pack(fill=tk.X, pady=5)
        
        remove_duplicates_var = tk.BooleanVar(value=False)
        tk.Checkbutton(duplicate_frame, text="Remove duplicate rows", 
                      variable=remove_duplicates_var).pack(anchor=tk.W)
        
        # Combination method
        method_frame = tk.LabelFrame(options_frame, text="Combination Method", padx=10, pady=10)
        method_frame.pack(fill=tk.X, pady=5)
        
        combine_method_var = tk.StringVar(value="append")
        tk.Radiobutton(method_frame, text="Append all data (recommended)", 
                      variable=combine_method_var, value="append").pack(anchor=tk.W)
        tk.Radiobutton(method_frame, text="Only keep common columns", 
                      variable=combine_method_var, value="inner").pack(anchor=tk.W)
        
        # Missing data handling
        missing_frame = tk.LabelFrame(options_frame, text="Missing Data Handling", padx=10, pady=10)
        missing_frame.pack(fill=tk.X, pady=5)
        
        missing_method_var = tk.StringVar(value="keep")
        tk.Radiobutton(missing_frame, text="Keep missing values as NaN", 
                      variable=missing_method_var, value="keep").pack(anchor=tk.W)
        tk.Radiobutton(missing_frame, text="Fill missing values with empty string", 
                      variable=missing_method_var, value="fill").pack(anchor=tk.W)
        
        # Buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="Start Combination", command=on_proceed,
                 bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        root.mainloop()
        root.destroy()
        
        return options if proceed[0] else None
    
    def combine_files(self, file_paths, options):
        """Combine the selected Excel files"""
        combined_data = []
        processing_log = []
        
        for i, file_path in enumerate(file_paths):
            try:
                processing_log.append(f"Processing: {os.path.basename(file_path)}")
                
                # Read the Excel file
                df = pd.read_excel(file_path)
                
                # Add source column if requested
                if options['add_source_column']:
                    df[options['source_column_name']] = os.path.basename(file_path)
                
                # Add custom columns if requested
                if options['add_custom_columns']:
                    df['BG'] = ''  # Empty string
                    df['MAG'] = ''  # Empty string
                    df['Category'] = ''  # Empty string
                    processing_log.append(f"  ✓ Added BG, MAG, and Category columns")
                
                combined_data.append(df)
                processing_log.append(f"  ✓ Added {len(df)} rows, {len(df.columns)} columns")
                
            except Exception as e:
                error_msg = f"  ✗ Error: {str(e)}"
                processing_log.append(error_msg)
                print(error_msg)
        
        if not combined_data:
            return None, processing_log
        
        # Combine all dataframes
        processing_log.append("\nCombining data...")
        
        if options['combine_method'] == 'inner':
            # Only keep common columns
            self.combined_df = pd.concat(combined_data, join='inner', ignore_index=True)
        else:
            # Append all data (outer join)
            self.combined_df = pd.concat(combined_data, join='outer', ignore_index=True)
        
        # Handle missing data
        if options['handle_missing'] == 'fill':
            self.combined_df = self.combined_df.fillna('')
        
        # Remove duplicates if requested
        if options['remove_duplicates']:
            initial_rows = len(self.combined_df)
            self.combined_df = self.combined_df.drop_duplicates()
            removed_duplicates = initial_rows - len(self.combined_df)
            processing_log.append(f"Removed {removed_duplicates} duplicate rows")
        
        # Log information about custom columns
        if options['add_custom_columns']:
            processing_log.append("Added custom columns: BG, MAG, Category")
        
        processing_log.append(f"\n✓ Final result: {len(self.combined_df)} rows, {len(self.combined_df.columns)} columns")
        
        return self.combined_df, processing_log
    
    def show_processing_dialog(self, processing_log):
        """Show processing results and log"""
        root = tk.Tk()
        root.title("Processing Results")
        root.geometry("600x400")
        
        proceed = [False]
        
        def on_proceed():
            proceed[0] = True
            root.quit()
        
        def on_cancel():
            root.quit()
        
        # Title
        tk.Label(root, text="File Processing Complete!", 
                font=("Arial", 14, "bold")).pack(pady=10)
        
        # Log display
        log_frame = tk.Frame(root)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(log_frame, text="Processing Log:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        
        log_text = scrolledtext.ScrolledText(log_frame, height=15, width=70)
        log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Add log content
        for line in processing_log:
            log_text.insert(tk.END, line + "\n")
        
        log_text.config(state=tk.DISABLED)  # Make read-only
        
        # Buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=15)
        
        tk.Button(button_frame, text="Save Combined File", command=on_proceed,
                 bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        root.mainloop()
        root.destroy()
        
        return proceed[0]
    
    def select_output_location(self, default_name="combined_excel_files.xlsx"):
        """Dialog to select output file location and name"""
        root = tk.Tk()
        root.withdraw()
        
        output_path = filedialog.asksaveasfilename(
            title="Save Combined Excel File As",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return output_path
    
    def run_interactive(self):
        """Run the interactive Excel combiner"""
        print("=== Interactive Excel File Combiner ===\n")
        
        # Step 1: Choose file selection method
        print("Step 1: Choose file selection method...")
        selection_method = self.get_file_selection_method()
        
        if not selection_method:
            print("No selection method chosen. Exiting.")
            return
        
        # Step 2: Select files
        print(f"Step 2: Selecting files ({selection_method} method)...")
        
        if selection_method == "individual":
            file_paths = self.select_multiple_files()
        else:  # folder
            file_paths = self.select_folder_files()
        
        if not file_paths:
            print("No files selected. Exiting.")
            return
        
        print(f"Selected {len(file_paths)} files")
        
        # Step 3: Preview files
        print("Step 3: Showing file preview...")
        if not self.preview_files_dialog(file_paths):
            print("Operation cancelled or no files to process.")
            return
        
        # Step 4: Get combination options
        print("Step 4: Getting combination options...")
        options = self.get_combination_options()
        
        if not options:
            print("Operation cancelled.")
            return
        
        # Step 5: Combine files
        print("Step 5: Combining files...")
        combined_df, processing_log = self.combine_files(file_paths, options)
        
        if combined_df is None:
            messagebox.showerror("Error", "Failed to combine any files. Check the processing log.")
            return
        
        # Step 6: Show processing results
        print("Step 6: Showing processing results...")
        if not self.show_processing_dialog(processing_log):
            print("Operation cancelled.")
            return
        
        # Step 7: Save combined file
        print("Step 7: Saving combined file...")
        output_path = self.select_output_location()
        
        if not output_path:
            print("No output location selected. Exiting.")
            return
        
        try:
            combined_df.to_excel(output_path, index=False)
            
            print(f"\n✅ SUCCESS!")
            print(f"Combined file saved as: {output_path}")
            print(f"Total rows: {len(combined_df)}")
            print(f"Total columns: {len(combined_df.columns)}")
            print(f"Files combined: {len(file_paths)}")
            
            messagebox.showinfo("Success!", 
                              f"Files combined successfully!\n\n"
                              f"Combined {len(file_paths)} files\n"
                              f"Total rows: {len(combined_df):,}\n"
                              f"Total columns: {len(combined_df.columns)}\n"
                              f"Saved to: {os.path.basename(output_path)}")
            
        except Exception as e:
            error_msg = f"Error saving combined file: {str(e)}"
            print(error_msg)
            messagebox.showerror("Error", error_msg)

def main():
    """Main function"""
    try:
        combiner = ExcelCombiner()
        combiner.run_interactive()
    except Exception as e:
        print(f"Error running Excel combiner: {str(e)}")
        print("Make sure all required packages are installed:")
        print("pip install pandas openpyxl tkinter")

if __name__ == "__main__":
    main()