import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from fuzzywuzzy import process

STANDARD_PARAMS = {
    # --- Generic / commercial ---
    "Part Number": ["part no", "pn", "device", "type number", "partnumber"],
    "Description": ["desc", "product description"],
    "Lifecycle Status": ["status", "life cycle", "marketing status"],
    "Unit Price (1ku)": ["price", "unit price", "price|quantity (usd)"],
    "Package Type": ["package", "pkg", "package name", "package type", "package type2"],
    "Manufacturer": ["mfr", "vendor", "maker"],
    "Product Link": ["product url", "product page", "link", "web link"],
    "Datasheet Link": ["datasheet", "pdf", "pdf data sheet", "datasheet link"],
    "ECAD Model": ["ecad", "ecad model"],
    "Certifications": ["cert", "compliance", "rohs"],
    "Automotive Grade": ["automotive", "aec-q", "automotive qualified"],
    "Operating Temp Range (°C)": [
        "temperature range", "operating temperature", "temp range",
        "operating temperature range (°c)", "operating temp range (Â°c)"
    ],
    "Mounting Type": ["mount type", "mounting"],
    "Type": ["type", "category", "product family"],
    "Eval Board Available": ["eval board", "evaluation board"],
    "Reliability Rating": ["reliability"],
    "OPN": ["opn", "orderable part number"],
    "Coupon Available": ["coupon", "discount"],

    # --- Supply & basic electrical (generic) ---
    "Number of Channels": ["channels", "channel count", "number of channels nom"],
    "VIN Min (V)": [
        "vin min", "input voltage min", "supply voltage (v) min",
        "vin min v (volt)", "vin min (v)2", "vin min (v)3", "vin min (v)4", "vin min (v)5"
    ],
    "VIN Max (V)": [
        "vin max", "input voltage max", "supply voltage (v) max", "vin max v (volt)"
    ],
    "Nominal Current (A)": [
        "imax", "nominal current", "output current", "recommended output current (a) max",
        "nominal current (a)3"
    ],
    "Current Limit Min (A)": ["current limit", "current limit min (a)"],
    "Current Limit Max (A)": ["current limit max (a)"],
    "Iq (µA)": [
        "quiescent current", "iq", "supply current (µa) (on) max",
        "iq (µa)4", "iq (µa)3", "iq (µa)6"
    ],
    "Shutdown Current (µA)": ["shutdown current", "isd", "supply current (µa) (off) max"],
    "Overvoltage Protection (V)": ["ovp", "over voltage protection", "vin ovp"],
    "Short Circuit Protection": ["scp", "short circuit", "short-circuit protection"],
    "Overtemperature Protection": ["otp", "over temp", "thermal protection", "overtemperature"],
    "Enable Pin": ["enable", "enable control", "en", "enable pin active level"],
    "Fault Indicator": ["fault flag", "fault", "status pin"],
    "Soft Start / Inrush Control": ["soft start", "inrush", "inrush current limiting"],

    # --- Switch/load specific (kept from your set) ---
    "RDS(on) (25°C) (mΩ)": [
        "ron (25c)", "rds on", "on resistance", "ron", "rds(on) @5v (mo)",
        "rds(on) (25°c) (mo)", "ron (mω)"
    ],
    "RDS(on) (150°C) (mΩ)": ["ron (150c)", "rds(on) (150°c) (mo)"],
    "Integrated FET": ["integrated fet", "fet included", "mosfet included"],
    "Back-to-Back FETs": ["b2b fets", "reverse blocking", "back-to-back fets"],

    # --- I²C / I³C Specific ---
    "Interface Type": ["interface", "protocol", "bus type", "i2c/i3c interface"],
    "Max Bus Speed (MHz)": [
        "max bus speed", "max data rate", "max frequency", "clock speed",
        "i2c speed", "i3c speed"
    ],
    "Addressing": [
        "addressing", "i2c address", "slave address", "device address",
        "address bits", "7-bit address", "10-bit address", "addresses"
    ],
    "I3C HDR Support": ["hdr support", "i3c hdr modes", "high data rate support"],
    "Hot Join Support": ["hot join", "hot-join support", "hj"],
    "In-Band Interrupt": ["in-band interrupt", "ibi", "ibi support"],
    "Number of Ports": ["number of ports", "ports", "channels", "bus count"],
    "Master/Slave Support": ["master/slave", "master mode", "slave mode", "multi-master support"],
    "Bus Voltage Levels": ["logic level", "bus voltage", "v ih", "v il"],
    "Pull-up Resistor": ["pull-up", "internal pull-up", "external pull-up"],

    # === Isolated Gate Driver Specific ===
    # Supply & output drive
    "Driver Supply Voltage Min (V)": [
        "vdd min", "vcc min", "driver vdd(min)", "driver supply min", "logic supply min"
    ],
    "Driver Supply Voltage Max (V)": [
        "vdd max", "vcc max", "driver vdd(max)", "driver supply max", "logic supply max"
    ],
    "Output Peak Source Current (A)": [
        "source current", "peak source", "io_source", "output source current", "gate source current",
        "peak source current (a)", "source (a)", "source (ma)"
    ],
    "Output Peak Sink Current (A)": [
        "sink current", "peak sink", "io_sink", "output sink current", "gate sink current",
        "peak sink current (a)", "sink (a)", "sink (ma)"
    ],
    "Output Peak Current (A)": [
        "peak output current", "peak gate current", "io (peak)", "drive current", "source/sink current"
    ],
    "Output Voltage Swing (V)": [
        "output high voltage", "voh", "gate drive voltage", "vgate", "vo range", "output voltage range"
    ],
    "UVLO On (V)": [
        "uvlo+", "uvlo rising", "uvlo on threshold", "uvlo on", "uvlo (rising)"
    ],
    "UVLO Off (V)": [
        "uvlo-", "uvlo falling", "uvlo off threshold", "uvlo off", "uvlo (falling)"
    ],

    # Timing & speed
    "Propagation Delay (ns)": [
        "tpd", "prop delay", "input-to-output delay", "delay (ns)"
    ],
    "Propagation Delay Matching (ns)": [
        "channel-to-channel delay", "delay matching", "skew", "tpd match"
    ],
    "Rise Time (ns)": ["tr", "rise time"],
    "Fall Time (ns)": ["tf", "fall time"],
    "Max Switching Frequency (kHz)": [
        "max pwm frequency", "switching freq max", "f_sw", "max frequency (khz)"
    ],
    "Dead-Time Control": ["dead time", "dead-time", "dt control", "shoot-through protection"],

    # Isolation & robustness
    "Isolation Voltage (Vrms)": [
        "isolation rating", "v_iso", "withstand voltage", "isolation voltage", "viorm (rms)"
    ],
    "Working Isolation Voltage (V)": [
        "viorm", "working voltage", "continuous working voltage"
    ],
    "CMTI (kV/µs) Min": [
        "cmti", "common mode transient immunity", "dv/dt immunity", "dvdt immunity"
    ],
    "Isolation Type": [
        "isolation technology", "capacitive isolation", "magnetic isolation", "optocoupler", "optical isolation"
    ],
    "Isolation Certifications": [
        "ul1577", "vde 0884-11", "iec 60747-17", "certified isolation", "reinforced/basic isolation"
    ],
    "Creepage (mm)": ["creepage distance", "creepage"],
    "Clearance (mm)": ["clearance distance", "clearance"],

    # Protection features
    "Desaturation Detection": ["desat", "desaturation", "short-circuit detect", "desat protection"],
    "Miller Clamp": ["miller clamp", "miller", "gate clamp"],
    "Soft Turn-Off": ["soft turn off", "soft-off", "fault soft-off"],
    "Active Clamping": ["active clamp", "gate active clamp", "dv/dt clamp"],
    "Fault Reporting": ["fault output", "fault pin", "fault flag", "nflt", "flt"],
    "Fault Reset": ["fault reset", "reset pin", "rstrt", "auto reset"],

    # Function / topology
    "Topology": ["half-bridge", "high-side/low-side", "single driver", "dual driver", "hb"],
    "Target Switch Type": ["igbt", "mosfet", "sic", "gan", "driver type", "switch type"],
    "Input Logic Compatibility": ["cmos", "ttl", "input thresholds", "vih/vil logic level"],
    "Enable/Disable Control": ["shutdown", "ena pin", "dis pin", "sd pin"],

    # Mechanical/env (extras useful for drivers)
    "Package Size (mm)": ["size_mm", "package size (l x w)", "body size", "package area"],
    "Operating Temp Min (°C)": ["ta min", "ambient min", "operating temp min"],
    "Operating Temp Max (°C)": ["ta max", "ambient max", "operating temp max"],
    "AEC-Q100 Grade": ["aec grade", "automotive grade code", "q100 grade"],

    # === NEW: Isolation Amplifier Specific ===
    # Core electrical performance
    "Common Mode Rejection Ratio (CMRR)": [
        "cmrr", "common mode rejection", "cmr", "common mode suppression", "cmrr (db)"
    ],
    "Input Type": [
        "input type", "signal type", "voltage input", "current input", "differential input",
        "single-ended", "differential", "ac/dc coupling"
    ],
    "Gain (V/V)": [
        "gain", "voltage gain", "amplifier gain", "fixed gain", "programmable gain",
        "gain (v/v)", "gain accuracy", "gain (db)"
    ],
    "Gain Error (%)": [
        "gain error", "gain accuracy", "gain tolerance", "gain deviation", "gain error (%)"
    ],
    "Bandwidth (Hz)": [
        "bandwidth", "frequency response", "3db bandwidth", "signal bandwidth",
        "bandwidth (hz)", "bandwidth (khz)", "bandwidth (mhz)", "-3db frequency"
    ],
    "Signal-to-Noise Ratio (SNR)": [
        "snr", "signal to noise ratio", "s/n ratio", "snr (db)", "noise performance"
    ],
    "Offset Voltage (mV)": [
        "offset voltage", "input offset", "vos", "dc offset", "output offset",
        "offset drift", "vos (mv)", "offset voltage (µv)"
    ],
    "Input Bias Current (nA)": [
        "input bias current", "bias current", "ib", "input current", "ib (na)",
        "input bias current (pa)", "input bias current (µa)"
    ],

    # Power specifications  
    "Supply Voltage (V)": [
        "supply voltage", "vcc", "vdd", "vs", "supply range", "dual supply",
        "single supply", "±vs", "supply voltage (v)"
    ],
    "Supply Current (mA)": [
        "supply current", "icc", "idd", "is", "operating current", "supply current (ma)",
        "supply current (µa)", "quiescent current (ma)"
    ],
    "Power Consumption (mW)": [
        "power consumption", "power dissipation", "total power", "power (mw)",
        "power consumption (w)", "operating power"
    ],

    # Isolation specifications (additional to existing)
    "Reinforced/Basic Isolation": [
        "reinforced isolation", "basic isolation", "isolation class", "safety rating",
        "insulation type", "isolation standard"
    ],

    # === NEW: Transformer Driver Specific ===
    # Electrical input/output ratings
    "Supply Voltage Range (V)": [
        "vdd min/max", "supply range", "vcc range", "operating voltage range",
        "supply voltage min/max", "vdd operating range"
    ],
    "Input Signal Type": [
        "input signal type", "logic level input", "differential input signal", "pwm input",
        "signal input type", "input signal format", "logic input", "cmos input", "ttl input"
    ],
    "Output Drive Voltage (V)": [
        "output drive voltage", "drive voltage", "transformer drive voltage", "vo drive",
        "output voltage swing", "drive level", "excitation voltage"
    ],
    "Output Drive Current (A)": [
        "output drive current", "drive current", "transformer drive current", "io drive",
        "excitation current", "drive capability (a)"
    ],
    "Load Drive Capability": [
        "load drive capability", "maximum load", "load capacity", "drive strength",
        "load drive current", "transformer load capability", "max load impedance"
    ],

    # Frequency performance
    "Operating Frequency Range (Hz)": [
        "operating frequency range", "frequency range", "min/max frequency", "operating freq",
        "frequency capability", "switching frequency range", "freq range (hz)", "freq range (khz)"
    ],
    "Duty Cycle Tolerance (%)": [
        "duty cycle tolerance", "duty cycle range", "pulse width tolerance", "duty cycle capability",
        "min/max duty cycle", "duty cycle limits"
    ],

    # Efficiency & thermal
    "Efficiency (%)": [
        "efficiency", "power efficiency", "conversion efficiency", "efficiency (%)",
        "power transfer efficiency", "driver efficiency"
    ],
    "Thermal Resistance (°C/W)": [
        "thermal resistance", "θja", "junction to ambient", "thermal resistance (c/w)",
        "rθja", "thermal impedance", "junction-to-ambient thermal resistance"
    ],

    # Protection & robustness
    "ESD Protection (kV)": [
        "esd protection", "electrostatic discharge", "esd rating", "esd withstand",
        "hbm esd", "cdm esd", "esd tolerance", "static discharge protection"
    ],

    # === NEW: Digital Isolator Specific ===
    # Channel configuration
    "Channel Direction": [
        "channel direction", "signal direction", "unidirectional", "bidirectional", 
        "mixed direction", "forward/reverse", "input/output direction"
    ],
    "Logic Signal Type": [
        "logic signal type", "logic family", "cmos logic", "ttl logic", "lvds",
        "signal compatibility", "input logic type", "output logic type"
    ],

    # Performance specifications
    "Data Rate (Mbps)": [
        "data rate", "max data rate", "signal speed", "transfer rate", "data rate (mbps)",
        "max frequency (mhz)", "switching speed", "signal bandwidth (mhz)"
    ],
    "Pulse Width Distortion (PWD)": [
        "pulse width distortion", "pwd", "pulse distortion", "timing distortion",
        "pulse width error", "duty cycle distortion"
    ],
    "Channel-to-Channel Skew (ns)": [
        "channel skew", "channel-to-channel skew", "interchannel skew", "skew (ns)",
        "timing skew", "channel delay difference", "output skew"
    ],

    # Power specifications (additional to existing)
    "VDD Supply Voltage Range (V)": [
        "vdd1", "vdd2", "supply voltage side 1", "supply voltage side 2",
        "input side supply", "output side supply", "dual supply voltage"
    ],
    "Dynamic Current Consumption (mA)": [
        "dynamic current", "switching current", "active current", "operating current (dynamic)",
        "current at switching", "dynamic supply current"
    ],
    "Integrated DC-DC Converter": [
        "integrated dc-dc", "isolated power", "power isolation", "dc-dc converter",
        "isolated supply", "power transfer", "integrated power"
    ],

    # === NEW: Level Shifter/Translator Specific ===
    # Channel & direction configuration
    "Direction": [
        "direction", "unidirectional", "bidirectional", "auto-direction", 
        "direction sensing", "auto-direction sensing", "signal direction"
    ],
    "Signal Type": [
        "signal type", "logic signal", "open-drain", "push-pull", "cmos signal",
        "ttl signal", "open-collector", "signal format"
    ],
    "Supported Interfaces": [
        "supported interfaces", "interface support", "i2c support", "spi support",
        "uart support", "gpio support", "protocol support", "bus support"
    ],

    # Voltage compatibility
    "VCCA Range (V)": [
        "vcca", "low side supply", "vcca range", "side a voltage", "input side voltage",
        "vcca min/max", "low voltage side", "vcca supply range"
    ],
    "VCCB Range (V)": [
        "vccb", "high side supply", "vccb range", "side b voltage", "output side voltage", 
        "vccb min/max", "high voltage side", "vccb supply range"
    ],
    "Logic High Threshold (VIH)": [
        "vih", "input high voltage", "logic high threshold", "high level input",
        "input high", "vih min", "logic 1 threshold"
    ],
    "Logic Low Threshold (VIL)": [
        "vil", "input low voltage", "logic low threshold", "low level input",
        "input low", "vil max", "logic 0 threshold"
    ],

    # Performance (additional to existing)
    "Max Data Rate/Bandwidth": [
        "max data rate", "bandwidth", "data bandwidth", "max frequency",
        "signal bandwidth", "data rate capability", "max switching rate"
    ],

    # Electrical characteristics
    "Drive Strength/Output Current": [
        "drive strength", "output drive current", "drive current", "output current capability",
        "drive capability", "sink/source current", "io drive strength"
    ],
    "Pull-up Requirements": [
        "pull-up requirements", "internal pull-up", "external pull-up", "pull-up resistor",
        "pull-up value", "pull-up range", "resistor requirements"
    ],
    "Leakage Current": [
        "leakage current", "input leakage", "output leakage", "off-state current",
        "standby current", "leakage (ua)", "quiescent leakage"
    ],

    # Protection & reliability (additional)
    "Latch-up Immunity": [
        "latch-up immunity", "latch-up protection", "iec compliance", "jedec compliance",
        "latch-up current", "latch-up resistance"
    ],

    # Integration features
    "Integrated Features": [
        "integrated features", "built-in features", "fail-safe", "hot-swap capability",
        "internal pull-ups", "fail-safe io", "hot-swap", "power sequencing"
    ]
}
# === Standardize column names ===
def standardize_columns(df, mapping_dict, threshold=85):
    new_cols = {}
    for col in df.columns:
        found = False
        for std_name, alternatives in mapping_dict.items():
            choices = [std_name] + alternatives
            best_match, score = process.extractOne(str(col), choices)
            if score >= threshold:
                new_cols[col] = std_name
                found = True
                break
        if not found:
            new_cols[col] = col
    return df.rename(columns=new_cols)

# === Combine duplicate columns ===
def combine_duplicate_columns(df):
    """Combine columns with the same name by merging their non-null values"""
    # Get unique column names
    unique_cols = []
    duplicate_groups = {}
    
    for col in df.columns:
        if col not in unique_cols:
            unique_cols.append(col)
            # Find all columns with this name
            same_name_cols = [i for i, c in enumerate(df.columns) if c == col]
            if len(same_name_cols) > 1:
                duplicate_groups[col] = same_name_cols
    
    # If no duplicates, return original dataframe
    if not duplicate_groups:
        return df
    
    # Create new dataframe with combined columns
    new_df = pd.DataFrame()
    
    for col in unique_cols:
        if col in duplicate_groups:
            # Combine all columns with the same name
            col_indices = duplicate_groups[col]
            combined_series = df.iloc[:, col_indices[0]].copy()
            
            # Merge values from duplicate columns
            for idx in col_indices[1:]:
                duplicate_col = df.iloc[:, idx]
                # Fill nulls in combined_series with values from duplicate_col
                combined_series = combined_series.fillna(duplicate_col)
                
                # For non-null values, concatenate if they're different
                mask = combined_series.notna() & duplicate_col.notna() & (combined_series != duplicate_col)
                if mask.any():
                    combined_series.loc[mask] = combined_series.loc[mask].astype(str) + " | " + duplicate_col.loc[mask].astype(str)
            
            new_df[col] = combined_series
        else:
            # Single column, add as is
            new_df[col] = df[col]
    
    return new_df

# === Add manufacturer column ===
def add_manufacturer_column(df, manufacturer_name):
    """Add manufacturer column as the second column if it doesn't exist"""
    # Combine duplicate columns instead of renaming them
    df = combine_duplicate_columns(df)
    
    if "Manufacturer" not in df.columns:
        # Insert manufacturer as second column (after Part Number if it exists)
        if "Part Number" in df.columns:
            part_num_idx = df.columns.get_loc("Part Number")
            df.insert(part_num_idx + 1, "Manufacturer", manufacturer_name)
        else:
            df.insert(0, "Manufacturer", manufacturer_name)
    else:
        # If manufacturer column exists but is empty/null, fill it
        df["Manufacturer"] = df["Manufacturer"].fillna(manufacturer_name)
        # If all values are the same or empty, replace with new value
        if df["Manufacturer"].nunique() <= 1:
            df["Manufacturer"] = manufacturer_name
    return df

# === Load any supported file ===
def load_file(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    
    if ext == '.csv':
        try:
            return pd.read_csv(file_path)
        except Exception as e:
            # Try different encoding if UTF-8 fails
            try:
                return pd.read_csv(file_path, encoding='latin-1')
            except:
                return pd.read_csv(file_path, encoding='cp1252')
    else:
        # Try multiple methods for Excel files
        try:
            return pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            try:
                # Try with xlrd engine for older files
                return pd.read_excel(file_path, engine='xlrd')
            except Exception as e2:
                try:
                    # Try reading as CSV in case it's misnamed
                    return pd.read_csv(file_path)
                except Exception as e3:
                    # If all fails, raise the original Excel error with more info
                    raise Exception(f"Cannot read file. Tried multiple methods. Original error: {str(e)}")

# === Save any supported file ===
def save_file(df, output_path):
    ext = os.path.splitext(output_path)[-1].lower()
    if ext == '.csv':
        df.to_csv(output_path, index=False)
    else:
        df.to_excel(output_path, index=False)

# === Process a single file ===
def process_any_file(file_path, output_dir, add_manufacturer=False, manufacturer_name=""):
    try:
        df = load_file(file_path)
        df_cleaned = standardize_columns(df, STANDARD_PARAMS)
        
        # Add manufacturer column if requested
        if add_manufacturer and manufacturer_name.strip():
            df_cleaned = add_manufacturer_column(df_cleaned, manufacturer_name.strip())
        
        filename = os.path.basename(file_path)
        out_path = os.path.join(output_dir, filename)
        save_file(df_cleaned, out_path)
        return True, filename
    except Exception as e:
        return False, f"{os.path.basename(file_path)}: {e}"

# === Handle single file ===
def select_single_file():
    filepath = filedialog.askopenfilename(filetypes=[
        ("Excel or CSV files", "*.xlsx *.xls *.csv")
    ])
    if not filepath:
        return

    out_dir = filedialog.askdirectory(title="Select Output Folder")
    if not out_dir:
        return

    # Get manufacturer settings from GUI
    add_mfr = add_manufacturer_var.get()
    mfr_name = manufacturer_entry.get()

    if add_mfr and not mfr_name.strip():
        messagebox.showwarning("Warning", "Please enter a manufacturer name or uncheck the option.")
        return

    success, result = process_any_file(filepath, out_dir, add_mfr, mfr_name)
    if success:
        messagebox.showinfo("Success", f"✅ File saved:\n{os.path.join(out_dir, result)}")
    else:
        messagebox.showerror("Error", f"❌ {result}")

# === Handle batch folder ===
def select_folder_for_batch():
    folder_path = filedialog.askdirectory(title="Select Folder with Excel/CSV Files")
    if not folder_path:
        return

    # Get manufacturer settings from GUI
    add_mfr = add_manufacturer_var.get()
    mfr_name = manufacturer_entry.get()

    if add_mfr and not mfr_name.strip():
        messagebox.showwarning("Warning", "Please enter a manufacturer name or uncheck the option.")
        return

    out_dir = os.path.join(folder_path, "standardized_output")
    os.makedirs(out_dir, exist_ok=True)

    processed = 0
    errors = []

    for file in os.listdir(folder_path):
        if file.lower().endswith(('.xlsx', '.xls', '.csv')):
            full_path = os.path.join(folder_path, file)
            success, result = process_any_file(full_path, out_dir, add_mfr, mfr_name)
            if success:
                processed += 1
            else:
                errors.append(result)

    msg = f"✅ Processed: {processed} file(s).\n📁 Saved in: {out_dir}"
    if errors:
        msg += f"\n⚠️ Errors in:\n" + "\n".join(errors)
        messagebox.showwarning("Completed with Errors", msg)
    else:
        messagebox.showinfo("Success", msg)

# === Toggle manufacturer entry field ===
def toggle_manufacturer_entry():
    if add_manufacturer_var.get():
        manufacturer_entry.config(state='normal')
        manufacturer_label.config(foreground='black')
    else:
        manufacturer_entry.config(state='disabled')
        manufacturer_label.config(foreground='gray')

# === GUI Setup ===
def main():
    global add_manufacturer_var, manufacturer_entry, manufacturer_label
    
    root = tk.Tk()
    root.title("🧩 Excel/CSV Parameter Standardizer (Batch Supported)")
    root.geometry("520x400")
    root.resizable(False, False)

    # Title
    title_label = tk.Label(
        root,
        text="Standardize Excel/CSV Column Names\nfor Semiconductor Parameter Tables",
        font=("Helvetica", 13),
        wraplength=420,
        justify="center"
    )
    title_label.pack(pady=25)

    # Manufacturer section
    manufacturer_frame = tk.Frame(root)
    manufacturer_frame.pack(pady=15)

    add_manufacturer_var = tk.BooleanVar()
    manufacturer_checkbox = tk.Checkbutton(
        manufacturer_frame,
        text="Add Manufacturer Column",
        variable=add_manufacturer_var,
        command=toggle_manufacturer_entry,
        font=("Helvetica", 10)
    )
    manufacturer_checkbox.pack()

    # Manufacturer input
    input_frame = tk.Frame(manufacturer_frame)
    input_frame.pack(pady=(5, 0))

    manufacturer_label = tk.Label(input_frame, text="Manufacturer Name:", font=("Helvetica", 9))
    manufacturer_label.pack(side=tk.LEFT)

    manufacturer_entry = tk.Entry(input_frame, width=25, font=("Helvetica", 9))
    manufacturer_entry.pack(side=tk.LEFT, padx=(5, 0))
    manufacturer_entry.config(state='disabled')  # Initially disabled

    # Add some common manufacturer suggestions as placeholder
    manufacturer_entry.insert(0, "e.g., Texas Instruments, Analog Devices, etc.")
    manufacturer_entry.config(foreground='gray')

    def on_entry_click(event):
        if manufacturer_entry.get() == "e.g., Texas Instruments, Analog Devices, etc.":
            manufacturer_entry.delete(0, "end")
            manufacturer_entry.config(foreground='black')

    def on_entry_leave(event):
        if manufacturer_entry.get() == "":
            manufacturer_entry.insert(0, "e.g., Texas Instruments, Analog Devices, etc.")
            manufacturer_entry.config(foreground='gray')

    manufacturer_entry.bind('<FocusIn>', on_entry_click)
    manufacturer_entry.bind('<FocusOut>', on_entry_leave)

    # Main buttons
    btn1 = tk.Button(root, text="📄 Standardize One File", font=("Helvetica", 11), command=select_single_file)
    btn1.pack(pady=10)

    btn2 = tk.Button(root, text="📂 Standardize Folder of Files", font=("Helvetica", 11), command=select_folder_for_batch)
    btn2.pack(pady=10)

    # Instructions
    instructions = tk.Label(
        root,
        text="💡 Tip: Check the manufacturer option to automatically add\na manufacturer column to your standardized files.",
        font=("Helvetica", 9),
        foreground="gray",
        wraplength=400,
        justify="center"
    )
    instructions.pack(pady=(20, 10))

    root.mainloop()

if __name__ == "__main__":
    main()