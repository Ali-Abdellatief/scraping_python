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
        self.standard_columns = {  # --- Generic / commercial ---
    "Part Number": [
        "part no", "pn", "device", "type number", "partnumber", "type", "part number", 
        "p/n", "model", "type no", "type_number", "part_number", "typenum", "type_no"
    ],
    "Manufacturer_real": [
        "manufacturer", "mfr", "brand", "company", "supplier", "vendor", "make", "maker"
    ],
    "Manufacturer_real": [
        "manufacturer_real", "real_manufacturer", "actual_manufacturer", "true_manufacturer"
    ],
    "Description": ["desc", "product description"],
    "Lifecycle Status": [
        "status", "life cycle", "marketing status", "product status", "prod status", 
        "product_status", "prod_status", "lifecycle", "active", "obsolete", "nrnd"
    ],
    "Unit Price (1ku)": ["price", "unit price", "price|quantity (usd)"],
    "Package Type": [
        "package", "pkg", "package name", "package type", "package type2", "package_name", 
        "pkg_name", "pack", "packaging", "case", "outline"
    ],
    "Package version": [
        "package version", "pkg version", "package ver", "pkg ver", "pack version", 
        "package_version", "pkg_version", "pack_ver"
    ],
    "Product Link": ["product url", "product page", "link", "web link"],
    "Datasheet Link": [
        "datasheet", "pdf", "pdf data sheet", "datasheet link", "data sheet", 
        "spec sheet", "specification", "specs", "documentation", "technical data"
    ],
    "ECAD Model": ["ecad", "ecad model"],
    "Certifications": ["cert", "compliance", "rohs"],
    "Automotive Grade": [
        "automotive", "aec-q", "automotive qualified", "aec", "qualified", 
        "automotive_qualified", "aec-q101", "auto qualified"
    ],
    "Operating Temp Range (°C)": [
        "temperature range", "operating temperature", "temp range",
        "operating temperature range (°c)", "operating temp range (Â°c)"
    ],
    "Mounting Type": ["mount type", "mounting"],
    "Type": ["type", "category", "product family"],
    "Eval Board Available": ["eval board", "evaluation board"],
    "Reliability Rating": ["reliability"],
    "OPN": [
        "opn", "orderable part number", "order part number", "ordering part number", 
        "order code", "part code", "ordering code", "sales code"
    ],
    "Coupon Available": ["coupon", "discount"],
    "Configuration": [
        "config", "configuration", "conf", "setup", "arrangement", "topology"
    ],
    "Release date": [
        "date", "release date", "release_date", "launch date", "introduction date", 
        "release", "launched"
    ],

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

    # --- MOSFET/Transistor Specific Parameters (from original dictionary) ---
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
    "RDSon [max] @ Tj = 25 °C (mΩ)": [
        "rdson 25c", "rdson@25c", "rdson_25c", "rdson low temp", "low temperature rdson"
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

    # --- Switch/load specific (updated) ---
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

    # === Isolation Amplifier Specific ===
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
        "bandwidth (hz)", "bandwidth (khz)", "bandwidth (mhz)", "-3db frequency","BW -3 dB (typ)"
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
    "Supply Voltage Type (V)": [
        "supply voltage type", "single-ended", "dual supply", "split supply",
        "single supply type", "dual supply type", "supply voltage type (v)","Vs Type"
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

    # === Transformer Driver Specific ===
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

    # === Digital Isolator Specific ===
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

    # === Level Shifter/Translator Specific ===
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
            patterns = self.standard_columns.get(standard_col, [])
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
        """Process an Excel or CSV file and map its columns"""
        try:
            # Check if it's a CSV file
            is_csv = file_path.lower().endswith('.csv')
            
            # Read file based on type
            if is_csv:
                df = pd.read_csv(file_path)
                print(f"Loaded CSV file: {file_path}")
            else:
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
                if output_file.lower().endswith('.csv'):
                    mapped_df.to_csv(output_file, index=False)
                    print(f"\nMapped dataframe saved to: {output_file}")
                else:
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
        """Open file dialog to select input Excel or CSV file"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        file_path = filedialog.askopenfilename(
            title="Select Input File (Excel or CSV)",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return file_path
    
    def get_sheet_names(self, file_path):
        """Get all sheet names from Excel file or return None for CSV"""
        try:
            # Check if it's a CSV file
            if file_path.lower().endswith('.csv'):
                return ["CSV Data"]  # CSV files don't have sheets, so we create a dummy name
            
            # For Excel files, get actual sheet names
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            print(f"Error reading file: {str(e)}")
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
        
        # Check if it's a CSV file (single sheet with dummy name)
        if len(sheet_names) == 1 and sheet_names[0] == "CSV Data":
            return "CSV Data"
        
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
        """Dialog to select output Excel file location and name"""
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
    
    def select_csv_output_location(self, default_name):
        """Dialog to select output CSV file location and name"""
        root = tk.Tk()
        root.withdraw()
        
        output_path = filedialog.asksaveasfilename(
            title="Save Mapped CSV File As",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return output_path
    
    def show_preview_dialog(self, df_formulas, column_mapping, unmapped_columns, mapping_scores):
        """Show preview of column mappings, allow editing, and copy hyperlinks."""
        root = tk.Tk()
        root.title("Column Mapping Preview & Edit")
        root.geometry("850x700")

        editable_mapping = column_mapping.copy()
        proceed = [False]
        add_manufacturer_real = tk.BooleanVar(value=False)

        def on_proceed():
            proceed[0] = True
            # Store the manufacturer selection and checkbox state
            proceed.append(add_manufacturer_real.get())
            proceed.append(manufacturer_var.get())
            root.quit()

        def on_cancel():
            root.quit()

        def on_edit(event):
            """Handle editing of the 'Mapped To' column."""
            item_id = mapped_tree.focus()
            if not item_id:
                return

            item = mapped_tree.item(item_id)
            original_col = item['values'][0]
            
            new_name = simpledialog.askstring(
                "Edit Mapping",
                f"Enter new standard name for '{original_col}':",
                initialvalue=editable_mapping.get(original_col, "")
            )

            if new_name is not None:
                editable_mapping[original_col] = new_name
                mapped_tree.item(item_id, values=(original_col, new_name, "Manual Edit"))

        def copy_hyperlink(event):
            item_id = mapped_tree.focus()
            if not item_id:
                return

            # Get the index of the selected row
            item_index = mapped_tree.index(item_id)

            item = mapped_tree.item(item_id)
            original_col = item['values'][0]
            mapped_col = item['values'][1]

            if "link" in mapped_col.lower() or "url" in mapped_col.lower():
                try:
                    hyperlink = df_formulas.loc[item_index, original_col]
                    root.clipboard_clear()
                    root.clipboard_append(hyperlink)
                    messagebox.showinfo("Copied", f"Copied to clipboard:\n{hyperlink}")
                except (IndexError, KeyError):
                    messagebox.showwarning("Empty Cell", f"The cell in column '{original_col}' at this row is empty or could not be found.")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not copy link: {e}")
            else:
                messagebox.showinfo("Not a Link", "This column is not identified as a hyperlink column.")

        def show_context_menu(event):
            context_menu.post(event.x_root, event.y_root)

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        if editable_mapping:
            mapped_frame = ttk.Frame(notebook)
            notebook.add(mapped_frame, text=f"Mapped Columns ({len(editable_mapping)})")

            instructions_frame = tk.Frame(mapped_frame)
            instructions_frame.pack(pady=5)
            tk.Label(instructions_frame, text="Double-click a row to edit.", font=("Arial", 9, "italic")).pack(side=tk.LEFT, padx=5)
            tk.Label(instructions_frame, text="Right-click for options.", font=("Arial", 9, "italic")).pack(side=tk.LEFT, padx=5)

            columns = ("Original", "Mapped To", "Confidence")
            mapped_tree = ttk.Treeview(mapped_frame, columns=columns, show="headings", height=15)

            for col in columns:
                mapped_tree.heading(col, text=col)
                mapped_tree.column(col, width=270)

            for original, mapped in editable_mapping.items():
                score = mapping_scores.get(original, "N/A")
                confidence = f"{score:.1f}%" if isinstance(score, (int, float)) else score
                mapped_tree.insert("", tk.END, values=(original, mapped, confidence))

            mapped_tree.bind("<Double-1>", on_edit)
            mapped_tree.bind("<Button-3>", show_context_menu)

            context_menu = tk.Menu(root, tearoff=0)
            context_menu.add_command(label="Copy Hyperlink", command=lambda: copy_hyperlink(None))

            mapped_scrollbar = ttk.Scrollbar(mapped_frame, orient=tk.VERTICAL, command=mapped_tree.yview)
            mapped_tree.configure(yscrollcommand=mapped_scrollbar.set)
            
            mapped_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            mapped_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

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

        # Add Manufacturer_real checkbox
        manufacturer_frame = tk.Frame(root)
        manufacturer_frame.pack(pady=5)
        
        tk.Checkbutton(manufacturer_frame, text="Add 'Manufacturer_real' column", 
                      variable=add_manufacturer_real, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Manufacturer selection dropdown
        manufacturer_label = tk.Label(manufacturer_frame, text="Select Manufacturer:", font=("Arial", 10))
        manufacturer_label.pack(side=tk.LEFT, padx=10)
        
        manufacturer_var = tk.StringVar(value="Texas Instruments")
        manufacturer_dropdown = ttk.Combobox(manufacturer_frame, textvariable=manufacturer_var, 
                                           values=["TI", "ADI", "Infineon", 
                                                  "ST", "Onsemi", "Microchip", 
                                                  "Broadcom", "NXP", "Toshiba", "Renesas Electronics",
                                                  "Seiko Epson", "Marvell Technology", "Maxim Integrated", 
                                                  "Mitsubishi Electric", "ABB", "Fuji Electric", "Semikron Danfoss",
                                                  "Vishay", "Hitachi", "Littelfuse", "ROHM", "Custom..."])
        manufacturer_dropdown.pack(side=tk.LEFT, padx=5)
        
        def on_manufacturer_change(event):
            if manufacturer_var.get() == "Custom...":
                custom_name = simpledialog.askstring("Custom Manufacturer", "Enter manufacturer name:")
                if custom_name:
                    manufacturer_var.set(custom_name)
                else:
                    manufacturer_var.set("TI")
        
        manufacturer_dropdown.bind("<<ComboboxSelected>>", on_manufacturer_change)
        
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Proceed with Mapping", command=on_proceed, 
                 bg="green", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        summary_text = f"Ready to map {len(editable_mapping)} columns"
        if unmapped_columns:
            summary_text += f" ({len(unmapped_columns)} will remain unmapped)"
        
        tk.Label(root, text=summary_text, font=("Arial", 10)).pack(pady=5)
        
        root.mainloop()
        root.destroy()
        
        if proceed[0]:
            return proceed[0], editable_mapping, proceed[1], proceed[2]  # proceed, mapping, add_manufacturer_real, manufacturer_name
        else:
            return proceed[0], None, None, None
    
    def run_interactive(self):
        """Run the interactive version continuously"""
        print("=== Interactive MOSFET Column Mapper ===\n")
        print("🔄 Program will run continuously - process multiple files without restarting!\n")
        
        # Store settings for reuse
        saved_threshold = 70
        saved_manufacturer = "Texas Instruments"
        saved_add_manufacturer_real = False
        
        while True:
            print("=" * 60)
            print("📁 Ready to process next Excel file...")
            print("=" * 60)
            
            # Step 1: Select input file
            print("\nStep 1: Select input Excel file...")
            input_file = self.select_input_file()
            
            if not input_file:
                print("No file selected. Exiting program.")
                break
            
            print(f"Selected: {os.path.basename(input_file)}")
            
            # Step 2: Get sheet names and let user select
            print("\nStep 2: Reading Excel file...")
            sheet_names = self.get_sheet_names(input_file)
            
            if not sheet_names:
                print("Could not read Excel file or no sheets found.")
                continue
            
            selected_sheet = self.select_sheet_dialog(sheet_names)
            if not selected_sheet:
                print("No sheet selected. Skipping this file.")
                continue
            
            print(f"Selected sheet: {selected_sheet}")
            
            # Step 3: Get threshold (with option to reuse previous)
            print("\nStep 3: Set matching threshold...")
            if saved_threshold is not None:
                reuse_threshold = messagebox.askyesno("Reuse Settings", 
                    f"Use previous threshold ({saved_threshold}%)?\n\nClick 'Yes' to reuse\nClick 'No' to set new threshold")
                if reuse_threshold:
                    threshold = saved_threshold
                    print(f"Reusing previous threshold: {threshold}%")
                else:
                    threshold = self.get_threshold_dialog()
                    saved_threshold = threshold
            else:
                threshold = self.get_threshold_dialog()
                saved_threshold = threshold
            
            print(f"Threshold set to: {threshold}%")
            
            # Step 4: Load and analyze the file
            print("\nStep 4: Analyzing columns...")
            try:
                # Check if it's a CSV file
                is_csv = input_file.lower().endswith('.csv')
                
                if is_csv:
                    # For CSV files, read directly with pandas
                    df_values = pd.read_csv(input_file)
                    # Create a dummy df_formulas for CSV (since no formulas)
                    df_formulas = df_values.copy()
                    print("Processing CSV file...")
                else:
                    # For Excel files, read with openpyxl to get formulas
                    import openpyxl
                    workbook = openpyxl.load_workbook(input_file, data_only=False)
                    sheet = workbook[selected_sheet]
                    data = []
                    for row in sheet.iter_rows(values_only=True):
                        data.append(list(row))
                    df_formulas = pd.DataFrame(data[1:], columns=data[0])

                    # Read with pandas for values
                    df_values = pd.read_excel(input_file, sheet_name=selected_sheet)
                    print("Processing Excel file...")

                column_mapping, unmapped_columns, mapping_scores = self.mapper.map_columns(df_values, threshold)
                
                print(f"Found {len(df_values.columns)} columns total")
                print(f"Mapped: {len(column_mapping)} columns")
                print(f"Unmapped: {len(unmapped_columns)} columns")
                
            except Exception as e:
                print(f"Error processing file: {str(e)}")
                messagebox.showerror("Error", f"Error processing file: {str(e)}")
                continue
            
            # Step 5: Show preview and get confirmation
            print("\nStep 5: Showing mapping preview...")
            result = self.show_preview_dialog(df_formulas, column_mapping, unmapped_columns, mapping_scores)
            proceed, final_mapping, add_manufacturer_real, manufacturer_name = result
            
            if not proceed:
                print("Operation cancelled by user. Skipping this file.")
                continue
            
            # Save settings for reuse
            saved_manufacturer = manufacturer_name
            saved_add_manufacturer_real = add_manufacturer_real
            
            # Step 6: Select output location
            print("\nStep 6: Select output location...")
            input_name = Path(input_file).stem
            is_csv = input_file.lower().endswith('.csv')
            
            if is_csv:
                default_output = f"{input_name}_mapped.csv"
                # For CSV files, show CSV save dialog
                output_file = self.select_csv_output_location(default_output)
            else:
                default_output = f"{input_name}_mapped.xlsx"
                # For Excel files, show Excel save dialog
                output_file = self.select_output_location(default_output)
            
            if not output_file:
                print("No output location selected. Skipping this file.")
                continue
            
            print(f"Output file: {os.path.basename(output_file)}")
            
            # Step 7: Process and save
            print("\nStep 7: Processing and saving...")
            try:
                mapped_df = df_values.rename(columns=final_mapping)
                
                # Add Manufacturer_real column if requested
                if add_manufacturer_real:
                    mapped_df['Manufacturer_real'] = manufacturer_name
                    print(f"Added 'Manufacturer_real' column filled with: {manufacturer_name}")
                
                # Save based on output file extension
                if output_file.lower().endswith('.csv'):
                    mapped_df.to_csv(output_file, index=False)
                    print("Saved as CSV file")
                else:
                    mapped_df.to_excel(output_file, index=False)
                    print("Saved as Excel file")
                
                print(f"\n✅ SUCCESS!")
                print(f"Mapped file saved as: {output_file}")
                print(f"Columns mapped: {len(final_mapping)}")
                if add_manufacturer_real:
                    print(f"Manufacturer_real column added with value: {manufacturer_name}")
                print(f"Threshold used: {threshold}%")
                
                success_message = f"File processed successfully!\n\nMapped {len(final_mapping)} columns\nSaved to: {os.path.basename(output_file)}"
                if add_manufacturer_real:
                    success_message += f"\nManufacturer_real column added with value: {manufacturer_name}"
                
                messagebox.showinfo("Success", success_message)
                
            except Exception as e:
                print(f"Error saving file: {str(e)}")
                messagebox.showerror("Error", f"Error saving file: {str(e)}")
                continue
            
            # Ask if user wants to continue
            print("\n" + "=" * 60)
            continue_processing = messagebox.askyesno("Continue Processing", 
                "File processed successfully! 🎉\n\nWould you like to process another Excel file?\n\n"
                "Click 'Yes' to continue with next file\n"
                "Click 'No' to exit program")
            
            if not continue_processing:
                print("\n👋 Exiting program. Goodbye!")
                break
        
        print("\nProgram finished.")

def main():
    """Main function - choose between command line and interactive mode"""
    if len(sys.argv) > 1:
        # Command line mode (original functionality)
        file_path = sys.argv[1]
        sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
        threshold = int(sys.argv[3]) if len(sys.argv) > 3 else 70
        output_file = sys.argv[4] if len(sys.argv) > 4 else None
        
        mapper = MOSFETColumnMapper()
        
        # Check if it's a CSV file
        if file_path.lower().endswith('.csv'):
            print(f"Processing CSV file: {file_path}")
            # For CSV files, we don't need sheet_name
            result = mapper.process_excel_file(file_path, None, threshold, output_file)
        else:
            print(f"Processing Excel file: {file_path}")
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