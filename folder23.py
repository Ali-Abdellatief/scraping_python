import os
import tkinter as tk
from tkinter import filedialog

# Hide the main Tkinter window
root = tk.Tk()
root.withdraw()

# Ask the user to select a directory
base_path = filedialog.askdirectory(title="Select the base folder to create subfolders")
if not base_path:
    print("No folder selected. Exiting.")
    exit(1)

# List of folder names
folders = [
    "Texas Instruments",
    "Analog Devices",
    "Infineon Technologies",
    "STMicroelectronics",
    "Onsemi",
    "Microchip Technology",
    "Broadcom",
    "NXP Semiconductors",
    "Toshiba",
    "Renesas Electronics",
    "Seiko Epson",
    "Marvell Technology",
    "Maxim Integrated",
    "Mitsubishi Electric",
    "ABB",
    "Fuji Electric",
    "Semikron Danfoss",
    "Vishay",
    "Hitachi",
    "Littelfuse",
    "ROHM"
]

for folder in folders:
    folder_path = os.path.join(base_path, folder)
    os.makedirs(folder_path, exist_ok=True)  # Create folder if it doesn't exist
    print(f"Created folder: {folder_path}")

print("✅ All folders are ready.")
