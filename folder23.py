import os

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
    "Seiko Epson"
]

# Base path where folders will be created (current directory)
base_path = os.getcwd()

for folder in folders:
    folder_path = os.path.join(base_path, folder)
    os.makedirs(folder_path, exist_ok=True)  # Create folder if it doesn't exist
    print(f"Created folder: {folder_path}")

print("✅ All folders are ready.")
