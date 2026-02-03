🚀 Excel Image Automation & Data Filtering Script
This Python project automates the process of inserting specific images into hundreds of Excel files based on cell values while performing text corrections and categorized file filtering.

🛠 Features
Smart Image Insertion: Automatically detects image IDs from Excel cells (handling VLOOKUP results) and inserts the corresponding .jpg file.

Dynamic Resizing: Resizes images to a specific ratio (2x2.5) while anchoring them to specific cells.

Object Protection: Uses win32com to interact with the Excel GUI, ensuring existing objects like QR codes and logos remain untouched.

Data Cleanup: Automatically finds and replaces specific text strings (e.g., correcting "Ağ karaman" to "Akkaraman").

Categorized Filtering: Sorts processed files into sub-folders based on specific criteria (Breeds: Akkaraman, Morkaraman, Melez).

🧰 Requirements
Windows OS (Required for win32com)

Microsoft Excel installed

Python 3.x

pypiwin32 library

To install the dependency, run:

Bash
pip install pypiwin32
📂 Project Structure
GIRIS_KLASORU: Source folder containing the raw .xlsx files.

RESIM_KLASORU: Folder containing the product/animal images (e.g., 12345.jpg).

ANA_CIKTI: Destination folder where sorted and processed files are saved.

🚀 How to Use
Clone this repository.

Update the folder paths in the script to match your local directory structure.

Run the script:

Bash
python filtr_final.py
📝 License
Distributed under the MIT License. See LICENSE for more information.
