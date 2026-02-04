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
 # Buraimport os
import win32com.client as win32

# --- YOLLAR ---
GIRIS_KLASORU = r"C:\Users\BlueRose\Desktop\Sahib"
RESIM_KLASORU = r"C:\Users\BlueRose\Desktop\Qoyun Şəkil"
CIKTI_KLASORU = r"C:\Users\BlueRose\Desktop\Sahib\Hazir_Exceller"

UZANTI = ".jpg"

if not os.path.exists(CIKTI_KLASORU):
    os.makedirs(CIKTI_KLASORU)

# Excel'i arka planda başlat
excel_app = win32.gencache.EnsureDispatch('Excel.Application')
excel_app.Visible = False
excel_app.DisplayAlerts = False

print("Boyutlar ayarlanıyor (2x2.5) ve operasyon başlıyor...")

for dosya_adi in os.listdir(GIRIS_KLASORU):
    if dosya_adi.endswith(".xlsx") and not dosya_adi.startswith("~$"):
        try:
            dosya_yolu = os.path.join(GIRIS_KLASORU, dosya_adi)
            cikti_yolu = os.path.join(CIKTI_KLASORU, dosya_adi)
            
            wb = excel_app.Workbooks.Open(dosya_yolu)
            ws = wb.ActiveSheet

            # C6'daki rakamı al
            c6_metni = str(ws.Range("C6").Text).strip()
            if not c6_metni:
                wb.Close(False)
                continue
            resim_id = c6_metni.split()[-1]

            # Resim yolları
            normal_yol = os.path.join(RESIM_KLASORU, resim_id + UZANTI)
            yedek_yol = os.path.join(RESIM_KLASORU, f"{resim_id}(1){UZANTI}")

            secilen_yol = None
            if os.path.exists(normal_yol):
                secilen_yol = normal_yol
            elif os.path.exists(yedek_yol):
                secilen_yol = yedek_yol

            if secilen_yol:
                hucre = ws.Range("E33")
                hucre.Value = "" 

                # Resim ekleme (AddPicture parametreleri: Yol, Link, Save, Sol, Üst, En, Boy)
                # İstediğin 2x2.5 oranını yakalamak için rakamları ona göre ayarladım:
                # Excel'de 1 birim yaklaşık 72 pixeldir.
                h_boy = 2 * 72    # Hündürlük: 2
                w_en = 2.5 * 72   # En: 2.5

                pic = ws.Shapes.AddPicture(secilen_yol, False, True, hucre.Left, hucre.Top, w_en, h_boy)
                
                # En-Boy oranını kilidini aç (Böylece tam verdiğimiz ölçü olur)
                pic.LockAspectRatio = False 
                
                wb.SaveAs(cikti_yolu)
                print(f"[TAMAM] {dosya_adi} -> Resim: {resim_id} (2 x 2.5)")
            else:
                print(f"[HATA] Resim bulunamadı: {resim_id}")
            
            wb.Close(True)

        except Exception as e:
            print(f"[HATA] {dosya_adi}: {e}")
            try: wb.Close(False)
            except: pass

excel_app.Quit()
print("\n--- TEBRİKLER ASLANIM! HER ŞEY HAZIR. ---") ")
