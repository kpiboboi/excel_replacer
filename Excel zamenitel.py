import os
from openpyxl import load_workbook
import re

# Papkaga yo'l:
folder_path = input("Iltimos, .xlsx fayllar joylashgan papkani manzilini kiriting: ") 

# Almashtirish (zmena) uchun:
replace_values = {}
while True:
    key = input("Kiriting almashtirish uchun matn kiriting (bo'sh qoldirish bilan tugatish): ")
    if not key:
        break
    value = input(f"Matning almashtirish qiymatini kiriting: ")
    replace_values[key] = value

if not replace_values:
    print("Almashtirish qiymatlari kiritilmagan. Skript to'xtatiladi.")
else:
    confirmation = input("Almashtirishni boshlash uchun 'yes' yoki bekor qilish uchun 'no' deb yozing: ")
    if confirmation.lower() == 'yes':

# Excel fayl obrabotka funksiyasi
        def process_excel_file(file_path):
            try:
                workbook = load_workbook(file_path)

                for sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]

                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value is not None and isinstance(cell.value, str):
                                for pattern, replacement in replace_values.items():
                                    cell.value = re.sub(pattern, replacement, cell.value)
                workbook.save(file_path)
                print(f"Fayl muvafaqqiyatli o'zgartirildi 🟩 : {file_path}")
            except Exception as e:
                print(f"Fayl obrabotkasida xatolik mavjud 🟥 : {file_path}. Sabab => : {str(e)}")


# Rekursiv fayl va papkalardan o'tuvchi funksiya
def process_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                process_excel_file(file_path)


# Funksiyani chaqirish
process_folder(folder_path)
print("✅✅✅ Fayllar muvafaqqiyatli o'zgartirildi ✅✅✅")