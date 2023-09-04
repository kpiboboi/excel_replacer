import os
from openpyxl import load_workbook
import re
from tqdm import tqdm

def process_excel_file(file_path, replace_values):
    try:
        workbook = load_workbook(file_path)
        replacements = 0

        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        for pattern, replacement in replace_values.items():
                            replaced_text, count = re.subn(pattern, replacement, cell.value)
                            if count > 0:
                                cell.value = replaced_text
                                replacements += count

        workbook.save(file_path)

        if replacements > 0:
            return f"Fayl muvaffaqiyatli o'zgartirildi 🟩: {file_path}\nFayldagi almashtirishlar soni: {replacements} ta"
        else:
            return f"Fayl o'zgartirilmadi 🟨: {file_path}\nFayldagi almashtirishlar yo'q ❌"
    except Exception as e:
        return f"Faylni qayta ishlashda xatolik 🟥: {file_path}. Sabab => : {str(e)}"

def main():
    log_file = open("log.txt", "w", encoding="utf-8")
    folder_path = input("Iltimos, .xlsx fayllari joylashgan papka manzilini kiriting 📂: ")

    replace_values = {}
    while True:
        key = input("O'zgartirilishi kerak bo'lgan matnini kiriting (tugatish uchun bo'sh qoldiring) 🔍: ")
        if not key:
            break
        value = input("O'zgartirish uchun yangi qiymatni kiriting 📝: ")
        replace_values[key] = value

    if not replace_values:
        print("O'zgartirish qiymatlari kiritilmadi. Skript yakunlandi")
    else:
        confirmation = input("O'zgartirishni boshlash uchun [ha] ni, bekor qilish uchun [yo'q] ni kiriting: ")
        if confirmation.lower() == 'ha':
            files = [os.path.join(root, file_name) for root, _, file_names in os.walk(folder_path) for file_name in file_names if file_name.endswith('.xlsx')]
            success_count = 0
            fail_count = 0
            log = []

            for file in tqdm(files, desc="FAYLLAR QAYTA ISHLASH JARAYONIDA 🔁:"):
                log_entry = process_excel_file(file, replace_values)
                log.append(log_entry)
                if "muvaffaqiyatli" in log_entry.lower():
                    success_count += 1
                else:
                    fail_count += 1

            for entry in log:
                log_file.write(entry + "\n")
                print(entry)

            log_file.close()
            print("❇️ OPERATSIYA MUVAFAQQIYATLI YAKUNLANDI ❇️")
            print(f"✅ Muvafaqqiyatli operatsiyalar soni: {success_count}")
            print(f"❌ Muvafaqqiyatsiz operatsiyalar soni: {fail_count}")
        else:
            print("⛔️ O'zgartirish bekor qilindi ⛔️")

if __name__ == "__main__":
    main()