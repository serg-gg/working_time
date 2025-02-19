import pandas as pd
import glob
import os
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext

# Определяем порядок месяцев для сортировки
months_order = {
    "Styczeń": 1, "Luty": 2, "Marzec": 3, "Kwiecień": 4, "Maj": 5, "Czerwiec": 6,
    "Lipiec": 7, "Sierpień": 8, "Wrzesień": 9, "Październik": 10, "Listopad": 11, "Grudzień": 12
}

def search_hours():
    search_surname = entry_name.get().strip().lower()
    if not search_surname:
        messagebox.showwarning("Błąd", "Wpisz nazwisko!")
        return
    
    folder_path = "C:\\Users\\uzytkownik\\Desktop\\Nowy folder"
    files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    all_workers = {}

    for file in files:
        if file.startswith(os.path.join(folder_path, "~$")):
            continue
        try:
            df = pd.read_excel(file, skiprows=1)
            required_columns = ["Imię", "Nazwisko", "Rbh dzienne", "Rbh nocne"]
            
            if not all(col in df.columns for col in required_columns):
                continue  
            
            df = df.dropna(subset=["Nazwisko", "Imię"])
            df = df.fillna(0).infer_objects(copy=False)

            # Приводим к единообразному формату (имя и фамилия с заглавной буквы)
            df["Nazwisko"] = df["Nazwisko"].str.strip().str.capitalize()
            df["Imię"] = df["Imię"].str.strip().str.capitalize()

            filtered_data = df[df["Nazwisko"].str.lower() == search_surname]

            if not filtered_data.empty:
                match = re.search(r"(\w+)_.*?(\d{4})", os.path.basename(file))
                if match:
                    month, year = match.groups()
                    year = int(year)
                    month_order = months_order.get(month.capitalize(), 0)
                    
                    if month_order == 0:
                        continue  

                    for _, row in filtered_data.iterrows():
                        imie = row["Imię"]
                        nazwisko = row["Nazwisko"]
                        try:
                            dzienne_godziny = float(row["Rbh dzienne"])  
                            nocne_godziny = float(row["Rbh nocne"])    
                        except ValueError:
                            continue  

                        full_name = f"{imie} {nazwisko}"
                        if full_name not in all_workers:
                            all_workers[full_name] = {}
                        
                        all_workers[full_name][(year, month_order, month)] = [dzienne_godziny, nocne_godziny]

        except PermissionError:
            messagebox.showerror("Błąd", f"Brak dostępu do pliku: {file}\nZamknij Excel!")
        except Exception as e:
            messagebox.showerror("Błąd", f"Problem z plikiem {file}:\n{e}")

    if all_workers:
        result_text = ""
        for full_name, worked_months in all_workers.items():
            result_text += f"👤 Pracownik: {full_name}\n📅 Pracował w miesiącach:\n"
            total_dzienne = total_nocne = 0

            for year, month_order, month in sorted(worked_months.keys()):
                dzienne, nocne = worked_months[(year, month_order, month)]
                total_dzienne += dzienne
                total_nocne += nocne
                result_text += f"- {month} {year}: {dzienne + nocne} h (Dziennych: {dzienne} h, Nocnych: {nocne} h)\n"
            
            total_wszystkie = total_dzienne + total_nocne
            result_text += f"\n📊 Łączne godziny:\n"
            result_text += f"- Dziennych: {total_dzienne} h\n"
            result_text += f"- Nocnych: {total_nocne} h\n"
            result_text += f"- Wszystkich: {total_wszystkie} h\n\n"

        text_result.config(state=tk.NORMAL)
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)
        text_result.config(state=tk.DISABLED)
    else:
        messagebox.showinfo("Brak wyników", f"Nie znaleziono danych dla '{search_surname}'.")

# === ГРАФИЧЕСКИЙ ИНТЕРФЕЙС ===
root = tk.Tk()
root.title("Godziny pracy")
root.geometry("500x400")

tk.Label(root, text="Wpisz nazwisko:", font=("Arial", 12)).pack(pady=5)
entry_name = tk.Entry(root, font=("Arial", 12))
entry_name.pack(pady=5)

btn_search = tk.Button(root, text="Szukaj", font=("Arial", 12), command=search_hours)
btn_search.pack(pady=10)

text_result = scrolledtext.ScrolledText(root, font=("Arial", 12), height=15, width=50, state=tk.DISABLED)
text_result.pack(pady=10)

root.mainloop()
