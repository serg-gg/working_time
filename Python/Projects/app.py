import pandas as pd
import glob
import os
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext

# Маппинг названий месяцев на числа для сортировки
MONTHS_MAP = {
    "styczeń": 1, "luty": 2, "marzec": 3, "kwiecień": 4,
    "maj": 5, "czerwiec": 6, "lipiec": 7, "sierpień": 8,
    "wrzesień": 9, "październik": 10, "listopad": 11, "grudzień": 12
}

def search_hours():
    search_name = entry_name.get().strip().lower()
    if not search_name:
        messagebox.showwarning("Błąd", "Wpisz nazwisko!")
        return
    
    folder_path = "C:\\Users\\uzytkownik\\Desktop\\Nowy folder"
    files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    person_filtered_data = []
    worked_months = {}

    for file in files:
        if file.startswith(os.path.join(folder_path, "~$")):
            continue
        try:
            df = pd.read_excel(file, skiprows=1, usecols="E,F,J,K")
            if "Nazwisko" not in df.columns:
                continue  
            df = df.dropna(subset=["Nazwisko"])
            df = df.fillna(0).infer_objects(copy=False)
            df["Nazwisko"] = df["Nazwisko"].str.lower()
            
            filtered_data = df[df["Nazwisko"] == search_name]
            
            if not filtered_data.empty:
                person_filtered_data.extend(filtered_data.values.tolist())
                match = re.search(r"(\w+)_\s*([^_]+)_\s*(\d{4})", os.path.basename(file))
                if match:
                    month, company, year = match.groups()
                    month = month.lower()
                    key = (int(year), MONTHS_MAP.get(month, 0), f"{month.capitalize()} {year} ({company})")
                    dzienne_godziny = filtered_data.iloc[:, 2].astype(float).sum()
                    nocne_godziny = filtered_data.iloc[:, 3].astype(float).sum()
                    
                    if key not in worked_months:
                        worked_months[key] = [0, 0, company]
                    worked_months[key][0] += dzienne_godziny
                    worked_months[key][1] += nocne_godziny

        except PermissionError:
            messagebox.showerror("Błąd", f"Brak dostępu do pliku: {file}\nZamknij Excel!")
        except Exception as e:
            messagebox.showerror("Błąd", f"Problem z plikiem {file}:\n{e}")

    if person_filtered_data:
        imie = person_filtered_data[0][0].capitalize()
        nazwisko = person_filtered_data[0][1].capitalize()
        
        companies = {data[2] for data in worked_months.values()}
        if len(companies) == 1:
            company_label = f"\n🏢 Spółka: {list(companies)[0]}\n"
        else:
            company_label = ""

        result_text = f"👤 Pracownik: {imie} {nazwisko}{company_label}\n\n📅 Pracował w miesiącach:\n"
        total_dzienne = total_nocne = 0
        company_hours = {}
        
        for year, month_num, month_label in sorted(worked_months):
            dzienne, nocne, company = worked_months[(year, month_num, month_label)]
            total_dzienne += dzienne
            total_nocne += nocne
            result_text += f"- {month_label}: {dzienne + nocne} h (Dziennych: {dzienne} h, Nocnych: {nocne} h)\n"
            
            if company not in company_hours:
                company_hours[company] = [0, 0]
            company_hours[company][0] += dzienne
            company_hours[company][1] += nocne
        
        total_wszystkie = total_dzienne + total_nocne
        result_text += f"\n📊 Łączne godziny:\n"
        result_text += f"- Dziennych: {total_dzienne} h\n"
        result_text += f"- Nocnych: {total_nocne} h\n"
        result_text += f"- Wszystkich: {total_wszystkie} h\n"
        
        if len(company_hours) > 1:
            result_text += "\n🏢 Podział na spółki:\n"
            for company, (dzienne, nocne) in company_hours.items():
                result_text += f"- {company}: {dzienne + nocne} h (Dziennych: {dzienne} h, Nocnych: {nocne} h)\n"

        text_result.config(state=tk.NORMAL)
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)
        text_result.config(state=tk.DISABLED)

    else:
        messagebox.showinfo("Brak wyników", f"Nie znaleziono danych dla '{search_name}'.")

# === ГРАФИЧЕСКИЙ ИНТЕРФЕЙС ===
root = tk.Tk()
root.title("Godziny pracy")
root.geometry("700x600")

tk.Label(root, text="Wpisz nazwisko:", font=("Arial", 12)).pack(pady=5)
entry_name = tk.Entry(root, font=("Arial", 12))
entry_name.pack(pady=5)

btn_search = tk.Button(root, text="Szukaj", font=("Arial", 12), command=search_hours)
btn_search.pack(pady=10)

text_result = scrolledtext.ScrolledText(root, font=("Arial", 12), height=20, width=70, state=tk.DISABLED)
text_result.pack(pady=10)

root.mainloop()

