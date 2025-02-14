import pandas as pd
import glob
import os
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext

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
            filtered_data = df[df["Nazwisko"].str.lower() == search_name]

            if not filtered_data.empty:
                person_filtered_data.extend(filtered_data.values.tolist())
                match = re.search(r"(\w+)_.*?(\d{4})", os.path.basename(file))
                if match:
                    month, year = match.groups()
                    key = f"{month.capitalize()} {year}"
                    dzienne_godziny = filtered_data.iloc[:, 2].astype(float).sum()
                    nocne_godziny = filtered_data.iloc[:, 3].astype(float).sum()
                    worked_months[key] = [dzienne_godziny, nocne_godziny]

        except PermissionError:
            messagebox.showerror("Błąd", f"Brak dostępu do pliku: {file}\nZamknij Excel!")
        except Exception as e:
            messagebox.showerror("Błąd", f"Problem z plikiem {file}:\n{e}")

    if person_filtered_data:
        imie = person_filtered_data[0][0]
        nazwisko = person_filtered_data[0][1]
        
        result_text = f"👤 Pracownik: {imie} {nazwisko}\n\n📅 Pracował w miesiącach:\n"
        total_dzienne = total_nocne = 0
        for month in sorted(worked_months):
            dzienne, nocne = worked_months[month]
            total_dzienne += dzienne
            total_nocne += nocne
            result_text += f"- {month}: {dzienne + nocne} h (Dziennych: {dzienne} h, Nocnych: {nocne} h)\n"
        
        total_wszystkie = total_dzienne + total_nocne
        result_text += f"\n📊 Łączne godziny:\n"
        result_text += f"- Dziennych: {total_dzienne} h\n"
        result_text += f"- Nocnych: {total_nocne} h\n"
        result_text += f"- Wszystkich: {total_wszystkie} h\n"

        text_result.config(state=tk.NORMAL)
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)
        text_result.config(state=tk.DISABLED)

    else:
        messagebox.showinfo("Brak wyników", f"Nie znaleziono danych dla '{search_name}'.")

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
