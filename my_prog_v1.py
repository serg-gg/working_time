import pandas as pd

# Укажи путь к файлу Excel
file_path = "C:\\Users\\uzytkownik\\Desktop\\Nowy folder\\sierpień_ PRACAINFO SP ZOO_ 2024.xlsx"

# Загружаем данные из Excel
df = pd.read_excel(file_path, skiprows=1, usecols="E,F,J,K")

# Проверяем, есть ли нужный столбец
if "Nazwisko" not in df.columns:
    print("Column 'Nazwisko' is not found")
    exit()

# Запрашиваем фамилию у пользователя
search_name = input("Wpisz Nazwisko: ").strip()

# Фильтруем данные по введённой фамилии (без учёта регистра)
filtered_data = df[df["Nazwisko"].str.lower() == search_name.lower()]
# Переменная для данный конкретного сотрудника
person_filtered_data = []
# Проверяем, найден ли сотрудник
if filtered_data.empty:
    print(f"Nazwisko '{search_name}' nie odnaleziono")
else:
    #print(f"Nazwisko - '{search_name}':")
    for column in filtered_data.columns:
        li = filtered_data[column].tolist()
        person_filtered_data.append(li)