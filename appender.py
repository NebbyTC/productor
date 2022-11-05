import openpyxl as excel
import sqlite3


# Pobieranie informacji od użytkownika
print("Podaj nazwę pliku do zaimportowania danych do bazy danych (z rozszerzeniem)")
nazwa_pliku = input(">>>")
print("podaj ilość importowanych danych (wierszy)")
ilosc_wierszy = int(input(">>>"))


# Wczytywanie 
try:
	w = excel.load_workbook(filename=nazwa_pliku)
	sheet = w.active
except Exception:
	print("Błąd: Nie znaleziono takiego pliku w lokalizacji bazy danych.")
	input()
	exit()


# Zapisywanie danych do bazy danych
db = sqlite3.connect("sodb.db")
mycursor = db.cursor()

for i in range(ilosc_wierszy):
	print(sheet[f"B{i+2}"].value, sheet[f"C{i+2}"].value, sheet[f"D{i+2}"].value, sheet[f"E{i+2}"].value)
	
	mycursor.execute(
		f"""INSERT INTO Spismama VALUES ("{sheet[f"B{i+2}"].value}", "{sheet[f"C{i+2}"].value}", "{sheet[f"D{i+2}"].value}", "{sheet[f"E{i+2}"].value}")"""
	)
db.commit()

print("Pomyślnie zaimportowano dane do bazy danych.")
input()
