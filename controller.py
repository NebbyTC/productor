import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter.filedialog import asksaveasfile
import os
import distutils.util
from decimal import Decimal
import subprocess
import sys
import platform

import openpyxl

from model import Model
from generator import Faktura


PLIK_USTAWIEN = "settings.txt"


class ChangeSettingsWindow(tk.Toplevel):
	""" Klasa okna zmiany ustawień """


	def __init__(self, controller):
		super().__init__(controller.view)
		self.geometry("500x400")
		self.title("Zmień ustawienia")
		self.group(controller.view)
		self.resizable(False, False)

		self.controller = controller

		# Definicje widgetów
		internal_appending_pane = tk.Frame(self)

		seller_panel = tk.LabelFrame(internal_appending_pane, text="Informacje o sprzedawcy")
		misc_panel = tk.LabelFrame(internal_appending_pane, text="Informacje do faktury")

		# Panel o sprzedawcy
		input_welcoming_label = tk.Label(seller_panel, text="Zmień dane o użytkowniku")

		nazwa_label = tk.Label(seller_panel, text="Sprzedawca:")
		self.nazwa_entry = tk.Entry(seller_panel, width=33)
		self.nazwa_entry.insert(tk.END, controller.sprzedawca)

		typ_label = tk.Label(seller_panel, text="Firma Sprzedawcy:")
		self.typ_entry = tk.Entry(seller_panel, width=33)
		self.typ_entry.insert(tk.END, controller.firma_sprzedawcy)

		cena_label = tk.Label(seller_panel, text="Email:")
		self.cena_entry = tk.Entry(seller_panel, width=33)
		self.cena_entry.insert(tk.END, controller.email)

		adres_label = tk.Label(seller_panel, text="Adres pocztowy:")
		self.adres_entry = tk.Entry(seller_panel, width=33)
		self.adres_entry.insert(tk.END, controller.kod_pocztowy)

		ulica_label = tk.Label(seller_panel, text="Ulica:")
		self.ulica_entry = tk.Entry(seller_panel, width=33)
		self.ulica_entry.insert(tk.END, controller.adres)

		input_submit = tk.Button(internal_appending_pane, text="Zatwierdź", command="dbinsert", width=10)
		input_submit.config(command=lambda: self.change_settings())

		dostawa_label = tk.Label(misc_panel, text="Dostawa:                 ")
		self.dostawa_entry = tk.Entry(misc_panel, width=33)
		self.dostawa_entry.insert(tk.END, controller.stawka_vat)

		# Pakowanie widgetów
		internal_appending_pane.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
		seller_panel.pack(side=tk.TOP, fill=tk.X, padx=4, pady=4)
		

		nazwa_label.grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
		self.nazwa_entry.grid(row=1, column=1, padx=4, pady=4)		

		typ_label.grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
		self.typ_entry.grid(row=2, column=1, padx=4, pady=4)

		cena_label.grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
		self.cena_entry.grid(row=3, column=1, padx=4, pady=4)

		adres_label.grid(row=4, column=0, sticky=tk.W, padx=4, pady=4)
		self.adres_entry.grid(row=4, column=1, padx=4, pady=4)

		ulica_label.grid(row=5, column=0, sticky=tk.W, padx=4, pady=4)
		self.ulica_entry.grid(row=5, column=1, padx=4, pady=4)


		misc_panel.pack(side=tk.TOP, fill=tk.X, padx=4, pady=4)

		dostawa_label.grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
		self.dostawa_entry.grid(row=1, column=1, padx=4, pady=4)


		input_submit.pack(side=tk.TOP, anchor=tk.E, padx=4, pady=4)

	
	def change_settings(self):
		""" Zmienia plik ustawien na podstawie widgetów w panelu ustawien """
		
		# Czyści obecne ustawienia
		with open(PLIK_USTAWIEN, "w+", encoding="utf-8") as w:
			w.write("")

		# Zapisuje ustawienia ustawione przez użytkownika
		opcje = {
			"sprzedawca": self.nazwa_entry.get(),
			"firma-sprzedawcy": self.typ_entry.get(),
			"email": self.cena_entry.get(),
			"kod-pocztowy": self.adres_entry.get(),
			"adres": self.ulica_entry.get(),
			"stawka-vat": self.dostawa_entry.get()
		}

		# Zmienia ustawienia w pliku ustawień
		with open(PLIK_USTAWIEN, "w", encoding="utf-8") as w:
			for klucz, wartosc in opcje.items(): 
				w.write(f"{klucz}: '{wartosc}';\n")

		self.destroy()


class UpdateProductWindow(tk.Toplevel):
	""" Klasa okna edycji produktu """


	def __init__(self, controller, _id: int, nazwa: str, typ: str, cena: float, stan: bool):
		super().__init__(controller.view)
		self.title("Edytuj produkt")
		self.geometry("500x300")
		self.group(controller.view)
		#self.resizable(False, False)

		self._id = _id
		self.controller = controller

		# Definicje widgetów
		internal_appending_pane = tk.LabelFrame(self, text="Zmień atrybuty produktu")

		input_welcoming_label = tk.Label(internal_appending_pane, text="Wprowadź dane do bazy danych")

		nazwa_label = tk.Label(internal_appending_pane, text="Nazwa:")
		self.nazwa_entry = tk.Entry(internal_appending_pane, width=33)
		self.nazwa_entry.insert(0, nazwa)

		typ_label = tk.Label(internal_appending_pane, text="Typ:")
		self.typ_entry = tk.Entry(internal_appending_pane, width=33)
		self.typ_entry.insert(0, typ)

		cena_label = tk.Label(internal_appending_pane, text="Cena:")
		self.cena_entry = tk.Entry(internal_appending_pane, width=33)
		self.cena_entry.insert(0, str(cena))
		
		stan_label = tk.Label(internal_appending_pane, text="Stan:")
		self.stan_entry = ttk.Combobox(internal_appending_pane, width=30)
		self.stan_entry["values"] = ("True", "False")
		self.stan_entry.current(not int(distutils.util.strtobool(stan)))
		self.stan_entry.bind("<<ComboboxSelected>>", lambda e: self.focus())

		input_submit = tk.Button(internal_appending_pane, text="Prześlij", command="dbinsert", width=10)
		input_submit.config(command=lambda: self.update())

		vat_label = tk.Label(internal_appending_pane, text="Stawka VAT:")
		self.vat_entry = tk.Entry(internal_appending_pane, width=33)
		self.vat_entry.insert(0, str(Model.get_product_vat(nazwa)))


		# Pakowanie widgetów
		internal_appending_pane.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		input_welcoming_label.grid(row=0, column=0, columnspan=2, padx=4, pady=4)

		nazwa_label.grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
		self.nazwa_entry.grid(row=1, column=1, padx=4, pady=4)		

		typ_label.grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
		self.typ_entry.grid(row=2, column=1, padx=4, pady=4)

		cena_label.grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
		self.cena_entry.grid(row=3, column=1, padx=4, pady=4)

		stan_label.grid(row=4, column=0, sticky=tk.W, padx=4, pady=4)
		self.stan_entry.grid(row=4, column=1, padx=4, pady=4)

		vat_label.grid(row=5, column=0, sticky=tk.W, padx=4, pady=4)
		self.vat_entry.grid(row=5, column=1, padx=4, pady=4)

		input_submit.grid(row=6, column=1, padx=4, pady=4, sticky=tk.E)

		self.mainloop()


	def update(self):
		""" Aktualizauje informacje o produkcie """
		Model.update_product(self._id, self.nazwa_entry.get(), self.typ_entry.get(), float(self.cena_entry.get()), distutils.util.strtobool(self.stan_entry.get()), float(self.vat_entry.get()))
		self.controller.reload_treeview()
		self.destroy()


class Appender(tk.Toplevel):
	""" Klasa odpowiedzialna za wprowadzanie danych z arkusza .xlsx """


	def __init__(self, controller):
		super().__init__(controller.view)
		self.geometry("500x400")
		self.title("Zmień ustawienia")
		self.group(controller.view)
		self.resizable(False, False)

		self.controller = controller

		# Definicje Gui
		sheet_panel = tk.LabelFrame(self, text="Podaj nazwę pliku akrusza")

		arkusz_label = tk.Label(sheet_panel, text="Nazwa arkusza(pelna):")
		self.akrusz_entry = tk.Entry(sheet_panel, width=33)

		wiersz_label = tk.Label(sheet_panel, text="Ilość wierszy:")
		self.wiersz_entry = tk.Entry(sheet_panel, width=33)

		self.stan_label = tk.Label(self, text="")

		submit = tk.Button(self, text="Zatwierdź", command=self.importuj_arkusz)

		# Pakowanie
		sheet_panel.pack(side=tk.TOP, fill=tk.X, padx=4, pady=4)

		arkusz_label.grid(row=0, column=0, padx=4, pady=4, sticky=tk.W)
		self.akrusz_entry.grid(row=0, column=1, padx=4, pady=4, sticky=tk.E)

		wiersz_label.grid(row=1, column=0, padx=4, pady=4, sticky=tk.W)
		self.wiersz_entry.grid(row=1, column=1, padx=4, pady=4, sticky=tk.E)

		self.stan_label.pack(side=tk.TOP)

		submit.pack(side=tk.BOTTOM, anchor=tk.E, padx=4, pady=4)

		self.mainloop()


	def importuj_arkusz(self):
		try:
			plik = self.akrusz_entry.get()
			wiersze = range(int(self.wiersz_entry.get()))
		except ValueError:
			self.stan_label.config(text="Błąd: Wprowadzone dane nie są poprawne.")

		# Wczytywanie akrusza
		try:
			arkusz = openpyxl.load_workbook(filename=plik)
			skoroszyt = arkusz.active
		except Exception:
			self.stan_label.config(text="Błąd: Podany akrusz nie istnieje.")

		# Dodawanie produktów 
		for wiersz in wiersze:
			Model.add_product(skoroszyt[f"B{wiersz+2}"].value, skoroszyt[f"C{wiersz+2}"].value, skoroszyt[f"D{wiersz+2}"].value, distutils.util.strtobool(skoroszyt[f"E{wiersz+2}"].value), skoroszyt[f"F{wiersz+2}"].value)
		
		self.controller.reload_treeview()
		self.stan_label.config(text="Pomyślnie wczytano dane z arkusza.")


class Controller:
	""" Klasa odpowiedzialna za kontrolę danych i interfejsu """


	def __init__(self, view):
		self.view = view


	def start_appender(self):
		""" Metoda uruchamiająca program imprortu z arkusza xlsx """
		if platform.system() == "Windows":
			Appender(self)
		else:
			Appender(self)

		
	def load_treeview(self):
		""" Wstawia wszystkie dane z bazy danych do wiodku górnego """
		records = Model.get_products()

		for row in records:
			self.view.widok.insert("", tk.END, text=row[0], values=(row[1], row[2], "{0:.2f}".format(float(row[3])), bool(row[4])))


	def append_to_treeview(self):
		""" Dodaje wpis do widoku górnego i do bazy danych """
		Model.add_product(self.view.nazwa_entry.get(), self.view.typ_entry.get(), self.view.cena_entry.get(), distutils.util.strtobool(self.view.stan_entry.get()))

		self.view.widok.insert("", tk.END, text=Model.get_last_product()[0], values=(
			self.view.nazwa_entry.get(), 
			self.view.typ_entry.get(), 
			self.view.cena_entry.get(),
			bool(distutils.util.strtobool(self.view.stan_entry.get()))
			)
		)
		self.view.prod_combo['values'] = self.get_product_names()
		
		# Czyszczenie entry widgetów
		self.view.nazwa_entry.delete(0, "end")
		self.view.typ_entry.delete(0, "end")
		self.view.cena_entry.delete(0, "end")
		self.view.stan_entry.delete(0, "end")


	def reload_treeview(self):
		""" Odświeża górny widok """
		self.clear_treeview()
		self.load_treeview()
		self.view.reload_bottom_treeview()


	def del_from_treeview(self):
		""" Usuwa z widoku gónego i z bazy danych na podstawie id """
		Model.delete_product(self.view.deletion_by_id_entry.get())

		# Odświeżanie górnego widoku
		self.view.deletion_by_id_entry.delete(0, tk.END)
		self.view.prod_combo['values'] = self.get_product_names()

		self.view.widok.delete(*self.view.widok.get_children())
		self.load_treeview()


	def clear_treeview(self):
		""" Usuwa wszystkie elementy dolnego i górnego widoku """
		self.view.widok.delete(*self.view.widok.get_children())


	def edit_product(self):
		""" Pozwala na edycje wybranego produktu """

		# Wydobywanie wartości
		_id = self.view.widok_wyszukiwanego.item(self.view.widok_wyszukiwanego.selection())["text"]
		wartosci = self.view.widok_wyszukiwanego.item(self.view.widok_wyszukiwanego.selection())["values"]

		# Wyświetlanie okna edycji
		user_input = UpdateProductWindow(self, _id, wartosci[0], wartosci[1], wartosci[2], wartosci[3])


	def load_settings(self):
		""" Wczytuje ustawienia z pliku ustawień """
		
		# Odczytuje opcje z pliku
		try:
			with open(PLIK_USTAWIEN, "r+", encoding="utf-8") as r:
				tresc = r.read()
		except FileNotFoundError:
			open(PLIK_USTAWIEN, "a+", encoding="utf-8").close()
			with open(PLIK_USTAWIEN, "a+") as w:
				w.write("sprzedawca: '';\nfirma-sprzedawcy: '';\nemail: '';\nkod-pocztowy: '';\nadres: '';\nstawka-vat: '0.0';")
				tresc = w.read()

		linijki = [x.replace("\n", "") for x in tresc.split(";")]
		ustawienia = {}

		for linijka in linijki:
			if linijka:
				opcja, wartosc = linijka.split(":")
				ustawienia[opcja] = eval(wartosc.strip())

		# Ustawianie ustawień właściwe
		self.sprzedawca = ustawienia["sprzedawca"]
		self.firma_sprzedawcy = ustawienia["firma-sprzedawcy"]
		self.email = ustawienia["email"]
		self.kod_pocztowy = ustawienia["kod-pocztowy"]
		self.adres = ustawienia["adres"]
		self.stawka_vat = ustawienia["stawka-vat"]


	def change_settings(self):
		""" Zmienia plik ustawien na podstawie widgetów w panelu ustawien """
		ChangeSettingsWindow(self)


	def get_product_names(self):
		""" Zwraca nazwy wszystkich produktów w bazie """
		return [produkt[1] for produkt in Model.get_products()]


	def add_to_cart(self, prod: str, quantity: int, concerns: str):
		""" Dodaje wpis produktu do faktury """
		self.view.kupione_produkty.insert("", tk.END, text=str(Model.get_product_id(prod)), 
			values=(
				prod, 
				quantity, 
				"{0:.2f}".format(Model.get_product_price(prod)*quantity),
				concerns
			)
		)


	def get_os_path(self, path: str):
		""" Zwraca scieżkę odpowiednią dla systemu operacyjnego """
		if platform.system() == "Windows":
			return path.replace("/", "\\")

		else:
			return path.replace("\\", "/")


	def create_invoice(self):
		""" Tworzy plik .xlsx z fakturą """
		rozszerzenia = [('Dokument PDF', '*.pdf'), ('Skoroszyt programu Excel', '*.xlsx')]
		nazwa = asksaveasfile(filetypes = rozszerzenia).name
		
		nazwa = self.get_os_path(nazwa)
		info = [
			self.view.sprzedawca_entry.get(),
			self.view.firma_sprzedawcy_entry.get(),
			self.view.email_sprzedawcy_entry.get(),
			self.view.adres_sprzedawcy_entry.get(),
			self.view.ulica_sprzedawcy_entry.get(),
			self.view.nabywca_entry.get(),
			self.view.nabywca_kod_pocztowy_entry.get(),
			self.view.nabywca_nr_tel_entry.get(),
			self.view.nabywca_email_entry.get(),
			self.view.typ_dostawy_entry.get(),
			self.view.forma_platnosci_entry.get(),
			self.view.adres_entry.get(),
			self.view.kupione_produkty,
			self.stawka_vat
		]

		plik = Faktura(nazwa, info, self.view.include_vat.get())
		plik.build()
			
