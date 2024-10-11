import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror, showinfo
import os
import distutils.util
from decimal import Decimal
import subprocess
import sys
import platform
import datetime
from dataclasses import dataclass

import openpyxl

import components as piktk
from model import Model
from generator import Faktura


class ChangeSettingsWindow(tk.Toplevel):
	""" Klasa okna zmiany ustawień """


	def __init__(self, controller):
		super().__init__(controller.view)
		self.geometry("500x400")
		self.title("Zmień ustawienia")
		self.group(controller.view)
		self.resizable(False, False)

		self.controller = controller

		# Definicje Listenerów
		def tab_changed(e):
			self.focus()

		# Definicje stylów
		style = ttk.Style(self)
		style.configure('lefttab.TNotebook', tabposition='wn')

		# Definicje widgetów
		notebook = ttk.Notebook(self, style='lefttab.TNotebook')
		notebook.bind('<<NotebookTabChanged>>', tab_changed)

		# Karta ustawień sprzedawcy
		sprzedawca_tab = tk.Frame(notebook)

		sprzedawca_caption = tk.Label(sprzedawca_tab, text="Dane sprzedawcy")

		self.sprzedawca_entry = piktk.LabeledEntry(sprzedawca_tab, caption="Sprzedawca")
		self.firma_sprzedawcy = piktk.LabeledEntry(sprzedawca_tab, caption="Firma sprzedawcy")
		self.email_sprzedawcy = piktk.LabeledEntry(sprzedawca_tab, caption="Email sprzedawcy")
		self.adres_sprzedawcy = piktk.LabeledEntry(sprzedawca_tab, caption="Adres sprzedawcy")
		self.ulica_sprzedawcy = piktk.LabeledEntry(sprzedawca_tab, caption="Ulica sprzedawcy ")


		sprzedawca_caption.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		piktk.Separator(sprzedawca_tab).pack(fill=tk.X, padx=4, pady=4)

		self.sprzedawca_entry.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.firma_sprzedawcy.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.email_sprzedawcy.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.adres_sprzedawcy.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.ulica_sprzedawcy.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)

		sprzedawca_tab.pack(fill=tk.BOTH)


		# Panel zatwierdzenia
		panel_zatwierdzania = tk.LabelFrame(self)

		submit_button = tk.Button(panel_zatwierdzania, text="Zatwierdź", command=self.change_settings)
		cancel_button = tk.Button(panel_zatwierdzania, text="Anuluj", command=self.destroy)

		submit_button.pack(side=tk.RIGHT, padx=4, pady=4)
		cancel_button.pack(side=tk.RIGHT, padx=4, pady=4)

		panel_zatwierdzania.pack(side=tk.BOTTOM, fill=tk.X, padx=4, pady=4)


		# Karta ustawień dostawy
		dostawa_tab = tk.Frame(notebook)

		dostawa_caption = tk.Label(dostawa_tab, text="Informacje o dostawie")

		self.uwzglednij_dostawe = piktk.Checkbox(dostawa_tab, caption="Uwzględnij kwotę dostawy w fakturze")
		self.kwota_dostawy = piktk.LabeledEntry(dostawa_tab, caption="Kwota dostawy(PLN)")

		dostawa_caption.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		piktk.Separator(dostawa_tab).pack(fill=tk.X, padx=4, pady=4)

		self.uwzglednij_dostawe.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.kwota_dostawy.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)

		dostawa_tab.pack(fill=tk.BOTH)

		# Dodawanie kart do notebooka
		notebook.add(sprzedawca_tab, text="Sprzedawca")
		notebook.add(dostawa_tab, text="Dostawa")

		notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

		# Listenery komponentów
		def dostawa_changed():
			if self.uwzglednij_dostawe.get():
				self.kwota_dostawy.enable()

			else:
				self.kwota_dostawy.disable()

		self.uwzglednij_dostawe.config(command = dostawa_changed)

		self.load_settings()

		self.mainloop()
	
	def change_settings(self):
		""" Zmienia plik ustawien na podstawie widgetów w panelu ustawien """
		
		# Zapisuje ustawienia ustawione przez użytkownika
		opcje = {
			"sprzedawca": self.sprzedawca_entry.get(),
			"firma-sprzedawcy": self.firma_sprzedawcy.get(),
			"email": self.email_sprzedawcy.get(),
			"kod-pocztowy": self.adres_sprzedawcy.get(),
			"adres": self.ulica_sprzedawcy.get(),
			"dostawa": self.kwota_dostawy.get(),
			"uwzglednij-dostawe": self.uwzglednij_dostawe.get()
		}

		Model.update_settings(opcje)
		self.destroy()

	def load_settings(self):
		""" Wczytuje ustawienia z pliku ustawień """
		
		# Odczytuje opcje z pliku
		ustawienia = Model.get_settings()

		# Ustawianie ustawień właściwe
		self.sprzedawca_entry.insert(ustawienia["sprzedawca"])
		self.firma_sprzedawcy.insert(ustawienia["firma-sprzedawcy"])
		self.email_sprzedawcy.insert(ustawienia["email"])
		self.adres_sprzedawcy.insert(ustawienia["kod-pocztowy"])
		self.ulica_sprzedawcy.insert(ustawienia["adres"])
		self.kwota_dostawy.insert(ustawienia["dostawa"])

		if bool(distutils.util.strtobool(ustawienia["uwzglednij-dostawe"])):
			self.uwzglednij_dostawe.set(True)
			self.kwota_dostawy.enable()

		else:
			self.uwzglednij_dostawe.set(False)
			self.kwota_dostawy.disable()


class UpdateProductWindow(tk.Toplevel):
	""" Klasa okna edycji produktu """


	def __init__(self, controller, _id: int, nazwa: str, typ: str, cena: float, stan: bool):
		super().__init__(controller.view)
		self.title("Edytuj produkt")
		self.geometry("350x250")
		self.group(controller.view)
		#self.resizable(False, False)

		self._id = _id
		self.controller = controller

		# Definicje widgetów
		entry_pane = tk.LabelFrame(self, text="Dane produktu")

		internal_pane = tk.Frame(entry_pane)

		self.nazwa_entry = piktk.LabeledEntry(internal_pane, caption="Nazwa", spacing=75)
		self.typ_entry = piktk.LabeledEntry(internal_pane, caption="Typ", spacing=75)
		self.cena_entry = piktk.LabeledEntry(internal_pane, caption="Cena", spacing=75)
		self.stan_entry = piktk.LabeledCombobox(internal_pane, options=["True", "False"], caption="Stan", spacing=75)
		self.vat_entry = piktk.LabeledEntry(internal_pane, caption="Vat", spacing=75)
		self._load_data(nazwa, typ, cena, stan)


		panel_zatwierdzania = tk.LabelFrame(self)

		submit_button = tk.Button(panel_zatwierdzania, text="Zatwierdź", command=self._update)
		cancel_button = tk.Button(panel_zatwierdzania, text="Anuluj", command=self.destroy)



		entry_pane.pack(side=tk.TOP, fill=tk.BOTH, padx=4, pady=4, expand=1)
		internal_pane.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		self.nazwa_entry.pack(side=tk.TOP, padx=4, pady=4)
		self.typ_entry.pack(side=tk.TOP, padx=4, pady=4)
		self.cena_entry.pack(side=tk.TOP, padx=4, pady=4)
		self.stan_entry.pack(side=tk.TOP, padx=4, pady=4)
		self.vat_entry.pack(side=tk.TOP, padx=4, pady=4)


		submit_button.pack(side=tk.RIGHT, padx=4, pady=4)
		cancel_button.pack(side=tk.RIGHT, padx=4, pady=4)

		panel_zatwierdzania.pack(side=tk.BOTTOM, fill=tk.X, padx=4, pady=4)
		self.mainloop()


	def _load_data(self, nazwa: str, typ: str, cena: float, stan: bool):
		""" Wstawia informacje o produkcie do pól tekstowych """
		self.nazwa_entry.insert(str(nazwa))
		self.typ_entry.insert(str(typ))
		self.cena_entry.insert(str(cena))
		self.stan_entry.select(str(stan))
		self.vat_entry.insert(str(Model.get_product_vat(nazwa)))


	def _update(self):
		""" Aktualizauje informacje o produkcie """
		Model.update_product(self._id, self.nazwa_entry.get(), self.typ_entry.get(), float(self.cena_entry.get()), distutils.util.strtobool(self.stan_entry.get()), float(self.vat_entry.get()))
		self.controller.reload_treeview()
		self.destroy()


class ExcelProduct:
	""" 
	Reprezentuje produkt wyciągnięty z arkusza excela
	"""

	def __init__(self, skoroszyt, wiersz):
		""" Wyciąga dane z wiersza i sprawdza ich poprawnośc """

		self.nazwa = skoroszyt[f"B{2 + wiersz}"].value
		self.typ = skoroszyt[f"C{2 + wiersz}"].value
		self.cena = skoroszyt[f"D{2 + wiersz}"].value
		self.dostepnosc = skoroszyt[f"E{2 + wiersz}"].value
		self.vat = skoroszyt[f"F{2 + wiersz}"].value

		if not self.typ: self.typ = "brak"


	def is_invalid(self):
		""" Sprawdza czy wszystkie pobrane dane są poprawne """

		try:
			self.nazwa = str(self.nazwa)

		except ValueError:
			return f"Błąd: Wprowadzone dane nie są poprawne(nazwa w produkcie {self.nazwa} nie jest tekstem)."

		try:
			self.cena = float(self.cena)

		except ValueError:
			return f"Błąd: Wprowadzone dane nie są poprawne(cena w produkcie {self.nazwa} nie jest liczbą)."

		except TypeError:
			if isinstance(self.cena, datetime.datetime):
				return f"Błąd: Wprowadzone dane nie są poprawne(cena w produkcie {self.nazwa} została podana jako data. Należy zmienić ją na prawidłową liczbę)."

		try:
			self.vat = float(self.vat)

		except ValueError:
			return f"Błąd: Wprowadzone dane nie są poprawne(stawka vat w produkcie {self.nazwa} nie jest liczbą)."

		except TypeError:
			if isinstance(self.cena, datetime.datetime):
				return f"Błąd: Wprowadzone dane nie są poprawne(cena w produkcie {self.nazwa} została podana jako data. Należy zmienić ją na prawidłową liczbę)."

		try:
			self.dostepnosc = bool(self.dostepnosc)

		except ValueError:
			return f"Błąd: Wprowadzone dane nie są poprawne(stan produktu {self.nazwa} nie może być rozpoznany jako wartość prawda/fałsz)."


		if self.cena < 0:
			return f"Błąd: Wprowadzone dane nie są poprawne(cena w produkcie {self.nazwa} jest ujemna)."

		elif self.vat < 0:
			return f"Błąd: Wprowadzone dane nie są poprawne(stawka vat w produkcie {self.nazwa} jest ujemna)."


		return 0


class Appender(tk.Toplevel):
	""" 
	Klasa odpowiedzialna za wprowadzanie danych z arkusza .xlsx 
	"""

	def __init__(self, controller):
		super().__init__(controller.view)
		self.geometry("500x400")
		self.title("Zmień ustawienia")
		self.group(controller.view)
		self.resizable(False, False)

		self.controller = controller
		self.perform_reload = False

		# Definicje Listenerów
		def tab_changed(e):
			self.focus()

		# Definicje stylów
		style = ttk.Style(self)
		style.configure('lefttab.TNotebook', tabposition='wn')

		# Definicje widgetów
		notebook = ttk.Notebook(self, style='lefttab.TNotebook')
		notebook.bind('<<NotebookTabChanged>>', tab_changed)

		# Panele
		excel_panel = tk.Frame(notebook)

		self.xlsx_path = piktk.PathEntry(excel_panel, caption="Scieżka do pliku")
		self.wiersz_entry = piktk.LabeledEntry(excel_panel, caption="Ilość wierszy")

		xlsx_caption = tk.Label(excel_panel, text="Importuj dane z arkusza")

		self.excel_progress_panel = tk.Frame(excel_panel)

		# Panel zatwierdzenia
		panel_zatwierdzania = tk.LabelFrame(self)

		submit_button = tk.Button(panel_zatwierdzania, text="Zatwierdź", command=self.importuj_arkusz)
		cancel_button = tk.Button(panel_zatwierdzania, text="Anuluj", command=self.destroy)

		submit_button.pack(side=tk.RIGHT, padx=4, pady=4)
		cancel_button.pack(side=tk.RIGHT, padx=4, pady=4)

		panel_zatwierdzania.pack(side=tk.BOTTOM, fill=tk.X, padx=4, pady=4)
		

		db_panel = tk.Frame(notebook)

		# Dodawanie kart do notebooka
		xlsx_caption.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		piktk.Separator(excel_panel).pack(fill=tk.X, padx=4, pady=4)
		self.xlsx_path.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.wiersz_entry.pack(side=tk.TOP, anchor=tk.W, padx=4, pady=4)
		self.excel_progress_panel.pack(side=tk.TOP, fill=tk.X, pady=(16, 0), padx=4)
		#self.set_progressbar(35)
		excel_panel.pack(fill=tk.BOTH, expand=1)

		notebook.add(excel_panel, text="          .xlsx")
		notebook.add(db_panel, text="           .db")
		notebook.tab(1, state="disabled")

		notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
		self.mainloop()


	def reload_treeview(self):
		""" Reloads the product view """

		self.controller.clear_treeview()
		self.controller.load_treeview()
		showinfo(title="Suckes", message="Operacja została zakończona powodzeniem.", parent = self)
		self.destroy()
		

	def importuj_arkusz(self):
		try:
			plik = self.xlsx_path.get()
			wiersze = int(self.wiersz_entry.get())

		except ValueError:
			showerror(title="Błąd", message="Błąd: Wprowadzone dane nie są poprawne.", parent = self)
			return

		# Wczytywanie akrusza
		try:
			arkusz = openpyxl.load_workbook(filename=plik, data_only=True)
			skoroszyt = arkusz.active

		except Exception:
			showerror(title="Błąd", message="Błąd: Podany akrusz nie istnieje.", parent = self)
			return

		# Dodawanie produktów 
		if skoroszyt[f"B2"].value == None:
			showerror(title="Błąd", message="Błąd: Podany arkusz jest nieprawidłowy.", parent = self)
			return

		for wiersz in range(wiersze):
			if skoroszyt[f"B{2 + wiersz}"].value == None:
				showerror(title="Błąd", message="Błąd: Akrusz nie posiada podanej liczb wierszy informacji.", parent = self)
				return

		if wiersze > 25:
			self.set_progressbar(wiersze)
		
		
		import threading
		thr = threading.Thread(target=self.load_products, args=(skoroszyt, wiersze))
		thr.start()


	def set_progressbar(self, wiersze):
		self.progress = ttk.Progressbar(self.excel_progress_panel)
		self.status_text = tk.Label(self.excel_progress_panel, text="Status...")

		self.progress.pack(fill=tk.X, pady=4, padx=4)
		self.status_text.pack(side=tk.TOP, padx=4, pady=4, anchor=tk.W)


	def truncate_name(self, string, limit):
		""" Przycina podanego stringa jeżeli jest dłuższy niż podany limit. """
		return (string[:(limit - 2)] + '..') if len(string) > (limit - 2) else string


	def load_products(self, skoroszyt, wiersze):
		step = 100 / wiersze

		for wiersz in range(wiersze):
			new_product = ExcelProduct(skoroszyt, wiersz)

			if error := new_product.is_invalid():
				showerror(title="Błąd", message=error, parent = self)
				return

			self.status_text.config(text = f"Importing {self.truncate_name(new_product.nazwa, 50)}...")
			Model.add_product(
				new_product.nazwa, 
				new_product.typ, 
				new_product.cena, 
				new_product.dostepnosc, 
				new_product.vat
			)
			self.progress.step(step)

		self.progress['value'] = 100
		self.status_text.config(text = f"Done.")

		self.after(0, self.reload_treeview)


@dataclass
class ProductData:
	""" 
	Klasa reprezentująca dane konkretnego produktu na fakturze.
	"""

	# Nazwa produktu
	nazwa: str

	# Kod produktu
	kod: str

	# Zamawiana ilość sztuk
	ilosc: int

	# Cena łączna za wszystkie zamówione sztuki produktu
	cena: Decimal

	# Stawka vat za jedną sztukę produktu
	vat: Decimal

	# Uwagi na temat produktu
	uwagi: str


class Controller:
	""" Klasa odpowiedzialna za kontrolę danych i interfejsu """


	def __init__(self, view):
		self.view = view


	def start_appender(self):
		""" Metoda uruchamiająca program imprortu z arkusza xlsx """
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
		ustawienia = Model.get_settings()

		# Ustawianie ustawień właściwe
		self.sprzedawca = ustawienia["sprzedawca"]
		self.firma_sprzedawcy = ustawienia["firma-sprzedawcy"]
		self.email = ustawienia["email"]
		self.kod_pocztowy = ustawienia["kod-pocztowy"]
		self.adres = ustawienia["adres"]
		self.uwzglednij_dostawe = distutils.util.strtobool(ustawienia["uwzglednij-dostawe"])
		self.dostawa = ustawienia["dostawa"]


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


	def get_os_supported_path(self, path: str):
		""" Zwraca scieżkę odpowiednią dla systemu operacyjnego """
		if platform.system() == "Windows":
			return path.replace("/", "\\")

		else:
			return path.replace("\\", "/")


	def create_invoice(self):
		""" Tworzy plik .xlsx z fakturą """
		rozszerzenia = (('Dokument PDF', '*.pdf'), ('Skoroszyt programu Excel', '*.xlsx'))
		nazwa = asksaveasfilename(filetypes = rozszerzenia, defaultextension=".pdf")
		nazwa = self.get_os_supported_path(nazwa)

		# Uzyskiwanie danych o produktach na fakturze
		self.load_settings()
		produkty = []

		nazwy_cech = ["nazwa", "ilosc", "cena", "uwagi"]
		for produkt in self.view.kupione_produkty.get_children():

			# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
			cechy_produktu = {"kod": str(self.view.kupione_produkty.item(produkt)["text"])}
			for nr_cechy, cecha in enumerate(self.view.kupione_produkty.item(produkt)['values']):
				cechy_produktu[nazwy_cech[nr_cechy]] = cecha

			# Przypisywanie atrybutów
			produkty.append(
				ProductData(
					cechy_produktu["nazwa"], 
					cechy_produktu["kod"], 
					int(cechy_produktu["ilosc"]), 
					Decimal(cechy_produktu["cena"]), 
					Decimal(Model.get_product_vat(cechy_produktu["nazwa"])),
					cechy_produktu["uwagi"]
				)
			)

		# Doliczanie dostawy
		if self.uwzglednij_dostawe:

			produkty.append(
				ProductData(
					"Dostawa", 
					0, 
					1, 
					Decimal(self.dostawa), 
					Decimal(0),
					""
				)
			)

		info = {
			"sprzedawca": self.view.sprzedawca_entry.get(),
			"firma_sprzedawcy": self.view.firma_sprzedawcy_entry.get(),
			"email_sprzedawcy": self.view.email_sprzedawcy_entry.get(),
			"adres_sprzedawcy": self.view.adres_sprzedawcy_entry.get(),
			"ulica_sprzedawcy": self.view.ulica_sprzedawcy_entry.get(),
			"nabywca": self.view.nabywca_entry.get(),
			"nabywca_kod_pocztowy": self.view.nabywca_kod_pocztowy_entry.get(),
			"nabywca_nr_tel": self.view.nabywca_nr_tel_entry.get(),
			"nabywca_email": self.view.nabywca_email_entry.get(),
			"typ_dostawy": self.view.typ_dostawy_entry.get(),
			"forma_platnosci": self.view.forma_platnosci_entry.get(),
			"adres": self.view.adres_entry.get(),
			"kupione_produkty": produkty,
		}

		plik = Faktura(nazwa, info, bool(self.view.include_vat.get()))
		plik.build()
			