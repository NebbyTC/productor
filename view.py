import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import platform

import components as pikttk


class WindowsWindow(tk.Tk):
	""" Klasa okna na systemie windows """


	def __init__(self):
		super().__init__()
		self.iconbitmap(default="resources/icon.ico")


class LinuxWindow(tk.Tk):
	""" Klasa okna na systemie Linux """


	def __init__(self):
		super().__init__(className='Software')
		ikona = tk.PhotoImage(file='resources/icon.gif')
		self.iconphoto(True, ikona)


OKNA = {
	"Linux": LinuxWindow,
	"Windows": WindowsWindow
}

KOLOR_TLA  = {
	"Linux": "#DDDDDD",
	"Darwin": "#DDDDDD",
	"": "#DDDDDD",
	"Windows": "SystemButtonFace"
}

OS = platform.system()
OKNO = OKNA[OS]


class About(tk.Toplevel):
	""" Klasa okna "O programie" """


	def __init__(self, master):
		super().__init__(master)
		self.geometry("300x300")
		self.title("O programie")
		self.resizable(False, False)
		self.group(master)

		# Definicje widgetów
		title = tk.Label(self, text="Informacje o programie\nPik Productor")

		trademark = tk.Label(self, text="© Pik Software Development S.A. 2015-2022")
		wersja = tk.Label(self, text="Wersja 1.0.1 rc-2")

		produkcja = tk.Text(self, relief=tk.FLAT, wrap=tk.WORD, font="TkDefaultFont", bg=KOLOR_TLA[OS], height=3)
		produkcja.insert(1.0, "Wyprodukowano przez Pik Software Development S.A. dla M.B. & Spółka")
		produkcja["state"] = tk.DISABLED

		licencja = tk.Text(self, relief=tk.FLAT, wrap=tk.WORD, font="TkDefaultFont", bg=KOLOR_TLA[OS])
		licencja.insert(1.0, "Licencję na ten produkt zgodnie z postanowieniami licencyjnymi dotyczącymi oprogramowania firmy Pik Software Development posiada:\n	M.B & Spółka")
		licencja["state"] = tk.DISABLED

		# Pakowanie widżetów
		title.pack(side=tk.TOP, fill=tk.X)

		pikttk.Separator(self).pack(fill=tk.X, padx=10, pady=5)

		trademark.pack(side=tk.TOP, anchor=tk.W, padx=4)
		wersja.pack(side=tk.TOP, anchor=tk.W, padx=4)

		pikttk.Separator(self).pack(fill=tk.X, padx=10, pady=5)

		produkcja.pack(fill=tk.X, padx=8)
		licencja.pack(fill=tk.X, padx=8)


class Gui(OKNO):
	""" Klasa reprezentująca interfejs graficzny programu """


	def __init__(self):
		super().__init__()
		self.geometry("1080x700")
		self.title("Productor")
		self.minsize(900, 680)

		# Definiowanie menu górnego
		self.menubar = tk.Menu(self)

		# Operacje Menu
		self.operacja_menu = tk.Menu(tearoff=False)
		self.operacja_menu.add_command(label="Dodaj nowy produkt", command=lambda: self.mennubar_focus(self.nazwa_entry))
		self.operacja_menu.add_command(label="Usuń produkt", command=lambda: self.mennubar_focus(self.deletion_by_id_entry))
		self.operacja_menu.add_command(label="Wyszukaj produkt", command=lambda: self.mennubar_focus(self.se))
		self.operacja_menu.add_separator()
		self.operacja_menu.add_command(label="Wyczyśc wyszukiwarkę", command=lambda: self.clear_search())
		self.menubar.add_cascade(label="Operuj", menu=self.operacja_menu)

		# Import menu
		self.import_menu = tk.Menu(tearoff=False)
		self.import_menu.add_command(label="Importuj arkusz excela", command=lambda: self.controller.start_appender())
		self.menubar.add_cascade(label="Importuj", menu=self.import_menu)

		# Export menu
		self.export_menu = tk.Menu(tearoff=False)
		self.export_menu.add_command(label="Exprtuj produkty do arkuszu", command=lambda: self.controller.start_exporter())
		self.menubar.add_cascade(label="Eksportuj", menu=self.export_menu)

		# Ustawienia menu
		self.menubar.add_command(label="Ustawienia", command=self.controller.change_settings)

		# Info menu
		self.info_menu = tk.Menu(tearoff=False)
		self.info_menu.add_command(label="Dokumentacja", command="#")
		self.info_menu.add_command(label="O programie", command=lambda: About(self))
		self.menubar.add_cascade(label="Informacje", menu=self.info_menu)

		# Konfiguracja
		self.config(menu=self.menubar)


		# Definiowanie nadrzędnego notebooka
		self.toplevel = ttk.Notebook(self)
		self.toplevel.pack(fill=tk.BOTH, expand=1)

		# Definiowanie podstawowych paneli
		self.product_database_panel = tk.Frame(self)
		self.invoice_panel = tk.Frame(self)

		db_img = tk.PhotoImage(file="resources/db_icon.png")
		self.toplevel.add(self.product_database_panel, text="Produkty", image=db_img, compound=tk.LEFT)

		paper_img = tk.PhotoImage(file="resources/faktura_icon.png")
		self.toplevel.add(self.invoice_panel, text="Faktura", image=paper_img, compound=tk.LEFT)


		""" Panel bazy danych """
		self.operation_pane = tk.Frame(self.product_database_panel, width=45)
		self.view_pane = tk.Frame(self.product_database_panel, relief=tk.SUNKEN, bd=4)

		self.appending_pane = tk.Frame(self.operation_pane, relief=tk.SUNKEN, bd=4, height=225)
		self.deleting_pane = tk.Frame(self.operation_pane, relief=tk.SUNKEN, bd=4, height=100)
		self.searching_pane = tk.Frame(self.operation_pane, relief=tk.SUNKEN, bd=4, height=100)

		# Panel dodawania
		self.internal_appending_pane = tk.Frame(self.appending_pane)

		self.input_welcoming_label = tk.Label(self.internal_appending_pane, text="Wprowadź dane do bazy danych")

		self.nazwa_label = tk.Label(self.internal_appending_pane, text="Nazwa:")
		self.nazwa_entry = tk.Entry(self.internal_appending_pane, width=33)

		self.typ_label = tk.Label(self.internal_appending_pane, text="Typ:")
		self.typ_entry = tk.Entry(self.internal_appending_pane, width=33)

		self.cena_label = tk.Label(self.internal_appending_pane, text="Cena:")
		self.cena_entry = tk.Entry(self.internal_appending_pane, width=33)
		
		self.stan_label = tk.Label(self.internal_appending_pane, text="Stan:")
		self.stan_entry = ttk.Combobox(self.internal_appending_pane, width=30)
		self.stan_entry["values"] = ("True", "False")
		self.stan_entry.current(0)
		self.stan_entry.bind("<<ComboboxSelected>>", lambda e: self.focus())

		self.input_submit = tk.Button(self.internal_appending_pane, text="Prześlij", width=10)
		self.input_submit.config(command=lambda: self.controller.append_to_treeview())

		# Panel usuwania
		self.internal_deleting_pane = tk.Frame(self.deleting_pane)

		self.deleteion_welcoming_label = tk.Label(self.internal_deleting_pane, text="Usuń dane z bazy danych")

		self.deletion_by_id_label = tk.Label(self.internal_deleting_pane, text="Id produktu:  ")
		self.deletion_by_id_entry = tk.Entry(self.internal_deleting_pane, width=25)

		self.deletion_submit = tk.Button(self.internal_deleting_pane, text="Usuń", command=lambda: self.del_from_db("odczynnik", self.e4))
		self.deletion_submit.config(command=lambda: self.controller.del_from_treeview())

		# Panel wyszukiwania
		self.search_welcoming_label = tk.Label(self.searching_pane, text="Wyszukaj")

		self.se = tk.Entry(self.searching_pane, width=49)
		self.se.bind('<KeyRelease>', self.__search_listener)
		
		self.result = tk.Label(self.searching_pane, text="\n")
				
		# Widok bazy danych
		rozmiar_kolumny, rozmiar_id = 100, 75

		self.widok = ttk.Treeview(self.view_pane, height=9, selectmode="none")
		self.widok["columns"] = ("one","two","three", "four")

		self.widok.column('#0', width=rozmiar_id)
		self.widok.column('one', width=rozmiar_kolumny)
		self.widok.column('two', width=rozmiar_kolumny)
		self.widok.column('three', width=rozmiar_kolumny)
		self.widok.column('four', width=rozmiar_kolumny)

		self.widok.heading("#0",text="Id")
		self.widok.heading("one", text="Nazwa")
		self.widok.heading("two", text="Typ")
		self.widok.heading("three", text="Cena")
		self.widok.heading("four", text="Stan")

		self.scrollbar = tk.Scrollbar(self.view_pane, width=20)
		self.scrollbar.config(command=self.widok.yview)

		self.widok_wyszukiwanego = ttk.Treeview(self.view_pane, height=5, selectmode="extended")
		self.widok_wyszukiwanego["columns"]=("one","two","three", "four")

		self.widok_wyszukiwanego.column('#0', width=rozmiar_id)
		self.widok_wyszukiwanego.column('one', width=rozmiar_kolumny)
		self.widok_wyszukiwanego.column('two', width=rozmiar_kolumny)
		self.widok_wyszukiwanego.column('three', width=rozmiar_kolumny)
		self.widok_wyszukiwanego.column('four', width=rozmiar_kolumny)

		self.widok_wyszukiwanego.heading("#0",text="Id")
		self.widok_wyszukiwanego.heading("one", text="Nazwa")
		self.widok_wyszukiwanego.heading("two", text="Typ")
		self.widok_wyszukiwanego.heading("three", text="Cena")
		self.widok_wyszukiwanego.heading('four', text="Stan")

		self.widok_wyszukiwanego.bind("<Button-3>", lambda e: self.controller.edit_product())


		# Pakowanie widżetów
		self.operation_pane.pack(fill=tk.Y, side=tk.LEFT)
		self.view_pane.pack(fill=tk.BOTH, expand=1, padx=(4, 8), pady=8)

		self.appending_pane.pack(padx=(8, 4), pady=(8, 4), fill=tk.X)
		self.deleting_pane.pack(padx=(8, 4), pady=4, fill=tk.X)
		self.searching_pane.pack(padx=(8, 4), pady=(4, 8), fill=tk.X)

		# Pakowanie panelu wprowadzania
		self.internal_appending_pane.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		self.input_welcoming_label.grid(row=0, column=0, columnspan=2, padx=4, pady=4)

		self.nazwa_label.grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
		self.nazwa_entry.grid(row=1, column=1, padx=4, pady=4)		

		self.typ_label.grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
		self.typ_entry.grid(row=2, column=1, padx=4, pady=4)

		self.cena_label.grid(row=3, column=0, sticky=tk.W, padx=4, pady=4)
		self.cena_entry.grid(row=3, column=1, padx=4, pady=4)

		self.stan_label.grid(row=4, column=0, sticky=tk.W, padx=4, pady=4)
		self.stan_entry.grid(row=4, column=1, padx=4, pady=4)

		self.input_submit.grid(row=5, column=1, padx=4, pady=4, sticky=tk.E)
		
		# Pakowanie panelu usuwania
		self.internal_deleting_pane.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		self.deleteion_welcoming_label.grid(row=0, column=0, columnspan=4, padx=4, pady=4)

		self.deletion_by_id_label.grid(row=1, column=0, padx=4, pady=4, sticky=tk.W)
		self.deletion_by_id_entry.grid(row=1, column=1, padx=4, pady=4)

		self.deletion_submit.grid(row=1, column=2, padx=4)

		# Pakowanie panelu wyszukiwania
		self.search_welcoming_label.grid(row=0, column=0, columnspan=3, padx=10, pady=4)

		self.se.grid(row=1, column=1, columnspan=2, padx=6, pady=4)

		self.result.grid(row=3, column=0, columnspan=3, rowspan=2)

		# Pakowanie widoku bazy danych
		self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
		self.widok.pack(fill=tk.BOTH, expand=1)
		self.widok_wyszukiwanego.pack(fill=tk.BOTH, side=tk.BOTTOM)

		# Załadowanie danych z bazy
		self.controller.load_treeview()


		""" Panel faktury """
		panel_gorny = tk.Frame(self.invoice_panel)

		panel_produktow = tk.Frame(panel_gorny, relief=tk.SUNKEN, borderwidth=4)

		panel_informacji = tk.Frame(panel_gorny)

		panel_sprzedawcy = tk.LabelFrame(panel_informacji, relief=tk.GROOVE, borderwidth=4, text="Informacje o sprzedawcy")

		panel_nabywcy = tk.LabelFrame(panel_informacji, relief=tk.GROOVE, borderwidth=4, text="Informacje o nabywcy")

		panel_odbiorcy = tk.LabelFrame(panel_informacji, relief=tk.GROOVE, borderwidth=4, text="Informacje o odbiorcy")

		panel_vat = tk.LabelFrame(panel_informacji, relief=tk.GROOVE, borderwidth=4, text="Informacje o podatku VAT")

		panel_kontroli = tk.LabelFrame(self.invoice_panel, relief=tk.SUNKEN, borderwidth=4, height=100)

		# Panel sprzedawcy
		self.controller.load_settings()

		sprzedawca_label = tk.Label(panel_sprzedawcy, text="Sprzedawca", width=15, anchor=tk.W)
		self.sprzedawca_entry = tk.Entry(panel_sprzedawcy, width=30)
		self.sprzedawca_entry.insert(tk.END, self.controller.sprzedawca)

		firma_sprzedawcy_label = tk.Label(panel_sprzedawcy, text="Firma sprzedawcy", width=15, anchor=tk.W)
		self.firma_sprzedawcy_entry = tk.Entry(panel_sprzedawcy, width=30)
		self.firma_sprzedawcy_entry.insert(tk.END, self.controller.firma_sprzedawcy)

		email_sprzedawcy_label = tk.Label(panel_sprzedawcy, text="Email sprzedawcy", width=15, anchor=tk.W)
		self.email_sprzedawcy_entry = tk.Entry(panel_sprzedawcy, width=30)
		self.email_sprzedawcy_entry.insert(tk.END, self.controller.email)

		adres_sprzedawcy_label = tk.Label(panel_sprzedawcy, text="Adres sprzedawcy", width=15, anchor=tk.W)
		self.adres_sprzedawcy_entry = tk.Entry(panel_sprzedawcy, width=30)
		self.adres_sprzedawcy_entry.insert(tk.END, self.controller.kod_pocztowy)

		ulica_sprzedawcy_label = tk.Label(panel_sprzedawcy, text="Ulica sprzedawcy", width=15, anchor=tk.W)
		self.ulica_sprzedawcy_entry = tk.Entry(panel_sprzedawcy, width=30)
		self.ulica_sprzedawcy_entry.insert(tk.END, self.controller.adres)

		# Panel nabywcy
		nabywca_label = tk.Label(panel_nabywcy, text="Nabywca", width=15, anchor=tk.W)
		self.nabywca_entry = tk.Entry(panel_nabywcy, width=30)

		nabywca_kod_pocztowy_label = tk.Label(panel_nabywcy, text="Kod pocztowy", width=15, anchor=tk.W)
		self.nabywca_kod_pocztowy_entry = tk.Entry(panel_nabywcy, width=30)

		nabywca_nr_tel_label = tk.Label(panel_nabywcy, text="nr. Tel", width=15, anchor=tk.W)
		self.nabywca_nr_tel_entry = tk.Entry(panel_nabywcy, width=30)

		nabywca_email_label = tk.Label(panel_nabywcy, text="email", width=15, anchor=tk.W)
		self.nabywca_email_entry = tk.Entry(panel_nabywcy, width=30)

		# Panel odbiorcy
		typ_dostawy_label = tk.Label(panel_odbiorcy, text="Typ dostawy", width=15, anchor=tk.W)
		self.typ_dostawy_entry = tk.Entry(panel_odbiorcy, width=30)

		forma_platnosci_label = tk.Label(panel_odbiorcy, text="Forma płatności", width=15, anchor=tk.W)
		self.forma_platnosci_entry = tk.Entry(panel_odbiorcy, width=30)

		adres_label = tk.Label(panel_odbiorcy, text="Adres", width=15, anchor=tk.W)
		self.adres_entry = tk.Entry(panel_odbiorcy, width=30)

		# Panel załączania podatku VAT
		self.include_vat = tk.IntVar()

		include_vat_button = tk.Checkbutton(panel_vat, text = "Uwzględnij podatek VAT", 
			variable = self.include_vat, onvalue = 1, offvalue = 0)

		# Panel kupowania produktów
		self.kupione_produkty = ttk.Treeview(panel_produktow, selectmode="none")
		self.kupione_produkty["columns"]=("one","two","three", "four")
		self.kupione_produkty.column('#0', width=50)
		self.kupione_produkty.column('one', width=100)
		self.kupione_produkty.column('two', width=100)
		self.kupione_produkty.column('three', width=100)
		self.kupione_produkty.column('four', width=100)
		self.kupione_produkty.heading("#0",text="Kod")
		self.kupione_produkty.heading("one", text="Nazwa towaru")
		self.kupione_produkty.heading("two", text="Ilośc")
		self.kupione_produkty.heading("three", text="Cena")
		self.kupione_produkty.heading("four", text="Uwagi")

		# Panel wprowadzania produktów
		internal_panel_kontroli = tk.Frame(panel_kontroli)

		prod_label = tk.Label(internal_panel_kontroli, text="Produkt")

		prod_label = tk.Label(internal_panel_kontroli, text="Produkt")
		self.prod_combo = ttk.Combobox(internal_panel_kontroli)
		self.prod_combo['values'] = self.controller.get_product_names()
		self.prod_combo.bind("<<ComboboxSelected>>", lambda e: self.focus())

		ilosc_label = tk.Label(internal_panel_kontroli, text="Ilość")
		ilosc_spin = ttk.Spinbox(
		    internal_panel_kontroli,
			from_=0,
			to=99,
			wrap=True
		)

		uwagi_label = tk.Label(internal_panel_kontroli, text="Uwagi")
		uwagi_entry = tk.Entry(internal_panel_kontroli, width=30)

		cart_submit = tk.Button(internal_panel_kontroli, text="Dodaj", 
			command=lambda: self.controller.add_to_cart(
					self.prod_combo.get(),
					int(ilosc_spin.get()),
					uwagi_entry.get()
				)
			)

		invoice_submit = tk.Button(internal_panel_kontroli, text="Eksportuj fakturę")
		invoice_submit.config(command=self.controller.create_invoice)


		# Pakowanie
		panel_gorny.pack(side=tk.TOP, fill=tk.BOTH, expand=1)


		panel_informacji.pack(side=tk.LEFT, pady=4, padx=(4, 0), anchor=tk.N)


		panel_sprzedawcy.pack(side=tk.TOP, fill=tk.X)

		sprzedawca_label.grid(row=0, column=0, padx=8, pady=8, sticky=tk.W)
		self.sprzedawca_entry.grid(row=0, column=1, padx=8, pady=8)

		firma_sprzedawcy_label.grid(row=1, column=0, padx=8, pady=8, sticky=tk.W)
		self.firma_sprzedawcy_entry.grid(row=1, column=1, padx=8, pady=8)

		email_sprzedawcy_label.grid(row=2, column=0, padx=8, pady=8, sticky=tk.W)
		self.email_sprzedawcy_entry.grid(row=2, column=1, padx=8, pady=8)

		adres_sprzedawcy_label.grid(row=3, column=0, padx=8, pady=8, sticky=tk.W)
		self.adres_sprzedawcy_entry.grid(row=3, column=1, padx=8, pady=8)

		ulica_sprzedawcy_label.grid(row=4, column=0, padx=8, pady=8, sticky=tk.W)
		self.ulica_sprzedawcy_entry.grid(row=4, column=1, padx=8, pady=8)


		panel_nabywcy.pack(side=tk.TOP, fill=tk.X)

		nabywca_label.grid(row=0, column=0, padx=8, pady=8, sticky=tk.W)
		self.nabywca_entry.grid(row=0, column=1, padx=8, pady=8)

		nabywca_kod_pocztowy_label.grid(row=1, column=0, padx=8, pady=8, sticky=tk.W)
		self.nabywca_kod_pocztowy_entry.grid(row=1, column=1, padx=8, pady=8)

		nabywca_nr_tel_label.grid(row=2, column=0, padx=8, pady=8, sticky=tk.W)
		self.nabywca_nr_tel_entry.grid(row=2, column=1, padx=8, pady=8)

		nabywca_email_label.grid(row=3, column=0, padx=8, pady=8, sticky=tk.W)
		self.nabywca_email_entry.grid(row=3, column=1, padx=8, pady=8)


		panel_odbiorcy.pack(side=tk.TOP, fill=tk.X)

		typ_dostawy_label.grid(row=0, column=0, padx=8, pady=8, sticky=tk.W)
		self.typ_dostawy_entry.grid(row=0, column=1, padx=8, pady=8)

		forma_platnosci_label.grid(row=1, column=0, padx=8, pady=8, sticky=tk.W)
		self.forma_platnosci_entry.grid(row=1, column=1, padx=8, pady=8)

		adres_label.grid(row=2, column=0, padx=8, pady=8, sticky=tk.W)
		self.adres_entry.grid(row=2, column=1, padx=8, pady=8)


		panel_vat.pack(side=tk.TOP, fill=tk.X)

		include_vat_button.grid(row=0, column=0, padx=8, pady=8)


		panel_kontroli.pack(side=tk.TOP, fill=tk.BOTH, padx=4, pady=(0, 4))
		internal_panel_kontroli.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

		prod_label.grid(row=0, column=0, columnspan=1)
		self.prod_combo.grid(row=1, column=0, padx=4)

		ilosc_spin.grid(row=1, column=1)
		ilosc_label.grid(row=0, column=1)

		uwagi_label.grid(row=0, column=2)
		uwagi_entry.grid(row=1, column=2, padx=4)

		cart_submit.grid(row=1, column=3)

		invoice_submit.grid(row=1, column=4, padx=(8, 0))

		panel_produktow.pack(side=tk.LEFT, fill=tk.BOTH, expand=1, padx=4, pady=(8, 4))
		self.kupione_produkty.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

		# Pętla
		self.mainloop()


	def clear_search(self):
		""" Czyści dolny widok i widgety powiązane z wyszukiwaniem """

		self.widok_wyszukiwanego.delete(*self.widok_wyszukiwanego.get_children())
		self.se.delete(0, 'end')

		for child in self.widok.get_children():
			self.widok.focus(child)
			self.widok.selection_remove(child)
		self.result.config(text="\nZnaleziono produktów: 0")
		self.result.focus()


	def __search_listener(self, event=None):
		""" Odpowiada za mechanizm dynamicznego wyszukiwania """

		# Odznaczanie wszystkiego
		for child in self.widok.get_children():
			self.widok.focus(child)
			self.widok.selection_remove(child)

		# Usuwanie wszystkiego z dolnego panelu
		for child in self.widok_wyszukiwanego.get_children():
			self.widok_wyszukiwanego.delete(child)

		# Dodaj
		licznik = 0
		for child in self.widok.get_children():
			if self.se.get() in self.widok.item(child)["values"][0] and self.se.get() != "":
				# Górny panel
				self.widok.focus(child)
				self.widok.selection_add(child)
				licznik += 1

				# Dolny Panel
				self.widok_wyszukiwanego.insert('', 'end', text=self.widok.item(child)["text"], values=(self.widok.item(child)["values"]))
		self.result.config(text="\nZnaleziono produktów: "+str(licznik))


	def reload_bottom_treeview(self):
		self.widok_wyszukiwanego.delete(*self.widok_wyszukiwanego.get_children())
		self.__search_listener()


	def mennubar_focus(self, widget):
		""" Daje fokus na wybrany widget i zmienia zakładkę notebooka """
		widget.focus()
		self.toplevel.select(0)


	def set_controller(self, controller_obj):
		""" Ustawia kontroler obiektu """
		self.controller = controller_obj


if __name__ == '__main__':
	Gui()