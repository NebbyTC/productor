import os
import datetime
import platform

import openpyxl
from openpyxl.styles import Alignment, Border, Side
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from decimal import Decimal

from model import Model


class Faktura:
	""" Klasa reprezentująca fakturę """

	# Określa maksymalną liczbę znaków na kolumnę
	MAX_CHAR: int = 18


	def __init__(self, nazwa: str, informacje: dict, uwzglednij_vat: bool):
		"""
		Zapisuje informacje do zawarcia na fakturze.

		Args:
			nazwa: 
				Scieżka pliku, do którego faktura ma
				zostać zapisana.

			informacje:
				Słownik zawierający informacje potrzebne
				do wystawienia faktury.

			uwzglednij_vat:
				Wartość true/false na podstawie, której
				program decyduje czy uwzględnić vat.

		Raises:
			ValueError:
				Zwracany jeżeli, któryś z argumentów będzie 
				miał niepoprawny typ lub zawartość.

		"""

		# Sprawdzanie argumentu 'nazwa'
		if not isinstance(nazwa, str):
			raise ValueError("[Productor] Argument 'nazwa' nie jest typu 'str.") 

		if (".pdf" in nazwa) or (".xlsx" in nazwa):
			self.nazwa = nazwa
		else:
			raise ValueError("[Productor] Argument 'nazwa' nie jest poprawną scieżką do pliku faktury.") 

		# Sprawdzanie argumentu 'informacje'
		if not isinstance(informacje, dict):
			raise ValueError("[Productor] Argument 'informacje' nie jest typu 'dict.") 

		wymagane_pola = [
			"sprzedawca", 
			"firma_sprzedawcy",
			"email_sprzedawcy",
			"adres_sprzedawcy",
			"ulica_sprzedawcy",
			"nabywca",
			"nabywca_kod_pocztowy",
			"nabywca_nr_tel",
			"nabywca_email",
			"typ_dostawy",
			"forma_platnosci",
			"adres",
			"kupione_produkty",
			"uwzglednij_dostawe",
			"dostawa"
		]

		for pole in wymagane_pola:

			if not pole in informacje.keys():
				raise ValueError(f"[Productor] Argument 'informacje' nie zawiera wymaganego klucza '{pole}'.") 

		self.info = informacje

		# Sprawdzanie argumetu uwzglednij_vat
		if not isinstance(uwzglednij_vat, bool):
			raise ValueError("[Productor] Argument 'uwzglednij_vat' nie jest typu 'bool.")

		self.uwzglednij_vat = uwzglednij_vat

		# Uzyskiwanie pozostałych danych
		self.os = platform.system()

		# Przetwarzanie informacji o kupionych produktach
		"""
		for nr_produktu, produkt in enumerate(self.info["kupione_produkty"].get_children()):

			# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
			cechy_produktu = {"kod": str(self.info["kupione_produkty"].item(produkt)["text"])}
			for nr_cechy, cecha in enumerate(self.info["kupione_produkty"].item(produkt)['values']):
				cechy_produktu[nazwy_cech[nr_cechy]] = cecha

		print()
		"""



	def build(self):
		""" Generuje fakturę uwzględniając atrybuty obiektu """
		if self.uwzglednij_vat:
			self.template_file = "template_vat"
		else:
			self.template_file = "template"

		if ".pdf" in self.nazwa:
			self.build_pdf(include_vat=self.uwzglednij_vat)

		elif ".xlsx" in self.nazwa:
			self.build_xlsx(include_vat=self.uwzglednij_vat)


	def build_xlsx(self, include_vat=False):
		""" 
		Odpowiada za stworzenie faktury w formacie .xlsx 

		Args:
			include_vat: 
				wartość typu bool, określająca czy w fakturze
				ma zaostać uwzględniony podatek vat.
		"""

		if self.os == "Windows":
			os.system(f'copy resources\\{self.template_file}.xlsx {self.nazwa}')

		else:
			os.system(f'cp resources/{self.template_file}.xlsx {self.nazwa}')

		skoroszyt = openpyxl.load_workbook(filename=self.nazwa)
		arkusz = skoroszyt.active
		__class__.MAX_CHAR = 18

		# Wstawianie tabeli
		zakupy = self.info["kupione_produkty"]
		suma_brutto = Decimal(0.0)
		suma_netto = Decimal(0.0)
		suma_vat = Decimal(0.0)

		wiersz = 31


		def kolejny_wiersz(nazwa: str):
			""" 
			Przesuwa kursor(zmienną 'wiersz') o
			tyle, ile zajmuje podana nazwa.

			Args:
				nazwa:
					Długość nazwy produktu.

			Raises:
				ValueError: 
					Zaracany, jeżeli argument 'nazwa' 
					nie jest typu str.
			"""
			nonlocal wiersz

			dlugosc = len(nazwa)
			ilosc_czesci = dlugosc/__class__.MAX_CHAR

			if ilosc_czesci > int(ilosc_czesci):
				wiersz += int(ilosc_czesci)+1

			else:
				wiersz += int(ilosc_czesci)


		def wpisz_nazwe(wiersz: int, nazwa: str):
			""" 
			Wpisuję podaną nazwę produktu do danego
			wiersza w odpowiedni sposób.

			Args:
				wiersz:
					Wiersz, od którego funkcja
					zacznie wstawiać nazwę.
				nazwa:
					Nazwa do wpisania do faktury.

			Raises:
				ValueError:
					Zwracany, jeżeli typy danych argumetów 
					nie są poprawne.
			""" 
			nonlocal arkusz

			if not isinstance(nazwa, str):
				raise ValueError("[Productor] Argument 'nazwa' nie jest typu 'str.")

			# Podział nazwy na równe części
			slowa = nazwa.split()
			czesci = []

			nr_slowa = 0
			czesc = ""
			while True:
				if len(slowa) <= nr_slowa:
					czesci.append(czesc)
					break

				if len(czesc + " " + slowa[nr_slowa]) <= __class__.MAX_CHAR:
					czesc = (czesc + " " + slowa[nr_slowa]).strip()
					nr_slowa += 1

				else:
					czesci.append(czesc)
					czesc = ""

			potrzebne_wiersze = len(czesci)

			# Wpisywanie części nazwy do kolejnych wierszy
			for i in range(potrzebne_wiersze):
				arkusz[f"C{wiersz+i}"] = czesci[i]


		def build_with_vat():
			""" 
			Generuje fakturę uwzględniając podatek vat.
			"""
			nonlocal wiersz, suma_netto, suma_brutto, suma_vat
			nazwy_cech = ["nazwa", "ilosc", "cena", "uwagi"]
			wiersz += 1

			for nr_produktu, produkt in enumerate(zakupy.get_children()):

				# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
				cechy_produktu = {"kod": str(zakupy.item(produkt)["text"])}
				for nr_cechy, cecha in enumerate(zakupy.item(produkt)['values']):
					cechy_produktu[nazwy_cech[nr_cechy]] = cecha

				# Obliczanie stawek
				stawka_vat = Decimal(Model.get_product_vat(cechy_produktu["nazwa"]))
				kwota_brutto = Decimal(float(cechy_produktu["cena"]))
				kwota_vat = kwota_brutto*(stawka_vat/100)
				kwota_netto = kwota_brutto*(stawka_vat/100)+kwota_brutto

				# Wpisywanie do tabeli
				arkusz[f"A{wiersz}"] = nr_produktu
				arkusz[f"B{wiersz}"] = int(cechy_produktu["kod"])
				wpisz_nazwe(wiersz, cechy_produktu["nazwa"])
				arkusz[f"E{wiersz}"] = cechy_produktu["ilosc"]
				arkusz[f"F{wiersz}"] = "{0:.2f}".format(kwota_brutto)
				arkusz[f"G{wiersz}"] = "{0:.2f}".format(stawka_vat)
				arkusz[f"H{wiersz}"] = "{0:.2f}".format(kwota_vat)
				arkusz[f"I{wiersz}"] = "{0:.2f}".format(kwota_netto)

				for kolumna in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
					arkusz[f"{kolumna}{wiersz}"].alignment = Alignment(horizontal='left')
				
				# Sumowanie stawek
				suma_netto += kwota_brutto*(stawka_vat/100)+kwota_brutto
				suma_brutto += kwota_brutto
				suma_vat += kwota_brutto*(stawka_vat/100)
				
				kolejny_wiersz(cechy_produktu["nazwa"])

			
			arkusz[f"E{wiersz+1}"] = "Razem:"

			arkusz[f"F{wiersz+1}"] = "{0:.2f}".format(suma_brutto)

			arkusz[f"H{wiersz+1}"] = "{0:.2f}".format(suma_vat)
			arkusz[f"I{wiersz+1}"] = "{0:.2f}".format(suma_netto)


		def build_without_vat():
			""" 
			Generuje fakturę bez podatku vat.
			"""
			nonlocal wiersz, suma_netto, suma_brutto, suma_vat
			nazwy_cech = ["nazwa", "ilosc", "cena", "uwagi"]

			for nr_produktu, produkt in enumerate(zakupy.get_children()):

				# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
				cechy_produktu = {"kod": str(zakupy.item(produkt)["text"])}
				for nr_cechy, cecha in enumerate(zakupy.item(produkt)['values']):
					cechy_produktu[nazwy_cech[nr_cechy]] = cecha

				# Wpisywanie do tabeli
				arkusz[f"A{wiersz}"] = nr_produktu
				arkusz[f"B{wiersz}"] = int(cechy_produktu["kod"])
				wpisz_nazwe(wiersz, cechy_produktu["nazwa"])
				arkusz[f"E{wiersz}"] = cechy_produktu["ilosc"]
				arkusz[f"F{wiersz}"] = cechy_produktu["cena"]
				arkusz[f"G{wiersz}"] = cechy_produktu["uwagi"]

				for kolumna in ["A", "B", "C", "D", "E", "F", "G"]:
					arkusz[f"{kolumna}{wiersz}"].alignment = Alignment(horizontal='left')

				suma_brutto += Decimal(float(cechy_produktu["cena"]))
				kolejny_wiersz(cechy_produktu["nazwa"])

			arkusz[f"E{wiersz+1}"] = "Razem:"
			arkusz[f"F{wiersz+1}"] = "{0:.2f}".format(suma_brutto)


		# Wypisywanie produktów na fakturze
		if include_vat:
			build_with_vat()
		else:
			build_without_vat()

		# Miejsce na podpis
		thin = Side(border_style="thin", color="000000")

		arkusz[f"B{wiersz+4}"].border = Border(bottom=thin)
		arkusz[f"C{wiersz+4}"].border = Border(bottom=thin)
		arkusz[f"D{wiersz+4}"].border = Border(bottom=thin)
		arkusz[f"E{wiersz+4}"].border = Border(bottom=thin)

		arkusz.merge_cells(f"B{wiersz+5}:E{wiersz+5}")
		arkusz[f"B{wiersz+5}"].alignment = Alignment(horizontal='center')
		arkusz[f"B{wiersz+5}"] = "Podpis osoby wystawiającej fakturę"

		# Dane czasowe
		arkusz["G5"] = str(datetime.datetime.now()).split(" ")[0]
		arkusz["G6"] = str(datetime.datetime.now()).split(" ")[0]
		arkusz["G7"] = str(datetime.datetime.now()).split(" ")[0]

		# Dane sprzedawcy
		arkusz["A6"] = self.info["sprzedawca"]
		arkusz["A7"] = self.info["firma_sprzedawcy"]
		arkusz["A8"] = self.info["email_sprzedawcy"]
		arkusz["A9"] = self.info["adres_sprzedawcy"]
		arkusz["A10"] = self.info["ulica_sprzedawcy"]

		# Dane nabywcy
		arkusz["A13"] = self.info["nabywca"]
		arkusz["A14"] = self.info["nabywca_kod_pocztowy"]
		arkusz["A15"] = self.info["nabywca_nr_tel"]
		arkusz["A16"] = self.info["nabywca_email"]

		# Dane odbiorcy
		arkusz["A19"] = self.info["typ_dostawy"]
		arkusz["A20"] = self.info["adres"]
		arkusz["A21"] = self.info["nabywca_nr_tel"]
		arkusz["A22"] = self.info["nabywca_kod_pocztowy"]
		arkusz["A23"] = self.info["nabywca_email"]

		# Sposó dostawy
		arkusz["C25"] = self.info["typ_dostawy"]

		# Forma platnosc, kwota, waluta
		arkusz["A28"] = self.info["forma_platnosci"]

		if include_vat:
			arkusz["C28"] = "{0:.2f}".format(suma_netto)
		else:
			arkusz["C28"] = "{0:.2f}".format(suma_brutto)
		arkusz["D28"] = "PLN"

		skoroszyt.save(self.nazwa)
		print(f"[Productor] Pomyślnie zapisano fakturę w pliku {self.nazwa}")


	def build_pdf(self, include_vat=False):
		""" 
		Odpowiada za stworzenie faktury w formacie .pdf 
		
		Args:
			include_vat: 
				wartość typu bool, określająca czy w fakturze
				ma zaostać uwzględniony podatek vat.
		"""
		pdfmetrics.registerFont(TTFont('Calibri', 'resources/calibri.ttf'))

		template = PdfReader(f"resources/{self.template_file}.pdf", decompress=False).pages[0]
		template_obj = pagexobj(template)	

		canvas = Canvas(self.nazwa)		
		canvas.setFont("Calibri", 11)	

		xobj_name = makerl(canvas, template_obj)
		canvas.doForm(xobj_name)

		font_height = 12
		ystart = 706
		xstart = 53

		""" Wstawianie danych """

		# Data
		canvas.drawString(xstart+300, ystart+13, str(datetime.datetime.now()).split(" ")[0])
		canvas.drawString(xstart+300, ystart-2, str(datetime.datetime.now()).split(" ")[0])
		canvas.drawString(xstart+300, ystart-17, str(datetime.datetime.now()).split(" ")[0])

		# Dane sprzedawcy
		canvas.drawString(xstart, ystart, self.info["sprzedawca"])
		canvas.drawString(xstart, ystart-14, self.info["firma_sprzedawcy"])
		canvas.drawString(xstart, ystart-14*2, self.info["email_sprzedawcy"])
		canvas.drawString(xstart, ystart-14*3, self.info["adres_sprzedawcy"])
		canvas.drawString(xstart, ystart-14*4, self.info["ulica_sprzedawcy"])

		# Dane nabywcy
		canvas.drawString(xstart, ystart-101, self.info["nabywca"])
		canvas.drawString(xstart, ystart-101-14, self.info["nabywca_kod_pocztowy"])
		canvas.drawString(xstart, ystart-101-14*2, self.info["nabywca_nr_tel"])
		canvas.drawString(xstart, ystart-101-14*3, self.info["nabywca_email"])

		# Dane odbiorcy
		canvas.drawString(xstart, ystart-188, self.info["typ_dostawy"])
		canvas.drawString(xstart, ystart-188-14, self.info["adres"])
		canvas.drawString(xstart, ystart-188-14*2, self.info["nabywca_nr_tel"])
		canvas.drawString(xstart, ystart-188-14*3, self.info["nabywca_kod_pocztowy"])
		canvas.drawString(xstart, ystart-188-14*4, self.info["nabywca_email"])

		# Forma platnosc, kwota, waluta
		canvas.drawString(xstart+85, ystart-278, self.info["typ_dostawy"])

		canvas.drawString(xstart, ystart-325, self.info["forma_platnosci"])

		canvas.drawString(xstart+78*2, ystart-325, "PLN")

		# Wykaz produktów
		dostawa = Decimal(self.info["dostawa"])
		zakupy = self.info["kupione_produkty"]

		suma_brutto = Decimal(0.0)
		suma_netto = Decimal(0.0)
		suma_vat = Decimal(0.0)

		kolumny = {"A": xstart, "B": xstart+52, "C": xstart+52*2, "D": xstart+70*3-3, "E": xstart+52*5-3, "F": xstart+52*6-3, "G": xstart+52*7-3, "H": xstart+52*8-3}
		wiersz = 1

		def kolejny_wiersz(nazwa: str):
			""" 
			Przesuwa kursor(zmienną 'wiersz') o
			tyle, ile zajmuje podana nazwa.

			Args:
				nazwa:
					Długość nazwy produktu.

			Raises:
				ValueError: 
					Zaracany, jeżeli argument 'nazwa' 
					nie jest typu str.
			"""
			nonlocal wiersz

			dlugosc = len(nazwa)
			ilosc_czesci = dlugosc/__class__.MAX_CHAR

			if ilosc_czesci > int(ilosc_czesci):
				wiersz += int(ilosc_czesci)+1

			else:
				wiersz += int(ilosc_czesci)

		def wpisz_nazwe(wiersz: int, nazwa: str):
			""" 
			Wpisuję podaną nazwę produktu do danego
			wiersza w odpowiedni sposób.

			Args:
				wiersz:
					Wiersz, od którego funkcja
					zacznie wstawiać nazwę.
				nazwa:
					Nazwa do wpisania do faktury.

			Raises:
				ValueError:
					Zwracany, jeżeli typy danych argumetów 
					nie są poprawne.
			""" 
			nonlocal canvas

			if not isinstance(nazwa, str):
				raise ValueError("[Productor] Argument 'nazwa' nie jest typu 'str.")

			# Podział nazwy na równe części
			slowa = nazwa.split()
			czesci = []

			nr_slowa = 0
			czesc = ""
			while True:
				if len(slowa) <= nr_slowa:
					czesci.append(czesc)
					break

				if len(czesc + " " + slowa[nr_slowa]) <= __class__.MAX_CHAR:
					czesc = (czesc + " " + slowa[nr_slowa]).strip()
					nr_slowa += 1

				else:
					czesci.append(czesc)
					czesc = ""

			potrzebne_wiersze = len(czesci)

			# Wpisywanie części nazwy do kolejnych wierszy
			max_extent = 0
			for i in range(potrzebne_wiersze):
				canvas.drawString(kolumny["C"], ystart-365-14*(wiersz+i), czesci[i])
				max_extent = wiersz+i

			if include_vat:
				canvas.drawImage("resources/intermittent.png", kolumny["A"], ystart-365-14*max_extent-4, width=200, height=1)
				canvas.drawImage("resources/intermittent.png", xstart+52*3+3, ystart-365-14*max_extent-4, width=200, height=1)
				canvas.drawImage("resources/intermittent.png", xstart+52*5+3, ystart-365-14*max_extent-4, width=200, height=1)

			else:
				canvas.drawImage("resources/intermittent.png", kolumny["A"], ystart-365-14*max_extent-4, width=200, height=1)
				canvas.drawImage("resources/intermittent.png", xstart+52*3+3, ystart-365-14*max_extent-4, width=200, height=1)

		def build_with_vat():
			""" 
			Generuje fakturę z podatkiem vat.
			"""
			nonlocal canvas, wiersz, kolumny, suma_netto, suma_brutto, suma_vat
			nazwy_cech = ["nazwa", "ilosc", "cena", "uwagi"]

			for nr_produktu, produkt in enumerate(zakupy.get_children()):

				# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
				cechy_produktu = {"kod": str(zakupy.item(produkt)["text"])}
				for nr_cechy, cecha in enumerate(zakupy.item(produkt)['values']):
					cechy_produktu[nazwy_cech[nr_cechy]] = cecha

				# Obliczanie stawek
				stawka_vat = Decimal(Model.get_product_vat(cechy_produktu["nazwa"]))
				kwota_brutto = Decimal(float(cechy_produktu["cena"]))
				kwota_vat = kwota_brutto*(stawka_vat/100)
				kwota_netto = kwota_brutto*(stawka_vat/100)+kwota_brutto

				# Wpisywanie do tabeli
				canvas.drawString(kolumny["A"], ystart-365-14*wiersz, str(nr_produktu))
				canvas.drawString(kolumny["B"], ystart-365-14*wiersz, str(cechy_produktu["kod"]))
				wpisz_nazwe(wiersz, cechy_produktu["nazwa"])
				canvas.drawString(kolumny["D"], ystart-365-14*wiersz, str(cechy_produktu["ilosc"]))
				canvas.drawString(kolumny["E"], ystart-365-14*wiersz, "{0:.2f}".format(kwota_brutto))
				canvas.drawString(kolumny["F"], ystart-365-14*wiersz, "{0:.2f}".format(stawka_vat))
				canvas.drawString(kolumny["G"], ystart-365-14*wiersz, "{0:.2f}".format(kwota_vat))
				canvas.drawString(kolumny["H"], ystart-365-14*wiersz, "{0:.2f}".format(kwota_netto))
				
				# Sumowanie stawek
				suma_netto += kwota_brutto*(stawka_vat/100)+kwota_brutto
				suma_brutto += kwota_brutto
				suma_vat += kwota_brutto*(stawka_vat/100)

				kolejny_wiersz(cechy_produktu["nazwa"])

			# Sumowanie kosztów
			price_calc_lvl = wiersz + 1

			canvas.drawString(xstart+70*3-3, ystart-365-14*(price_calc_lvl), "Razem:")
			canvas.drawString(xstart+52*5-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_brutto))

			canvas.drawString(xstart+52*7-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_vat))
			canvas.drawString(xstart+52*8-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_netto))

			# Wstawianie kwoty do zapłaty
			canvas.drawString(xstart+78*1+20, ystart-325, "{0:.2f}".format(suma_netto))

			# Rysowanie miejsc na podpisy
			canvas.drawString(xstart, ystart-365-14*(price_calc_lvl+4), "___________________________________")
			canvas.drawString(xstart+13, ystart-365-14*(price_calc_lvl+5), "Podpis osoby wystawiającej fakturę")


		def build_without_vat():
			""" 
			Generuje fakturę bez podatku vat.
			"""
			nonlocal canvas, wiersz, kolumny, suma_netto, suma_brutto, suma_vat
			nazwy_cech = ["nazwa", "ilosc", "cena", "uwagi"]

			for nr_produktu, produkt in enumerate(zakupy.get_children()):

				# Otrzymywanie własności produktu (np. nazwa, cena, kod, vat)
				cechy_produktu = {"kod": str(zakupy.item(produkt)["text"])}
				for nr_cechy, cecha in enumerate(zakupy.item(produkt)['values']):
					cechy_produktu[nazwy_cech[nr_cechy]] = cecha
				print(cechy_produktu)

				# Obliczanie stawek
				kwota_brutto = Decimal(float(cechy_produktu["cena"]))

				# Wpisywanie do tabeli
				canvas.drawString(kolumny["A"], ystart-365-14*wiersz, str(nr_produktu))
				canvas.drawString(kolumny["B"], ystart-365-14*wiersz, str(cechy_produktu["kod"]))
				wpisz_nazwe(wiersz, cechy_produktu["nazwa"])
				canvas.drawString(kolumny["D"], ystart-365-14*wiersz, str(cechy_produktu["ilosc"]))
				canvas.drawString(kolumny["E"], ystart-365-14*wiersz, "{0:.2f}".format(kwota_brutto))
	
				# Sumowanie stawek
				suma_brutto += kwota_brutto

				kolejny_wiersz(cechy_produktu["nazwa"])

			# Sumowanie kosztów
			price_calc_lvl = wiersz + 1

			canvas.drawString(xstart+70*3-3, ystart-365-14*(price_calc_lvl), "Razem:")
			canvas.drawString(xstart+52*5-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_brutto))

			# Wstawianie kwoty do zapłaty
			canvas.drawString(xstart+78*1+20, ystart-325, "{0:.2f}".format(suma_brutto))

			# Rysowanie miejsc na podpisy
			canvas.drawString(xstart, ystart-365-14*(price_calc_lvl+4), "___________________________________")
			canvas.drawString(xstart+13, ystart-365-14*(price_calc_lvl+5), "Podpis osoby wystawiającej fakturę")


		# Wypisywanie produktów na fakturze
		if include_vat:
			build_with_vat()
		else:
			build_without_vat()

		# Finalizacja
		canvas.save()
		print("[Productor] Pomyślnie zapisano fakturę w pliku {}".format(self.nazwa))

