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


	def __init__(self, nazwa: str, informacje: list, uwzglednij_vat: bool):
		self.nazwa, self.info, self.vat = nazwa, informacje, uwzglednij_vat
		self.os = platform.system()


	def build(self):
		""" Generuje fakturę uwzględniając atrybuty obiektu """
		if self.vat:
			self.template_file = "template_vat"
		else:
			self.template_file = "template"

		if ".pdf" in self.nazwa:
			self.build_pdf(include_vat=self.vat)

		elif ".xlsx" in self.nazwa:
			self.build_xlsx(include_vat=self.vat)


	def build_xlsx(self, include_vat=False):
		""" Odpowiada za stworzenie faktury w formacie .xlsx """
		if self.os == "Windows":
			os.system(f'copy resources\\{self.template_file}.xlsx {self.nazwa}')

		else:
			os.system(f'cp resources/{self.template_file}.xlsx {self.nazwa}')

		skoroszyt = openpyxl.load_workbook(filename=self.nazwa)
		arkusz = skoroszyt.active

		# Wstawianie tabeli
		DOSTAWA = Decimal(29.99)
		zakupy = self.info[12]
		suma = Decimal(0.0)
		suma_netto = 0.0
		suma_vat = 0.0

		if include_vat:
			counter = 32
			wiersze = 0

			for lp, wiersz in enumerate(zakupy.get_children()):

				values = [str(zakupy.item(wiersz)["text"])]
				for value in zakupy.item(wiersz)['values']:
					values.append(value)

				suma += Decimal(float(values[3]))
				stawka_vat = Model.get_product_vat(str(values[1]))
				cena_brutto = float(values[3])

				# Wpisywanie do tabeli
				arkusz[f"A{counter}"] = lp
				arkusz[f"B{counter}"] = str(values[0])
				arkusz[f"C{counter}"] = str(values[1])
				arkusz[f"E{counter}"] = str(values[2])
				arkusz[f"F{counter}"] = "{0:.2f}".format(cena_brutto)
				arkusz[f"G{counter}"] = "{0:.2f}".format(stawka_vat)
				arkusz[f"H{counter}"] = "{0:.2f}".format(cena_brutto*(stawka_vat/100))
				arkusz[f"I{counter}"] = "{0:.2f}".format(cena_brutto*(stawka_vat/100)+cena_brutto)

				arkusz[f"A{counter}"].alignment = Alignment(horizontal='left')
				
				suma_netto += cena_brutto*(stawka_vat/100)+cena_brutto
				suma_vat += cena_brutto*(stawka_vat/100)

				counter += 1
				wiersze += 1

		else:
			counter = 31
			wiersze = 0

			for lp, wiersz in enumerate(zakupy.get_children()):

				values = [str(zakupy.item(wiersz)["text"])]
				for value in zakupy.item(wiersz)['values']:
					values.append(value)

				suma += Decimal(float(values[3]))

				# Wpisywanie do tabeli
				arkusz[f"A{counter}"] = lp
				arkusz[f"B{counter}"] = str(values[0])
				arkusz[f"C{counter}"] = str(values[1])
				arkusz[f"E{counter}"] = str(values[2])
				arkusz[f"F{counter}"] = str(values[3])
				arkusz[f"G{counter}"] = str(values[4])

				arkusz[f"A{counter}"].alignment = Alignment(horizontal='left')
				
				counter += 1
				wiersze += 1

		if include_vat:
			arkusz[f"E{counter+1}"] = "Razem:"

			arkusz[f"F{counter+1}"] = "{0:.2f}".format(suma)

			arkusz[f"H{counter+1}"] = "{0:.2f}".format(suma_vat)
			arkusz[f"I{counter+1}"] = "{0:.2f}".format(suma_netto)

		else:
			arkusz[f"E{counter+1}"] = "Razem:"

			arkusz[f"F{counter+1}"] = "{0:.2f}".format(suma)


		# Miejsce na podpis
		thin = Side(border_style="thin", color="000000")

		arkusz[f"B{counter+3+1}"].border = Border(bottom=thin)
		arkusz[f"C{counter+3+1}"].border = Border(bottom=thin)
		arkusz[f"D{counter+3+1}"].border = Border(bottom=thin)
		arkusz[f"E{counter+3+1}"].border = Border(bottom=thin)

		arkusz.merge_cells(f"B{counter+4+1}:E{counter+4+1}")
		arkusz[f"B{counter+4+1}"].alignment = Alignment(horizontal='center')
		arkusz[f"B{counter+4+1}"] = "Podpis osoby wystawiającej fakturę"


		# Dane czasowe
		arkusz["G5"] = str(datetime.datetime.now()).split(" ")[0]
		arkusz["G6"] = str(datetime.datetime.now()).split(" ")[0]
		arkusz["G7"] = str(datetime.datetime.now()).split(" ")[0]

		# Dane sprzedawcy
		arkusz["A6"] = self.info[0]
		arkusz["A7"] = self.info[1]
		arkusz["A8"] = self.info[2]
		arkusz["A9"] = self.info[3]
		arkusz["A10"] = self.info[4]

		# Dane nabywcy
		arkusz["A13"] = self.info[5]
		arkusz["A14"] = self.info[6]
		arkusz["A15"] = self.info[7]
		arkusz["A16"] = self.info[8]

		# Dane odbiorcy
		arkusz["A19"] = self.info[9]
		arkusz["A20"] = self.info[11]
		arkusz["A21"] = self.info[7]
		arkusz["A22"] = self.info[6]
		arkusz["A23"] = self.info[8]

		# Sposó dostawy
		arkusz["C25"] = self.info[9]

		# Forma platnosc, kwota, waluta
		arkusz["A28"] = self.info[10]

		if include_vat:
			arkusz["C28"] = "{0:.2f}".format(suma_netto)
		else:
			arkusz["C28"] = "{0:.2f}".format(suma)
		arkusz["D28"] = "PLN"

		skoroszyt.save(self.nazwa)
		print("[Productor] Pomyślnie zapisano fakturę w pliku {}".format(self.nazwa))


	def build_pdf(self, include_vat=False):
		""" Odpowiada za stworzenie faktury w formacie .pdf """
		print(self.nazwa)
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

		# Sprzedawca
		canvas.drawString(xstart, ystart, self.info[0])
		canvas.drawString(xstart, ystart-14, self.info[1])
		canvas.drawString(xstart, ystart-14*2, self.info[2])
		canvas.drawString(xstart, ystart-14*3, self.info[3])
		canvas.drawString(xstart, ystart-14*4, self.info[4])

		# Nadawca
		canvas.drawString(xstart, ystart-101, self.info[5])
		canvas.drawString(xstart, ystart-101-14, self.info[6])
		canvas.drawString(xstart, ystart-101-14*2, self.info[7])
		canvas.drawString(xstart, ystart-101-14*3, self.info[8])

		# Odbiorca
		canvas.drawString(xstart, ystart-188, self.info[9])
		canvas.drawString(xstart, ystart-188-14, self.info[11])
		canvas.drawString(xstart, ystart-188-14*2, self.info[7])
		canvas.drawString(xstart, ystart-188-14*3, self.info[6])
		canvas.drawString(xstart, ystart-188-14*4, self.info[8])

		# Dolne małe informacje
		canvas.drawString(xstart+85, ystart-278, self.info[9])

		canvas.drawString(xstart, ystart-325, self.info[10])

		canvas.drawString(xstart+78*2, ystart-325, "PLN")

		# Wykaz produktów
		DOSTAWA = Decimal(self.info[13])
		zakupy = self.info[12]
		suma = Decimal(0.0)
		suma_kwoty_vat = 0.0
		suma_netto = 0.0
		counter = 31
		wiersze = 0

		for lp, wiersz in enumerate(zakupy.get_children()):

			values = [str(zakupy.item(wiersz)["text"])]
			for value in zakupy.item(wiersz)['values']:
				values.append(value)

			suma += Decimal(float(values[3]))

			# Wpisywanie do tabeli
			if include_vat:
				stawka_vat = Model.get_product_vat(values[1])
				cena_brutto = float(values[3])

				canvas.drawString(xstart, ystart-365-14*(lp+1), str(lp))

				canvas.drawString(xstart+52, ystart-365-14*(lp+1), str(values[0]))
				canvas.drawString(xstart+52*2, ystart-365-14*(lp+1), str(values[1]))
				canvas.drawString(xstart+70*3-3, ystart-365-14*(lp+1), str(values[2]))
				canvas.drawString(xstart+52*5-3, ystart-365-14*(lp+1), str(cena_brutto))
				canvas.drawString(xstart+52*6-3, ystart-365-14*(lp+1), str(stawka_vat))
				canvas.drawString(xstart+52*7-3, ystart-365-14*(lp+1), "{0:.2f}".format(cena_brutto*(stawka_vat/100)))
				canvas.drawString(xstart+52*8-3, ystart-365-14*(lp+1), "{0:.2f}".format(cena_brutto*(stawka_vat/100)+cena_brutto))

				suma_kwoty_vat += float("{0:.2f}".format(cena_brutto*(stawka_vat/100)))
				suma_netto += float("{0:.2f}".format(cena_brutto*(stawka_vat/100)+cena_brutto))

			else:
				canvas.drawString(xstart, ystart-365-14*lp, str(lp))

				canvas.drawString(xstart+52, ystart-365-14*lp, str(values[0]))
				canvas.drawString(xstart+52*2, ystart-365-14*lp, str(values[1]))
				canvas.drawString(xstart+70*3-3, ystart-365-14*lp, str(values[2]))
				canvas.drawString(xstart+52*5-3, ystart-365-14*lp, str(values[3]))
				canvas.drawString(xstart+52*6-3, ystart-365-14*lp, str(values[4]))
			
			counter += 1
			wiersze += 1

		# Sumowanie kosztów
		if include_vat:
			price_calc_lvl = wiersze + 2

			canvas.drawString(xstart+70*3-3, ystart-365-14*(price_calc_lvl), "Razem:")
			canvas.drawString(xstart+52*5-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma))

			canvas.drawString(xstart+52*7-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_kwoty_vat))
			canvas.drawString(xstart+52*8-3, ystart-365-14*(price_calc_lvl), "{0:.2f}".format(suma_netto))

		else:
			price_calc_lvl = wiersze + 4 		

			canvas.drawString(xstart+52*6-3, ystart-365-14*price_calc_lvl, "Koszt:")
			canvas.drawString(xstart+52*6-3, ystart-365-14*(price_calc_lvl+1), "{0:.2f}".format(DOSTAWA))
			canvas.drawString(xstart+52*6-3, ystart-365-14*(price_calc_lvl+2), "{0:.2f}".format(suma+DOSTAWA))

			canvas.drawString(xstart+52*5-3, ystart-365-14*(price_calc_lvl+1), "Dostawa:")
			canvas.drawString(xstart+52*5-3, ystart-365-14*(price_calc_lvl+2), "Razem:")

			canvas.drawString(xstart+104, ystart-325, "{0:.2f}".format(suma))

		# Wstawianie kwoty do zapłaty
		canvas.drawString(xstart+78*1+20, ystart-325, "{0:.2f}".format(suma_netto))

		# Rysowanie miejsc na podpisy
		canvas.drawString(xstart, ystart-365-14*(price_calc_lvl+4), "___________________________________")
		canvas.drawString(xstart+13, ystart-365-14*(price_calc_lvl+5), "Podpis osoby wystawiającej fakturę")

		# Finalizacja
		canvas.save()
		print("[Productor] Pomyślnie zapisano fakturę w pliku {}".format(self.nazwa))
