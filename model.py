import sqlite3


class Model:
	""" Klasa odpowiedzialna za obsługę danych """

	PLIK_USTAWIEN = "settings.txt"
	PLIK_BAZY_DANYCH = "baza_danych.db"


	@classmethod
	def get_products(cls) -> tuple:
		""" Zwraca wszystkie dane z bazy danych "produkty" """
		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute("SELECT * FROM produkty")
			return cursor.fetchall()


	@classmethod
	def get_last_product(cls) -> list:
		""" Zwraca ostatni produkt z bazy danych """
		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute("SELECT * FROM produkty")
			return cursor.fetchall()[-1]


	@classmethod
	def add_product(cls, nazwa: str, typ: str, cena: float, stan: bool, vat: float = 0.00) -> None:
		""" Dodaje wpis do bazy danych """
		nazwa, typ, cena, stan, vat = nazwa, typ, float(cena), int(stan), float(vat)

		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""INSERT INTO produkty VALUES (NULL, "{nazwa}", "{typ}", {cena}, {stan}, {vat})"""
			)
			conn.commit()


	@classmethod
	def delete_product(cls, id_: int) -> None:
		""" Usuwa z bazy danych na podstawie podanego id """
		id_ = int(id_)

		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f""" DELETE FROM produkty WHERE id = "{id_}"; """
			)

			# Zerowanie id (jeżeli wszystkie produkty zostały usunięte)
			cursor.execute(
				""" UPDATE SQLITE_SEQUENCE SET SEQ=0 WHERE NAME='produkty'; """
			)

			conn.commit()


	@classmethod
	def update_product(cls, id_: int, nazwa: str, typ: str, cena: float, stan: bool, vat: float) -> None:
		""" Aktualizuje inforamacje o produkcie """
		id_, nazwa, typ, cena, stan, vat = int(id_), str(nazwa), str(typ), float(cena), int(stan), float(vat)

		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""UPDATE produkty\n SET id='{id_}', nazwa='{nazwa}', typ='{typ}', cena='{cena}', stan={stan}, vat='{vat}'\n WHERE id='{id_}';"""
			)
			conn.commit()


	@classmethod
	def get_product_id(cls, nazwa: str) -> int:
		""" Zwraca id podanego produktu """
		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT id FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()

		return cursor.fetchall()[0][0]


	@classmethod
	def get_product_price(cls, nazwa: str) -> float:
		""" Zwraca cenę podanego produktu """
		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT cena FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()
		
		return float(cursor.fetchall()[0][0])


	@classmethod
	def get_product_vat(cls, nazwa: str) -> float:
		""" Zwraca cenę podanego produktu """
		with sqlite3.connect(cls.PLIK_BAZY_DANYCH) as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT vat FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()
		
		return float(cursor.fetchall()[0][0])


	@classmethod
	def get_settings(cls) -> dict:
		""" 
		Zwraca ustawienia w postaci słownika.
		
		Returns:
			Słownik zawierający ustawienia
			programu.
		"""

		try:
			with open(cls.PLIK_USTAWIEN, "r+", encoding="utf-8") as r:
				tresc = r.read()
		except FileNotFoundError:
			open(cls.PLIK_USTAWIEN, "a+", encoding="utf-8").close()
			with open(cls.PLIK_USTAWIEN, "a+") as w:
				w.write("sprzedawca: '';\nfirma-sprzedawcy: '';\nemail: '';\nkod-pocztowy: '';\nadres: '';\ndostawa: '0.0';\nuwzglednij-dostawe: 'false';")
				tresc = w.read()

		linijki = [x.replace("\n", "") for x in tresc.split(";")]
		ustawienia = {}

		for linijka in linijki:
			if linijka:
				opcja, wartosc = linijka.split(":")
				ustawienia[opcja] = eval(wartosc.strip())

		return ustawienia


	@classmethod
	def update_settings(cls, ustawienia: dict) -> None:
		""" 
		Aktualizuje plik z ustawieniami.

		Args:
			ustawienia:
				Słownik zawierający wartości
				do wpisania do pliku z
				ustawieniami.
		
		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(ustawienia, dict):
			raise ValueError("[ProductorInternal] Argument 'ustawienia' nie jest typu 'dict'.")

		with open(cls.PLIK_USTAWIEN, "w+", encoding="utf-8") as w:
			w.write("")

		with open(cls.PLIK_USTAWIEN, "w", encoding="utf-8") as w:
			for klucz, wartosc in ustawienia.items(): 
				w.write(f"{klucz}: '{wartosc}';\n")
