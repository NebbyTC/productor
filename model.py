import sqlite3


class Model:
	""" Klasa odpowiedzialna za obsługę danych """


	@staticmethod
	def get_products() -> tuple:
		""" Zwraca wszystkie dane z bazy danych "produkty" """
		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute("SELECT * FROM produkty")
			return cursor.fetchall()


	@staticmethod
	def get_last_product() -> list:
		""" Zwraca ostatni produkt z bazy danych """
		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute("SELECT * FROM produkty")
			return cursor.fetchall()[-1]


	@staticmethod
	def add_product(nazwa: str, typ: str, cena: float, stan: bool, vat: float = 0.00) -> None:
		""" Dodaje wpis do bazy danych """
		nazwa, typ, cena, stan, vat = nazwa, typ, float(cena), int(stan), float(vat)

		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""INSERT INTO produkty VALUES (NULL, "{nazwa}", "{typ}", {cena}, {stan}, {vat})"""
			)
			conn.commit()


	@staticmethod
	def delete_product(id_: int) -> None:
		""" Usuwa z bazy danych na podstawie podanego id """
		id_ = int(id_)

		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f""" DELETE FROM produkty WHERE id = "{id_}"; """
			)

			# Zerowanie id (jeżeli wszystkie produkty zostały usunięte)
			cursor.execute(
				""" UPDATE SQLITE_SEQUENCE SET SEQ=0 WHERE NAME='produkty'; """
			)

			conn.commit()


	@staticmethod
	def update_product(id_: int, nazwa: str, typ: str, cena: float, stan: bool, vat: float) -> None:
		""" Aktualizuje inforamacje o produkcie """
		id_, nazwa, typ, cena, stan, vat = int(id_), str(nazwa), str(typ), float(cena), int(stan), float(vat)

		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""UPDATE produkty\n SET id='{id_}', nazwa='{nazwa}', typ='{typ}', cena='{cena}', stan={stan}, vat='{vat}'\n WHERE id='{id_}';"""
			)
			conn.commit()


	@staticmethod
	def get_product_id(nazwa: str) -> int:
		""" Zwraca id podanego produktu """
		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT id FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()

		return cursor.fetchall()[0][0]


	@staticmethod
	def get_product_price(nazwa: str) -> float:
		""" Zwraca cenę podanego produktu """
		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT cena FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()
		
		return float(cursor.fetchall()[0][0])


	@staticmethod
	def get_product_vat(nazwa: str) -> float:
		""" Zwraca cenę podanego produktu """
		with sqlite3.connect("baza_danych.db") as conn:
			cursor = conn.cursor()

			cursor.execute(
				f"""SELECT vat FROM produkty WHERE nazwa=="{nazwa}" """
			)
			conn.commit()
		
		return float(cursor.fetchall()[0][0])
