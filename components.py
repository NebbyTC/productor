import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename
import platform
from typing import Union


class Separator(tk.Frame):
	""" Widget oddzielający widgety poziomo """

	def __init__(self, master: tk.Frame, color: str = "grey2"):
		""" 
		Tworzy widget i sprawdza poprawność 
		argumentów.

		Args:
			master: 
				Okno rodzic.
			color: 
				Kolor widżetu, domyślnie 'grey2'.

		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(color, str):
			raise ValueError("[ProductorInternal] Argument 'color' nie jest typu 'str'.")

		super().__init__(master, bg=color, relief=tk.RAISED, borderwidth=4, height=1)


class LabeledEntry(tk.Frame):
	""" Podpisany komponent wejścia """

	def __init__(self, master: tk.Frame, caption: str, spacing: int = 110):
		""" 
		Tworzy komponent i sprawdza poprawność 
		argumentów.

		Args:
			master: 
				Okno rodzic.
			caption:
				Podpis komponent wejścia.

		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(master, tk.Frame):
			raise ValueError("[ProductorInternal] Argument 'master' nie jest typu 'tk.Frame'.")

		elif not isinstance(spacing, int):
			raise ValueError("[ProductorInternal] Argument 'spacing' nie jest typu 'int'.")

		elif not isinstance(caption, str):
			raise ValueError("[ProductorInternal] Argument 'caption' nie jest typu 'str'.")

		super().__init__(master)

		self.label = tk.Label(self, text=caption)
		self.entry = tk.Entry(self, width=24)

		self.grid_columnconfigure(0, minsize=spacing)
		self.grid_columnconfigure(1, minsize=spacing)

		self.label.grid(row=0, column=0, sticky=tk.W)
		self.entry.grid(row=0, column=1, padx=4)


	def get(self) -> str:
		""" 
		Zwraca wartość w komponencie 'self.entry'.

		Returns:
			Wartość w komponencie 'self.entry'.
		"""

		return self.entry.get()


	def insert(self, text: str) -> None:
		""" 
		Wstawia podaną wartość do pola
		tekstowego kontrolki.

		Args:
			text:
				Wartość tekstowa do
				wstawienia.

		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(text, str):
			raise ValueError("[ProductorInternal] Argument 'text' nie jest typu 'str'.")

		self.entry.insert(tk.END, text)


	def disable(self) -> None:
		""" 
		Wyłącza komponent z użytku.
		"""
		self.entry.config(state=tk.DISABLED)


	def enable(self) -> None:
		""" 
		Dopuszcza komponent z użytku.
		"""
		self.entry.config(state=tk.NORMAL)


class Checkbox(tk.Checkbutton):
	""" Bardziej intuicyjna kontrolka guizka prawda/fałsz """

	def __init__(self, master: tk.Frame, caption: str, command: "function" = None):
		""" 
		Tworzy komponent na podstawie argumentów.

		Args:
			master: 
				Kontrolka/ramka rodzic.

			caption:
				Podpis kontrolki.

			command:
				Argument opcjonalny, zawiera
				funkcję wywoływaną podczas
				zdarzenia kliknięcia.
		"""

		if not isinstance(master, tk.Frame):
			raise ValueError("[ProductorInternal] Argument 'master' nie jest typu 'tk.Frame'.")

		elif not isinstance(caption, str):
			raise ValueError("[ProductorInternal] Argument 'caption' nie jest typu 'str'.")

		self.value = tk.IntVar()
		super().__init__(master, text=caption, variable=self.value, onvalue=1, offvalue=0, command=command)


	def get(self) -> bool:
		""" 
		Zwraca wartość kontrolki.

		Returns:
			Wartość kontrolki (true/false).
		"""

		return bool(self.value.get())

	def set(self, value: bool) -> None:
		"""
		Ustawia wartość kontrolki.

		Args:
			value:
				Wartość true/false komórki
				
		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(value, bool):
			raise ValueError("[ProductorInternal] Argument 'value' nie jest typu 'bool'.")

		self.value.set(int(value))


class PathEntry(LabeledEntry):
	""" Komponent wejścia z wyszukiwarką plików """

	def __init__(self, master: tk.Frame, caption: str):
		""" 
		Tworzy komponent i dziedziczy 

		Args:
			master: 
				Okno rodzic.
			caption:
				Podpis komponent wejścia.
		"""
		super().__init__(master, caption)

		self.button = tk.Button(self, text="Szukaj", command=self.__open)
		self.button.grid(row=0, column=2)

	def __open(self) -> None:
		""" 
		Przekazuje na wejście wyszukaną
		w eksploratorze scieżkę.
		"""

		path = askopenfilename(
			filetypes = (
				("Arkusz programu Excel","*.xlsx*"),
				("Wszystkie pliki","*.*")
			), 
			parent = self.master
		)

		if platform.system() == "Windows":
			path = path.replace("/", "\\")

		else:
			path =  path.replace("\\", "/")

		self.entry.delete(0, tk.END)
		self.entry.insert(tk.END, path)


class LabeledCombobox(tk.Frame):
	""" Combobox z podpisem """

	def __init__(self, master: tk.Frame, options: list[str], caption: str, spacing: int = 110):
		""" 
		Tworzy komponent i sprawdza poprawność 
		argumentów.

		Args:
			master: 
				Okno rodzic.
			options:
				Opcje wyboru widżetu.
			caption:
				Podpis komponent wejścia.

		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if not isinstance(master, tk.Frame):
			raise ValueError("[ProductorInternal] Argument 'master' nie jest typu 'tk.Frame'.")

		elif not isinstance(spacing, int):
			raise ValueError("[ProductorInternal] Argument 'spacing' nie jest typu 'int'.")

		elif not isinstance(caption, str):
			raise ValueError("[ProductorInternal] Argument 'caption' nie jest typu 'str'.")

		elif not isinstance(options, list):
			raise ValueError("[ProductorInternal] Argument 'options' nie jest typu 'list'.")

		super().__init__(master)
		self.__options = [str(elem) for elem in options]

		self.label = tk.Label(self, text=caption)
		self.combo = ttk.Combobox(self, width=21)

		self.combo["values"] = self.__options
		self.combo.bind("<<ComboboxSelected>>", lambda e: self.focus())

		self.grid_columnconfigure(0, minsize=spacing)
		self.grid_columnconfigure(1, minsize=spacing)

		self.label.grid(row=0, column=0, sticky=tk.W)
		self.combo.grid(row=0, column=1, padx=4)


	def get(self) -> str:
		""" 
		Zwraca wartość w komponencie 'self.combo'.

		Returns:
			Wartość w komponencie 'self.combo'.
		"""

		return self.combo.get()


	def select(self, option: Union[int, str]) -> None:
		""" 
		Wstawia podaną wartość do pola
		tekstowego kontrolki.

		Args:
			text:
				Wartość tekstowa do
				wstawienia.

		Raises:
			ValueError:
				Zwracany, jeżeli typy argumentów 
				nie są poprawne.
		"""

		if isinstance(option, str):
			self.combo.current(self.__options.index(option))

		elif isinstance(option, int):
			self.combo.current(option)

		else:
			raise ValueError("[ProductorInternal] Argument 'option' nie jest typu 'str' ani 'int'.")
