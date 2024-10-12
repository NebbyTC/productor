import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
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


OPEN = "pik.open"
SAVE = "pik.save"

class PathEntry(LabeledEntry):
	""" Komponent wejścia z wyszukiwarką plików """

	def __init__(self, master: tk.Frame, caption: str, mode: str = OPEN):
		""" 
		Tworzy komponent i dziedziczy 

		Args:
			master: 
				Okno rodzic.
			caption:
				Podpis komponent wejścia.
		"""
		super().__init__(master, caption)

		self.mode = mode
		self.defualt_extension = ".xlsx"
		self.extensions = (
			("Arkusz programu Excel","*.xlsx"),
			("Wszystkie pliki","*.*")
		)

		self.button = tk.Button(self, text="Szukaj", command=self.__open)
		self.button.grid(row=0, column=2)


	def __open(self) -> None:
		""" 
		Przekazuje na wejście wyszukaną
		w eksploratorze scieżkę.
		"""

		if self.mode == OPEN:
			path = askopenfilename(
				filetypes = self.extensions,
				parent = self.master
			)
		elif self.mode == SAVE:
			path = asksaveasfilename(
				filetypes = self.extensions,
				defaultextension = self.defualt_extension,
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


class ProgressToplevel():
	""" 
	Klasa wzór okna, które realizuje w tle
	funkcjonalność ładowania
	"""


	def set_progressbar(self, wiersze):
		""" Ustawia pasek ładowania """
		self.progress = ttk.Progressbar(self.excel_progress_panel)
		self.status_text = tk.Label(self.excel_progress_panel, text="Status...")

		self.progress.pack(fill=tk.X, pady=4, padx=4)
		self.status_text.pack(side=tk.TOP, padx=4, pady=4, anchor=tk.W)


	def truncate_name(self, string, limit):
		""" Przycina podanego stringa jeżeli jest dłuższy niż podany limit. """
		return (string[:(limit - 2)] + '..') if len(string) > (limit - 2) else string




class _AutocompletionLabel(tk.Label):
	""" A label displaying what autocompletion hints """

	def __init__(self, *args, **kwagrs):
		super().__init__(*args, **kwagrs)

		self.bind("<Button-1>", self.left_click)


	def left_click(self, e):
		self.master.belongs_to.delete(0, tk.END)
		self.master.belongs_to.insert(0, self["text"])
		self.master.belongs_to.destroy_popups()


class AutoCombobox(ttk.Combobox):# <-- later change to a more broad Autocompletion class for all input serving widgets
	""" A combobox with autocompletion feathures """

	def __init__(self, master, root_window):
		super().__init__(master)
		self.bind("<KeyRelease>", lambda e: self.display_autocompletions())
		self.bind("<FocusOut>", lambda e: self.destroy_popups())
		self.bind("<FocusIn>", lambda e: self.display_autocompletions())

		self.popups = []
		self.root_window = root_window
		self.previous_autocompletions = None


	def destroy_popups(self):
		""" 
		Responsible for destroying all the popups with 
		autocompletion hints 
		"""

		self.previous_autocompletions = None

		for elem in self.popups:
			elem.destroy()


	def truncate_name(self, string, limit):
		""" Przycina podanego stringa jeżeli jest dłuższy niż podany limit. """
		return (string[:(limit - 2)] + '..') if len(string) > (limit - 2) else string


	def display_autocompletions(self):
			""" Displays matching autompletions bellow of the Combobox widget """
			
			def get_matching_autocompletions(inputed_text):
				""" Returns autocompletions matching with given text """
				matching_autocompletions = []

				# Zwróć, jeśli informacja zaczyna się tak jak inputed_text
				for hint in self['values']:
					if hint.startswith(inputed_text):
						matching_autocompletions.append(hint)

				return matching_autocompletions

			def insert_autocompletion(autocompletions):
				""" Inserts the top autocompletion into the combobox """

				self.delete(0, tk.END)
				self.insert(0, autocompletions[0])

				self.destroy_popups()

			# Getting the autocompletions
			inputed_text = self.get()
			if not inputed_text: 
				self.destroy_popups()
				return

			autocompletions = get_matching_autocompletions(inputed_text)
			if not len(autocompletions):
				self.destroy_popups()
				return

			if self.previous_autocompletions == autocompletions:
				return

			self.bind("<Tab>", lambda e: insert_autocompletion(autocompletions))

			# Positioning the space for autocompletions right
			popup = tk.Toplevel()
			popup.attributes('-topmost',True)
			popup.belongs_to = self

			x = self.winfo_rootx()
			y = self.winfo_rooty() + self.winfo_height()
	
			popup.wm_overrideredirect(True)
			popup.wm_geometry("+%d+%d" % (x, y))

			self.destroy_popups()

			# Placing the autocompletions
			total_height = 0
			counter = 0

			for autocompletion in autocompletions:
				hint_label = tk.Label(
					popup, text = self.truncate_name(autocompletion, 20), anchor = 'w', bg="white"
				)
				hint_label.pack(fill=tk.X)
				popup.update()
				total_height += hint_label.winfo_height()
				counter += 1
				if counter == 1: # <-- display max. 3 autocompletion hints
					break

			popup.geometry(f"{self.winfo_width()}x{total_height}")
			self.root_window.moving_parts.append(popup)
			self.popups.append(popup)
			self.previous_autocompletions = autocompletions

				