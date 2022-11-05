import tkinter as tk
import sys

import view
import controller


class App(view.Gui):
	""" Klasa główna całego programu, łącząca wszystkie trzy elementy """


	def __init__(self):
		""" Załączenie wszystkich trzech modułów razem """
		kontroler = controller.Controller(self)
		self.set_controller(kontroler)

		super().__init__()


if __name__ == '__main__':
	app = App()


""" 
TODO:

1. Dodać dostawę jako produkt na fakturze
"""