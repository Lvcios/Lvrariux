# _*_coding:utf8_*_
#librerias a instalar 
#xlrd, xlwt, xlutils
import xlrd
from Tkinter import *
import tkFileDialog

class Ventana:
	def __init__(self):
		self.root = Tk()
		self.fileName = ''
		self.libro = ''
		self.hoja = ''
	
	def inicio(self):
		self.root.mainloop()
	
	def botones(self):
		btnCarga = Button(self.root, text = 'Cargar archivo', command = self.cargarArchivo , width = 30)
		btnCrea = Button(self.root, text = 'Crear Horario',command = self.setExcelBook, width = 30)
		btnCarga.grid(row = 1, column = 0)
		btnCrea.grid(row = 3, column = 0)
	
	def cargarArchivo(self):
		file = tkFileDialog.askopenfile(mode = 'r', title='Elije un archivo', filetypes = [('Excel File','*.xlsx'),])
		if file != None:
			self.fileName = str(file.name)
			fileLabel = Label(self.root, text = file.name )
			fileLabel.grid(row = 2, column = 0)
	
	def getFileName(self):
		return self.fileName
	
	def setExcelBook(self):
		self.libro = xlrd.open_workbook(self.getFileName())
		self.hoja = self.libro.sheet_by_index(0)
		print 'carga correcta'
"""
class Horario:
	def __init__(self, libro):
		self.libro = xlrd.open_workbook(libro)
		self.hoja = libro.sheet_by_index(0)
"""
ventana =  Ventana()
ventana.botones()
ventana.inicio()
