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
		self.clasesPorDia = range(450)
		self.Salones = range(5) #dias
		self.record = 0
	
	def inicio(self):
		self.root.mainloop()
	
	def botones(self):
		btnCarga = Button(self.root, text = 'Cargar archivo', command = self.cargarArchivo , width = 30)
		btnCrea = Button(self.root, text = 'Crear Horario',command = self.creaHorario, width = 30)
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
		#print 'carga de libro de excel correcta'
		#self.setMatrix()
		#print 'creacion de matrices correcta'
		
	def setMatrix(self):
		for i in range(len(self.clasesPorDia)):
			self.clasesPorDia[i]=range(9)#id, secuencia, materia	
		
		for i in range(len(self.Salones)):
			self.Salones[i] = range(100) #100 salones disponibles, aunque solo 11 realmente
	
		for i in range(len(self.Salones)):
			for j in range(len(self.Salones[i])):
				self.Salones[i][j] = range(30) #30 periodos para cada salon
		
		for i in range(len(self.Salones)):
			for j in range(len(self.Salones[i])):
				for k in range(len(self.Salones[i][j])):
					self.Salones[i][j][k] = 0 #periodo igual a 0 de inicio, indica que el periodo no ha sido asignado a ninguna clase.
	
	def classRecord(self, dia):
		sec_temp = ''
		mat_temp = ''
		i = 1
		for rx in range(1, self.hoja.nrows):
			if sec_temp != self.hoja.row(rx)[0].value:
				if not self.hoja.row(rx)[dia].value == '':
					self.clasesPorDia[i][0] = i #id
					self.clasesPorDia[i][1] = self.hoja.row(rx)[0].value #Secuencia
					self.clasesPorDia[i][2] = self.hoja.row(rx)[1].value #Materia
					self.clasesPorDia[i][3] = self.hoja.row(rx)[dia].value #horario
					self.clasesPorDia[i][4] = 0 #Asignado si o no
					self.clasesPorDia[i][5] = 'x' #Salon asignado
					self.clasesPorDia[i][6] = (int(self.clasesPorDia[i][3].split('-')[0].split(':')[0]) - 7) * 2 + int(self.clasesPorDia[i][3].split('-')[0].split(':')[1])/30 #Periodo incial
					self.clasesPorDia[i][7] = (int(self.clasesPorDia[i][3].split('-')[1].split(':')[0]) - 7) * 2 + int(self.clasesPorDia[i][3].split('-')[1].split(':')[1])/30 #Periodo final
					self.clasesPorDia[i][8] = self.clasesPorDia[i][7] - self.clasesPorDia[i][6] #Numero de periodos 
					i = i +1
				sec_temp = self.hoja.row(rx)[0].value
				mat_temp = self.hoja.row(rx)[1].value
			else:
				if mat_temp != self.hoja.row(rx)[1].value:
					if not self.hoja.row(rx)[dia].value == '':
						self.clasesPorDia[i][0] = i
						self.clasesPorDia[i][1] = self.hoja.row(rx)[0].value
						self.clasesPorDia[i][2] = self.hoja.row(rx)[1].value
						self.clasesPorDia[i][3] = self.hoja.row(rx)[dia].value
						self.clasesPorDia[i][4] = 0
						self.clasesPorDia[i][5] = 'x'
						self.clasesPorDia[i][6] = (int(self.clasesPorDia[i][3].split('-')[0].split(':')[0]) - 7) * 2 + int(self.clasesPorDia[i][3].split('-')[0].split(':')[1])/30 #Periodo incial
						self.clasesPorDia[i][7] = (int(self.clasesPorDia[i][3].split('-')[1].split(':')[0]) -7 ) * 2 + int(self.clasesPorDia[i][3].split('-')[1].split(':')[1])/30 #Periodo final
						self.clasesPorDia[i][8] = self.clasesPorDia[i][7] - self.clasesPorDia[i][6] #Numero de periodos
						i = i +1
					mat_temp = self.hoja.row(rx)[1].value
		return i
		
	def creaHorario(self):
		self.setExcelBook()
		print 'Libro cargado correctamente'
		self.setMatrix()
		print 'Matrices creadas correctamente'
		for i in range(5):
			self.record = self.classRecord(i + 6)
			print 'Dia ' + str(i) + ' leido correctamente'
		for 
			
"""
class Horario:
	def __init__(self, libro):
		self.libro = xlrd.open_workbook(libro)
		self.hoja = libro.self.hojaeet_by_index(0)
"""
ventana =  Ventana()
ventana.botones()
ventana.inicio()
