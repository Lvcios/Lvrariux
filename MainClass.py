# _*_coding:utf8_*_
#librerias a instalar 
#xlrd, xlwt, xlutils
import xlrd
import xlwt
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
		self.penalizaciones = 0
		self.totalPenalizaciones = 0
		self.wb = xlwt.Workbook()
	
	def inicio(self):
		self.root.mainloop()
	
	def botones(self):
		btnCarga = Button(self.root, text = 'Cargar archivo', command = self.cargarArchivo , width = 30)
		btnCrea = Button(self.root, text = 'Crear Horario',command = self.creaHorario, width = 30)
		btnMuestra = Button(self.root, text = 'Muestra Horario',command = self.muestraHorario, width = 30)
		btnMuestraClasesPorDia = Button(self.root, text = 'Muestra db',command = self.muestraClasesPorDia, width = 30)
		btnCarga.grid(row = 1, column = 0)
		btnCrea.grid(row = 3, column = 0)
		btnMuestra.grid(row = 4, column = 0)
		btnMuestraClasesPorDia.grid(row = 5, column = 0)
	
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
		
	
	def verifyClassTimeSlot(self,dia):
		pr, i, penalizaciones = 0, 1, 0
		while (i < self.record):
			if pr == self.clasesPorDia[i][6] and self.clasesPorDia[i][5] == 'x':
				retorno = self.addClass(self.clasesPorDia[i][6],self.clasesPorDia[i][7],self.clasesPorDia[i][0],dia)
				pr = retorno[1]
				penalizaciones = retorno[0] + penalizaciones
				i = 0
			else:
				i = i + 1
		#print 'Penalizaciones: ' + str(penalizaciones)
	
	def addClass(self, pi,pf,idClase,dia):
		penalizacion = 0
		permiso = False
		salon = 0
		while(permiso == False):
			for i in range(pi,pf):
				if self.Salones[dia][salon][i] == 0:
					permiso = True
				else:
					permiso = False
					salon = salon + 1
					break
		if permiso:
			for i in range(pi,pf):
				self.Salones[dia][salon][i] = idClase	
			if (salon >= 11):
				penalizacion = penalizacion + 1
			self.clasesPorDia[idClase][4] = 1
			self.clasesPorDia[idClase][5] = salon
		return penalizacion, self.clasesPorDia[idClase][7]
	
	def creaHorario(self):
		self.setExcelBook()
		print 'Libro cargado correctamente'
		self.setMatrix()
		print 'Matrices creadas correctamente'
		#creacion de nuevo libro excel:
		ws = self.wb.add_sheet('Test',cell_overwrite_ok=True)
		for dia in range(5):
			self.record = self.classRecord(dia + 6)
			for i in range(20):
				self.verifyClassTimeSlot(dia)
			for i in range(len(self.clasesPorDia)):
				if self.clasesPorDia[i][5] == 'x':
					self.addClass(self.clasesPorDia[i][6],self.clasesPorDia[i][7],self.clasesPorDia[i][0],dia)
			for i in range(len(self.clasesPorDia)):
				if not self.clasesPorDia[i][0] == 0:
					#print self.clasesPorDia[i]
					ws.write(i,0,self.clasesPorDia[i][0])
					ws.write(i,1,self.clasesPorDia[i][1])
					ws.write(i,2,self.clasesPorDia[i][2])
					ws.write(i,3,self.clasesPorDia[i][3])
					ws.write(i,4,self.clasesPorDia[i][4])
					ws.write(i,5,self.clasesPorDia[i][5])
					ws.write(i,6,self.clasesPorDia[i][6])
					ws.write(i,7,self.clasesPorDia[i][7])
					ws.write(i,8,self.clasesPorDia[i][8])
					#ws.write(0, 0, 'Test', style0)
					if self.clasesPorDia[i][5] >=11 :
						self.penalizaciones = self.penalizaciones + 1
			self.totalPenalizaciones  = self.totalPenalizaciones  + self.penalizaciones
			print self.penalizaciones
			self.wb.save('example.xls')
		
	def muestraHorario(self):
		for i in range(len(self.Salones)):
			print 'Dia ' + str(i)
			for j in range(len(self.Salones[i])):
				print 'Salon ' + str(j)
				print self.Salones[i][j]
				#raw_input()
	
	def muestraClasesPorDia(self):
		for rx in range(1, self.hoja.nrows):
			print self.clasesPorDia[rx]

ventana =  Ventana()
ventana.botones()
ventana.inicio()
