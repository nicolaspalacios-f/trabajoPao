import sys
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from datetime import date 
from tkinter import *
from tkinter import filedialog  
import jinja2
import pdfkit



filename = ""
def generatepdf(contexto, filename):
    context = {'html': contexto}
    loader = jinja2.FileSystemLoader('./')
    enviroment = jinja2.Environment(loader=loader)
    template = enviroment.get_template('plantilla.html')
    output = template.render(context)
    config = pdfkit.configuration(wkhtmltopdf="C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")
    pdfkit.from_string(output, str(filename) + '.pdf', configuration=config)

def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Elegir Archivo",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*")))
      
    return filename

def generateFile(filename):
    file = pd.read_excel(filename)
    file = file.fillna(0)
    file = file.drop(file.columns[0],axis=1)
    file['Fecha de agendamiento'] = pd.to_datetime(file['Fecha de agendamiento'])
    file['Fecha de agendamiento'] = file['Fecha de agendamiento'].dt.date
    for i in range(len(file.index)):
        atributoPersona = {}
        for cliente in file:
            atributo = (file[cliente][i])  
            if  type(atributo) == date:
                atributo = str(atributo)
            else:
                atributo = (file[cliente][i])    
            if atributo == 0:
                pass
            else:
                atributoPersona[cliente] = atributo
        contexto = htmlGenerator(atributoPersona)
        generatepdf(contexto[0], contexto[1])
                
    
def htmlGenerator(datos): 
    html = ""
    nombre = list(datos.values())[2]
    for key in datos:
        if "," in str(datos[key]):
            html += "<tr><br><td>" + str(key) + "</td><td>"
            split= datos[key].split(",")
            for element in split:
                html += element + "<br>"
            html += "</td></tr>"
        else:
            html += "<tr><td>" + str(key) + "</td><td>" + str(datos[key]) + "</td></tr>"
    rta = [html, nombre]
    return(rta)
            ##contexto[key] = atributo
        
class My_Window(QMainWindow):
    filename = ""
    def getFilename(self):
        return self.filename
    def setFilename(self, filename):
        self.filename = filename
    def __init__(self):
        super(My_Window, self).__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(1200,200,400,400)
        self.setWindowTitle("Migracion")

        self.lbl_Titulo = QtWidgets.QLabel(self)
        self.lbl_Titulo.setText("Elija la ubicacion del archivo Forms")
        self.lbl_Titulo.adjustSize()
        self.lbl_Titulo.move(50,50)

        self.btn_location = QtWidgets.QPushButton(self)
        self.btn_location.setText("Ubicacion")
        self.btn_location.move(250,42)
        self.btn_location.clicked.connect(self.btn_location_clicked)
        
        self.btn_generate = QtWidgets.QPushButton(self)
        self.btn_generate.setText("Generar")
        self.btn_generate.move(250,82)
        self.btn_generate.clicked.connect(self.btn_generate_clicked)
        
        
    def btn_location_clicked(self):
        self.setFilename(browseFiles())

    def btn_generate_clicked(self):
        filename = self.getFilename()
        filename = "C:\\Users\\nicolas\\Desktop\\Trabajo\\Pao\\prueba forms.xlsx"
        if filename != "":
            
            generateFile(filename)
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Seleccione un archivo')
            msg.setWindowTitle("Error")
            msg.exec_()

def window():
    app = QApplication(sys.argv)
    win = My_Window()
    win.show()
    sys.exit(app.exec_())
    
window()
