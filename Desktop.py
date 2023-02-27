import sys
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox

from tkinter import *
  

from tkinter import filedialog  
filename = ""
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
    file = file.drop(file.columns[0],axis=1)
    columnas = file.columns.tolist()
    for persona in file:
        atributoPersona = {}
        for atributo in file[persona]:
            
            if atributo == "NaN":
                pass
            else:
                print(file[persona]+ ":" +atributo)
                ##atributoPersona[file[persona]] = atributo
                pass
        


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
