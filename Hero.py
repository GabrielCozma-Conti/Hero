# -*- coding: utf-8 -*-
"""
Created on Wed Mar  4 16:24:47 2020

@author: uib73024
"""

from PyQt5 import QtWidgets,QtCore
from PyQt5.uic import loadUi
import sys
import xls_read
import cc_modify


class Hero(QtWidgets.QMainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        
        loadUi("uir.ui", self)
        
        self.setWindowTitle("HERO")
        
        self.browseButton.clicked.connect(self.BrowseFunction)
        self.openButton.clicked.connect(self.OpenFunction)
        self.exitButton.clicked.connect(self.ExitFunction)
        self.commitCodeChangeButton.clicked.connect(self.CommitCodeChangeFunction)
        self.codeChangeComboBox.currentIndexChanged.connect(lambda: self.IndexChangeFunction)
        self.homeButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.nextSelectXlsButton.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        
        self.fileName = 0
        
    def IndexChangeFunction(self):
        self.index_code_change = self.codeChangeComboBox.currentIndex()
    def BrowseFunction(self):
        self.fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open XLS",
               (QtCore.QDir.homePath()), "XLS (*.xlsx *.xlsm)") 
        self.pathLineEdit.setText(self.fileName)
        
    def OpenFunction(self):
        if self.fileName:
            self.file = xls_read.ModuleWorkbook(self.fileName)
            self.selectXlsLabel.setText('FISIERUL A FOST DESCHIS CU SUCCES')
            self.nextSelectXlsButton.setEnabled(True)
            self.names = self.file.GetAllSheetNames()
            self.codeChangeComboBox.addItems(self.names)
        else:
            self.selectXlsLabel.setText('EROARE LA DESCHIDEREA FISIERULUI')
            self.nextSelectXlsButton.setEnabled(False)
         
    def CommitCodeChangeFunction(self):
        print(self.codeChangeComboBox.currentIndex())
        name = self.names[self.codeChangeComboBox.currentIndex()]
        
        print(str(name))
        self.file.SetSheet(name)
        max = self.file.GetMaxRowNumber()
        for i in range(2,max):
            CC = cc_modify.cc_modify(int(i), self.file)
        
    def ChangeCurrentIndexFunction(index, self):
        return lambda: self.stackedWidget.setCurrentIndex(index)
        
    def ExitFunction(self):
        return sys.exit(0)
    
    
def main():
    app = QtWidgets.QApplication(sys.argv)
    window = Hero()
    window.show()
    app.exec_()
    
if __name__ == '__main__':
    main()