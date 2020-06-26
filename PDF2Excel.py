import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from os import listdir
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton
import sys

class MyWindow(QMainWindow):
  def __init__(self):
      super(MyWindow,self).__init__()
      self.initUI()

  def button_clicked(self):
      self.ConvertPDFtoExcel()

  def initUI(self):
      self.setFixedSize(600, 280)
      self.setWindowTitle("PDF2Excel")
      icon = QtGui.QIcon()
      icon.addPixmap(QtGui.QPixmap("accoding.jpg"), QtGui.QIcon.Selected, QtGui.QIcon.On)
      # self.setWindowIcon(QtGui.QIcon('accoding.jpg'))
      self.setWindowIcon(icon)
      self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)

      self.lblInstruction = QtWidgets.QLabel(self)
      self.lblInstruction.setText("1. The program will extract only tables from PDF files, in other words, it'll ignore text paragraphs.")
      self.lblInstruction.move(50,0)
      self.lblInstruction.resize(500, 100)
      self.lblInstruction.setWordWrap(True)

      self.lblHyperlink = QtWidgets.QLabel(self)
      self.lblHyperlink.setText("2. Use <a href=''>VBA Text2Column</a> to coerce string into general format for the produced Excel file.")
      self.lblHyperlink.move(50,40)
      self.lblHyperlink.resize(500, 100)
      self.lblHyperlink.setOpenExternalLinks(True)
      self.lblHyperlink.setWordWrap(True)

      self.lblAuthor = QtWidgets.QLabel(self)
      self.lblAuthor.setText("Created by Mike Y. Liao [<a href='mailto:n9102125@gmail.com'>n9102125@gmail.com</a>]")
      self.lblAuthor.move(120,200)
      self.lblAuthor.resize(600, 100)
      self.lblAuthor.setOpenExternalLinks(True)

      self.btnConvert = QPushButton(self)
      self.btnConvert.setText("Select a folder where PDF files located")
      self.btnConvert.move(50, 120)
      self.btnConvert.resize(500,100)
      self.btnConvert.clicked.connect(self.button_clicked)
      

  def ConvertPDFtoExcel(self):
    Main(self)

def window():
  app = QApplication(sys.argv)
  win = MyWindow()
  win.show()
  sys.exit(app.exec_())

def showdialog(self, message):
  return QMessageBox.information(self,'Info', message, QMessageBox.Ok | QMessageBox.Cancel)

def ShowInfoDialog(self, message):
  msgBox = QMessageBox(self)
  msgBox.setIcon(QMessageBox.Information)
  msgBox.setText(message)
  # msgBox.setInformativeText(message)
  msgBox.setStandardButtons(QMessageBox.Ok)
  msgBox.exec_()

def HideTkWindow():
  root = tk.Tk()
  root.withdraw()

def GetDirPath():
  return filedialog.askdirectory()

def Main(self):
  HideTkWindow()
  dirPath = GetDirPath()
  try:
    if dirPath:
      isAnyPDF = PDFsToExcels(self, dirPath)
      if isAnyPDF:
        reply = showdialog(self, 'The PDF files have been successfully converted! Would you like to open the directory where files located?')
        if reply == QMessageBox.Ok:
          dirPath = dirPath.replace('/', '\\')
          os.startfile(dirPath)
      else:
        ShowInfoDialog(self, 'There is no PDF file.')
  except:
    showdialog(self, 'Something wrong happened!\nClose Excel files and try again.')
    
def PDFsToExcels(self, dirPath):
  filesFullPath = list_filesFullPath(dirPath, 'pdf')
  if len(filesFullPath) > 0:
    for f in filesFullPath:
      PDFToExcel(f)
    return True
  else:
    return False

def PDFToExcel(filesPath):
  with pdfplumber.open(filesPath) as pdf:
      writer = pd.ExcelWriter(f'{filesPath.replace(".pdf", "")}.xlsx', engine='xlsxwriter')
      i = 1
      for page in pdf.pages:
        tables = page.find_tables()        
        if len(tables) > 0:
          j = 0
          for j in range(len(tables)):
            tb = tables[j].extract()
            df = pd.DataFrame(tb[1:], columns=tb[0])
            sheetName = f'Sheet{i}' if len(tables) == 1 else f'Sheet{i}_{j + 1}'
            df.to_excel(writer, sheet_name=sheetName, index=False)
            j += 1
        i += 1
      writer.save()

def list_files(directory, extension):
  return [f for f in listdir(directory) if f.endswith('.' + extension)]

def list_filesFullPath(directory, extension):  
  files = list_files(directory, extension)
  return [f'{directory}/{f}' for f in files]

if __name__ == '__main__':
  window()