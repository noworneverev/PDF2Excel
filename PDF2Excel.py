import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from os import listdir
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QPushButton, QComboBox, QLabel, QGridLayout, QSizePolicy, QWidget
import sys

class MyWindow(QMainWindow):
  def __init__(self):
      super(MyWindow,self).__init__()
      self.initUI()

  def button_clicked(self):
      self.ConvertPDFtoExcel()

  def initUI(self):
      self.setFixedSize(600, 350)
      self.setWindowTitle("PDF2Excel")
      icon = QtGui.QIcon()
      icon.addPixmap(QtGui.QPixmap("accoding.jpg"), QtGui.QIcon.Selected, QtGui.QIcon.On)
      # self.setWindowIcon(QtGui.QIcon('accoding.jpg'))
      self.setWindowIcon(icon)
      self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)

      wid = QWidget(self)
      self.setCentralWidget(wid)
      layout = QGridLayout()
      wid.setLayout(layout)

      self.lblInstruction = QLabel(self)
      self.lblInstruction.setText("1. The program will extract only tables from PDF files, in other words, it'll ignore text paragraphs.")
      self.lblInstruction.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
      self.lblInstruction.setWordWrap(True)
      
      self.lblHyperlink = QLabel(self)
      self.lblHyperlink.setText("2. Use <a href='https://github.com/noworneverev/PDF2Excel/releases/download/1.0.0/Text2Column.xlam'>VBA Text2Column</a> to coerce string into general format for the produced Excel file.")
      self.lblHyperlink.setOpenExternalLinks(True)
      self.lblHyperlink.setWordWrap(True)

      self.lblStrategy = QLabel(self)
      myFont=QtGui.QFont()
      myFont.setBold(True)
      self.lblStrategy.setFont(myFont)
      self.lblStrategy.setText("Table-extraction strategy:")

      self.lblVerticalStrategy = QLabel(self)
      self.lblVerticalStrategy.setText("    Vertical Strategy")

      self.lblHorizontalStrategy = QLabel(self)
      self.lblHorizontalStrategy.setText("    Horizontal Strategy")
      self.lblHorizontalStrategy.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)

      self.cboVerticalStrategy = QComboBox(self)
      self.cboVerticalStrategy.addItems(["lines", "lines_strict", "text"])
      
      self.cboHorizontalStrategy = QComboBox(self)
      self.cboHorizontalStrategy.addItems(["lines", "lines_strict", "text"])
      
      self.btnConvert = QPushButton(self)
      self.btnConvert.setText("Select a folder where PDF files located")
      self.btnConvert.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
      self.btnConvert.clicked.connect(self.button_clicked)

      self.lblAuthor = QLabel(self)
      self.lblAuthor.setText("Created by Mike Y. Liao [<a href='mailto:n9102125@gmail.com'>n9102125@gmail.com</a>]")
      self.lblAuthor.setAlignment(QtCore.Qt.AlignCenter)
      self.lblAuthor.setOpenExternalLinks(True)

      layout.addWidget(self.lblInstruction,0,0,1,0,QtCore.Qt.AlignVCenter)
      layout.addWidget(self.lblHyperlink,1,0,1,0,QtCore.Qt.AlignVCenter)
      layout.addWidget(self.lblStrategy,3,0)
      layout.addWidget(self.lblVerticalStrategy,4,0)
      layout.addWidget(self.lblHorizontalStrategy,5,0)
      layout.addWidget(self.cboVerticalStrategy,4,1)
      layout.addWidget(self.cboHorizontalStrategy,5,1)
      layout.addWidget(self.btnConvert,6,0,1,0)
      layout.addWidget(self.lblAuthor,8,0,1,0)

      layout.setContentsMargins(20,20,20,20)
      layout.setRowMinimumHeight(2,10)
      

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
      PDFToExcel(self, f)
    return True
  else:
    return False

def PDFToExcel(self, filesPath):
  with pdfplumber.open(filesPath) as pdf:
      writer = pd.ExcelWriter(f'{filesPath.replace(".pdf", "")}.xlsx', engine='xlsxwriter')
      i = 1
      table_settings = {
        "vertical_strategy": self.cboVerticalStrategy.currentText(),
        "horizontal_strategy": self.cboHorizontalStrategy.currentText()
      }
      for page in pdf.pages:
        # tables = page.find_tables()     
        tables = page.find_tables(table_settings)     
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