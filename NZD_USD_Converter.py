import openpyxl
import requests
import math
import PyQt6
import sys
from PyQt6.QtWidgets import (
    QApplication,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QLabel
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon

from converter_backend import ExcelConverters

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('NZD-USD Converter')
        self.setWindowIcon(QIcon('Ilovemoneyjpg.jpg'))
        self.resize(400,200) #x,y

        layout= QVBoxLayout()
        self.setLayout(layout)

        labelin = QLabel("Input file path:")
        layout.addWidget(labelin)

        self.iF_inLocation = QLineEdit()
        layout.addWidget(self.iF_inLocation)

        labelout = QLabel("Save location file path:")
        layout.addWidget(labelout)

        self.iF_outLocation = QLineEdit()
        layout.addWidget(self.iF_outLocation)

        labelcol = QLabel('Column:')
        layout.addWidget(labelcol)

        self.iF_sheetColumn = QLineEdit()
        layout.addWidget(self.iF_sheetColumn)

        labelempty = QLabel('')
        layout.addWidget(labelempty)

        self.button_nz_us=QPushButton('NZ To US')#clicked=self.run_nz_us
        self.button_nz_us.clicked.connect(self.nzus)
        layout.addWidget(self.button_nz_us)

        self.button_us_nz = QPushButton('US To NZ')#clicked=self.run_us_nz
        self.button_us_nz.clicked.connect(self.usnz)
        layout.addWidget(self.button_us_nz)

    def nzus(self):
        ExcelConverters.nz_usconverter(f'{self.iF_sheetColumn.text()}', f'{self.iF_inLocation.text()}', f'{self.iF_outLocation.text()}')

    def usnz(self):
        ExcelConverters.us_nzconverter(f'{self.iF_sheetColumn.text()}', f'{self.iF_inLocation.text()}', f'{self.iF_outLocation.text()}')

    #def run_nz_us(self):
        #ExcelConverters.nz_usconverter(iF_sheetColumn, iF_inLocation, iF_outLocation)
    #def run_us_nz(self):
        #ExcelConverters.us_nzconverter(iF_sheetColumn, iF_inLocation, iF_outLocation)

app=QApplication(sys.argv)
app.setStyleSheet("""
    QWidget{
            font size: 25px;
    }
    QPushButton{
            font size: 25px;
    }
    
""")

window=MyApp()
window.show()
sys.exit(app.exec())
