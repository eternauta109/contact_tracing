
import datetime as dt
import sys
import PyQt5
from PyQt5 import uic
from PyQt5 import QtWidgets
from pandas._config.config import options
""" from pyzbar.pyzbar import decode """
from PyQt5.QtWidgets import  QTableWidgetItem,QMainWindow, QApplication, QMessageBox, QHeaderView, QWidget
""" from usb_scanner.scanner import barcode_reader """
import os
import sqlite3
import pandas as pd
import xlsxwriter

def resource_path(relative_path):
   base_path=getattr(sys,'_MEPASS',os.path.dirname(os.path.abspath(__file__)))
   return os.path.join(base_path,relative_path)

#carico il database o ne creo uno se non esiste
ROOT_DIR = os.path.dirname(os.path.abspath(__file__)) 
DB_PATH = os.path.join(ROOT_DIR, 'tracing.db')
GUI_PATH=resource_path("tracingGUI.ui")
QRY_PATH=resource_path("exportToExcell.ui")

print (ROOT_DIR)
print(DB_PATH)
print('ui',GUI_PATH)
print('query', QRY_PATH)

"""Ui_MainWindow, QtBaseClass=uic.loadUiType(GUI_PATH)"""

tracingDB=sqlite3.connect(DB_PATH)
tracingDB.execute("""CREATE TABLE IF NOT EXISTS tracing(
    id integer PRIMARY KEY,
    daytime date,
    time date,
    codfisc text,
    ticket text,
    numIngressi integer
    );""")

class queryUi(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(QRY_PATH,self)
        
        self.pushButton.clicked.connect(self.toexcell)
        self.tablewdg.setColumnCount(6)

    def toexcell(self):
        print("to excell")
        fromTime=self.fromTime.date().toPyDate()
        toTime=self.toTime.date().toPyDate()
        start_day=fromTime.strftime("%d-%m-%y")
        end_day=toTime.strftime("%d-%m-%y")
        start_hour=self.fromTime.time().toString()  
        start_hour=start_hour[0:5]    
        end_hour=self.toTime.time().toString()
        end_hour=end_hour[0:5]
        
        print(start_hour,end_hour,start_day,end_day)
        
        query=f"SELECT * FROM tracing WHERE daytime BETWEEN ? and ? and time BETWEEN ? and ?"
        result=tracingDB.execute(query,(start_day,end_day,start_hour,end_hour))  
        self.tablewdg.setRowCount(0)
        print(query)
        for row_number, row_data in enumerate(result):   
            print(row_number)  
            print(row_data)     
            self.tablewdg.insertRow(row_number)
            for column_number, data in enumerate(row_data): 
                print(data)   
                print('ll',enumerate(row_data))       
                self.tablewdg.setItem(row_number,column_number, QTableWidgetItem(str(data)))
        
        
        return
        with pd.ExcelWriter("Output.xlsx", engine="xlsxwriter", options={'strings_to_numbers':True, "string_to_formulas":False}) as writer:
            try:
                df=pd.read_sql("Select * from tracing",tracingDB)
                df.to_excel(writer, sheet_name="db contact", header=True, index=False)
                print('operazione riuscita')
            except:
                print("c'è un problema")
                
                
class ui(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(GUI_PATH,self)
        self.btnremoveitem.clicked.connect(self.removeitem)
        self.acceptButton.clicked.connect(self.addItem)
        self.expToExcell.clicked.connect(self.queryUi)
        
        header=self.tablewdg.horizontalHeader()
        print(header)
        header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents )
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch )
        self.tablewdg.setColumnCount(6)
        """ self.tablewdg.setHorizontalHeaderLabels(['id', 'data',time, 'codice fiscale','ticket','numIngressi'])    """   
        self.tablewdg.setColumnHidden(3,True)
        self.tablewdg.setColumnHidden(0,True)
        
        """ self.setTabOrder(self.acceptButton,self.codFiscInput) """
       
        self.load()     

    def queryUi(self):
        self.w=queryUi()
        self.w.show()    
        
    def closeEvent(self,event):
        tracingDB.close()    
        
    def removeitem(self):
        if self.tablewdg.selectedItems():
            rawindex =  self.tablewdg.currentRow()
            
            dbindex = int(self.tablewdg.item(rawindex,0).text())
            ##non so bene che fa
            #a=len(set(index.row() for index in self.tablewdg.selectedIndexes()))

            ############################################
            ## le line succ entrano nell-item         ##
            ##field1 = self.tbl_anggota.item(r,0).text()
            ##field2 = self.tbl_anggota.item(r,1).text()
            ############################################
            query=f"DELETE FROM tracing WHERE id={dbindex}"
            tracingDB.execute(query)
            tracingDB.commit()
            self.load()
            """ print(rawindex) """
    
    def addItem(self):        
        
        day=dt.date.today()
        time=dt.datetime.now()
        current_day=day.strftime("%d-%m-%y")
        current_time=time.strftime("%H:%M")

        print(current_day,current_time)
        if (self.codFiscInput.text()=="")\
        or (self.ticketInput.text()==""):
           QMessageBox.about(self, "ma che stai a fa", "hai lasciato qualche campo vuoto!")
           return      
        """ detectedBarcodes = barcode_reader(self.ticketInput.text) """
        """ print(detectedBarcodes) """
        lista=(current_day,current_time, self.codFiscInput.text(),self.ticketInput.text(), self.inputIngressi.text())
        
        tracingDB.execute("INSERT INTO tracing (daytime,time,codfisc,ticket,numIngressi)\
             VALUES (?,?,?,?,?) ",lista)
        tracingDB.commit()
       
        self.codFiscInput.setText("")
        self.ticketInput.setText("") 
        self.inputIngressi.setText("1")       
        self.load()
            
    def load(self):         
        query="SELECT * FROM tracing ORDER BY id DESC"
        result=tracingDB.execute(query)  
        """ demos=enumerate(result) 
        print(demos)   """   
        self.tablewdg.setRowCount(0)
        for row_number, row_data in enumerate(result):   
            """ print(row_number)  
            print(row_data)  """      
            self.tablewdg.insertRow(row_number)
            for column_number, data in enumerate(row_data): 
                """ print(data)   
                print('ll',enumerate(row_data))  """           
                self.tablewdg.setItem(row_number,column_number, QTableWidgetItem(str(data)))
        
def main():
    app=QApplication([])
    window=ui()  
    window.show()
    app.exec()
        
main()