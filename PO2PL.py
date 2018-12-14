from PyQt5.QtWidgets import (QMainWindow, QTextEdit, 
    QAction, QFileDialog, QApplication)
from PyQt5.QtGui import QIcon
import sys
import xlrd
import xlwt

class Example(QMainWindow):
    
    def __init__(self):
        super().__init__()
        
        # display window
        self.initUI()
        
        # input variables
        self.inWorkSheets = []
        self.defSheetName = 'Sheet1'
        self.defIdxs = [2, 3, 4]
        self.defHeader = ['Item', 'Quantity', 'Unit', 'Description', 'Received by']
        
        # output variables
        self.aggList = []
        self.outWB = xlwt.Workbook()
        self.outWS = self.outWB.add_sheet(self.defSheetName)
        
    
    '''
    Function for opening excel file and worksheet to be readable by the program
    '''
    def openXL(self, file):
        wb = xlrd.open_workbook(file)
        ws = wb.sheet_by_name(self.defSheetName)
        
        return ws
        
    
    '''
    Function for reading the data in each sheet
    '''
    def readData(self, tSheet):
        i = 0
        tempSheetData = []
        
        poNumber = tSheet.cell_value(3,6) + ' ' + str(int(tSheet.cell_value(3,5)))
        poNumber = poNumber.upper()
        
        # append PO number/ Ref number F4~G4 of PO files to tempSheetData.
        # always set as 0th element followed by items purchased
        tempSheetData.append(poNumber)
        
        
        # this append the items purchased to tempSheetData starting at 1th element
        for row in range(14,35):
            if tSheet.row_values(row)[2] == "":
                break
            
            j = 0
            temp = []
            for x in tSheet.row_values(row):
                if j in self.defIdxs:
                    if isinstance(x, float):
                        x = int(x)
                    temp.append(x)
                j += 1
            
            tempSheetData.append(temp)
            i += 1
            
        self.aggList.append(tempSheetData)
    
    
    '''
    Function for creating an excel file
    '''
    def createFile(self):
        wb_out = xlwt.Workbook()
        ws_out = wb_out.add_sheet(self.defSheetName)

        return wb_out, ws_out
    
    
    '''
    Function for updating the data inside the excel file
    '''
    def updateFile(self):
        colCnt = 0
        
        # Header is printed for the file
        for name in self.defHeader:
            self.outWS.write(0, colCnt, name)
            colCnt += 1
            
            
        rowCnt = 1
        itemCnt = 1
        for sheet in self.aggList:
            for row in sheet:
            # PO number is printed
                if not isinstance(row, list):
                    self.outWS.write(rowCnt, 3, row)

                    rowCnt += 1
                    
                else:
                    itemColCnt = 1
                    
                    self.outWS.write(rowCnt, 0, itemCnt)
                    for data in row:
                        self.outWS.write(rowCnt, itemColCnt, data)

                        itemColCnt += 1

                    itemCnt += 1
                    rowCnt += 1
            
            
        self.outWB.save('Sample_PO/pl/packing_list001.xls')
        
    
    '''
    initialize display of graphical user interface
    simple window and a few simple functions
    '''
    def initUI(self):      

        self.textEdit = QTextEdit()
        self.setCentralWidget(self.textEdit)
        self.statusBar()

        openFile = QAction(QIcon('open.png'), 'Open', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Open new File')
        openFile.triggered.connect(self.showDialog)
        
        exitButton = QAction(QIcon('exit24.png'), 'Exit', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip('Exit application')
        exitButton.triggered.connect(self.close)
        
        createButton = QAction(QIcon(), 'Create', self)
        createButton.setShortcut('Ctrl+K')
        createButton.setStatusTip('Create Packing List')
        createButton.triggered.connect(self.updateFile)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(openFile)
        fileMenu.addAction(createButton)
        fileMenu.addAction(exitButton)
        
        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('File dialog')
        self.show()
        
        
    
    '''
    Dialog box for opening excel files
    '''
    def showDialog(self):

        fname = QFileDialog.getOpenFileNames(self, 'Open file', '/home')
        
        if fname[0]:
            i = 1
            for x in fname[0]:
                if i == 1:
                    tempSheet = self.openXL(x)
                    self.readData(tempSheet)
                    self.textEdit.setText(x)
                else:
                    tempSheet = self.openXL(x)
                    self.readData(tempSheet)
                    self.textEdit.append(x)
                    
                i += 1


        
if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())