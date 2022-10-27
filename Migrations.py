import pandas
import openpyxl
import numpy
import plotly
import plotly.graph_objs as go
import dash
from dash import html
from dash import dcc
import sqlite3
from PyQt5.QtWidgets import QWidget, QMessageBox, QApplication,QPushButton ,QFileDialog,QLabel
import sys


class PathWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.fname = ""

        self.setGeometry(300, 300, 550, 150)
        self.setWindowTitle('Миграция населения республики Татарстан')

        self.lbl = QLabel(self)
        self.lbl.resize(450, 20)
        self.lbl.move(60, 40)
        self.lbl.setStyleSheet("border: 1px solid black;")

        btn = QPushButton('Путь', self)
        btn.move(60, 60)
        btn.clicked.connect(self.showDialog)

        btn = QPushButton('Выполнить', self)
        btn.move(415, 60)
        btn.clicked.connect(self.run)

        self.show()

    def showDialog(self):
        self.fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')[0]
        self.lbl.setText(self.fname)

    def run(self):

        tuplePT = self.pt()
        tupleM = self.m()
        tupleN = self.n()
        tupleEP = self.ep(tupleM, tupleN)
        tupleOP = self.op(tuplePT)
        tupleMP = self.mp(tupleOP, tupleEP)
        tupleKOP = self.kop(tupleOP)
        tupleKMP = self.kmp(tupleMP, tupleOP)

        self.drawGraph(tupleOP, tupleMP)

        self.safeInExcel(tupleEP, 'ЕП')
        self.safeInExcel(tupleOP, 'ОП')
        self.safeInExcel(tupleMP, 'МП')
        self.safeInExcel(tupleKOP, 'Коп')
        self.safeInExcel(tupleKMP, 'Кмп')

        self.createTableDB(tupleOP)

    def closeEvent(self, event):

        reply = QMessageBox.question(self, 'Message',
            "Уверены, что хотите выйти?", QMessageBox.Yes |
            QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


#pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt pt

    def pt(self):
        excel_data_Pt = pandas.read_excel(self.fname, sheet_name='Pt', header=2)
        listPT = []

        for value in excel_data_Pt.values:
            dictPT = {}
            iterPT = 0
            for column in excel_data_Pt.columns:
                dictPT[column] = value[iterPT]
                iterPT = iterPT + 1
            listPT.append(dictPT)

        tuplePT = tuple(listPT)
        return tuplePT

# m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m m
    def m(self):
        excel_data_M = pandas.read_excel(self.fname, sheet_name='M', header=2)
        listM = []

        for value in excel_data_M.values:
            dictM = {}
            iterM = 0
            for column in excel_data_M.columns:
                dictM[column] = value[iterM]
                iterM = iterM + 1
            listM.append(dictM)

        tupleM = tuple(listM)
        return tupleM

#n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n n
    def n(self):
        excel_data_N = pandas.read_excel(self.fname, sheet_name='N', header=2)
        listN = []

        for value in excel_data_N.values:
            dictN = {}
            iterN = 0
            for column in excel_data_N.columns:
                dictN[column] = value[iterN]
                iterN = iterN + 1
            listN.append(dictN)

        tupleN = tuple(listN)
        return tupleN

#EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP EP
    def ep(self,tupleM,tupleN):
        listEP = []

        for n in tupleN:
            dictEP = {}
            Mdict = list(filter(lambda item: item['МО'] == n['МО'], tupleM))
            dictEP['МО'] = n['МО']
            for keys in n.keys():
                if keys != 'МО':
                    dictEP[keys] = n[keys] - Mdict[0][keys]
            listEP.append(dictEP)

        tupleEP = tuple(listEP)
        return tupleEP

#OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP OP
    def op(self,tuplePT):
        listOP = []

        for pt in tuplePT:
            dictOP = {}
            dictOP['МО'] = pt['МО']
            countOfColumns = len(pt)
            counter = 0
            for keys in pt:
                if keys != 'МО' and counter<countOfColumns-1:
                    dictOP[keys] = pt[keys+1] - pt[keys]
                counter+=1
            listOP.append(dictOP)

        tupleOP = tuple(listOP)
        return tupleOP

#MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP MP
    def mp(self, tupleOP,tupleEP):
        listMP = []

        for op in tupleOP:
            dictMP = {}
            MPdict = list(filter(lambda item: item['МО'] == op['МО'], tupleEP))
            dictMP['МО'] = op['МО']
            for keys in op.keys():
                if keys != 'МО':
                    dictMP[keys] = op[keys] - MPdict[0][keys]
            listMP.append(dictMP)

        tupleMP = tuple(listMP)
        return tupleMP

    # KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP KOP
    def kop(self,tupleOP):

        listKOP = []

        for op in tupleOP:
            dictKOP = {}
            dictKOP['МО'] = op['МО']
            for keys in op.keys():
                if keys != 'МО':
                    dictKOP[keys] = op[keys] / (op[keys] / 2)
            listKOP.append(dictKOP)

        tupleKOP = tuple(listKOP)
        return tupleKOP

 # KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP KMP
    def kmp(self, tupleMP, tupleOP):

        listKMP = []

        for mp in tupleMP:
            dictKMP = {}
            OPdict = list(filter(lambda item: item['МО'] == mp['МО'], tupleOP))
            dictKMP['МО'] = mp['МО']
            for keys in mp.keys():
                if keys != 'МО':
                    dictKMP[keys] = mp[keys] / (OPdict[0][keys] / 2)
            listKMP.append(dictKMP)

        tupleKMP = tuple(listKMP)
        return tupleKMP


# save save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save  save
    def safeInExcel(self,result, nameOfPage):
        wb = openpyxl.load_workbook(self.fname)

        for ir in range(0, len(result)):
            counterOfColumns = 0
            for keys in result[0].keys():
                if keys != 'МО':
                    wb[nameOfPage].cell(4 + ir, 2 + counterOfColumns).value = result[ir][keys]
                    counterOfColumns = counterOfColumns + 1

        wb.save(self.fname)

    def drawGraph(self, tupleOP, tupleMP):
        periods = []
        BarOP = {}
        if len(tupleOP) != 0:
            for key in tupleOP[0].keys():
                if key != 'МО':
                    periods.append(key)
                listYearOP = []
                for op in tupleOP:
                    listYearOP.append(op[key])
                BarOP[key] = tuple(listYearOP)

        BarMP = {}
        if len(tupleMP) != 0:
            for key in tupleMP[0].keys():
                listYearMP = []
                for op in tupleMP:
                    listYearMP.append(op[key])
                BarMP[key] = tuple(listYearMP)

        list_menu = []
        idx_period = 0
        for num, name_period in enumerate(periods):
            list_visible = [False] * len(periods) * 2

            list_visible[idx_period] = True
            idx_period = idx_period + 1
            list_visible[idx_period] = True
            idx_period = idx_period + 1
            temp_dict = dict(label=name_period, method='update', args=[{'visible': list_visible}])
            list_menu.append(temp_dict)

        fig = go.Figure()

        visible = True
        for Bmp in BarMP:
            for Bop in BarOP:
                if Bmp != 'МО' and Bmp == Bop:
                    fig.add_trace(go.Bar(x=BarMP['МО'], y=BarMP[Bmp], name='МП - миграционный прирост населения',
                                         visible=visible))
                    fig.add_trace(
                        go.Bar(x=BarOP['МО'], y=BarOP[Bop], name='ОП - общий прирост населения', visible=visible))
                    visible = False

        fig.update_layout(legend_orientation="v",
                          title="Миграция населения",
                          xaxis_title="Наименование МО",
                          yaxis_title="Численность, чел",
                          margin=dict(l=0, r=0, t=30, b=0),
                          updatemenus=list([dict(buttons=list_menu, active=0)])
                          )

        fig.show()

    def createTableDB(self, tupleOP):
        if len(tupleOP) != 0:

            columns = ""
            requestText = "CREATE TABLE OP (_id INTEGER PRIMARY KEY AUTOINCREMENT, "

            for op in tupleOP[0]:
                if op == "МО":
                    requestText = requestText + " " + str(op) + " TEXT NOT NULL, "
                    columns = columns + str(op) + ","
                else:
                    requestText = requestText + " Year" + str(op) + " INTEGER, "
                    columns = columns + "Year" + str(op) + ","

            requestText = requestText[:-2]
            requestText = requestText + ")"

            columns = columns[:-1]

            connection = sqlite3.connect('shows.db')
            cursor = connection.cursor()
            cursor.execute(requestText)
            connection.commit()

        for op in tupleOP:
            sqlite_insert_query = "INSERT INTO OP (" + columns + ")  VALUES  ("
            for key in tupleOP[0].keys():
                if isinstance(op[key], str):
                    sqlite_insert_query = sqlite_insert_query + "'" + op[key] + "',"
                else:
                    sqlite_insert_query = sqlite_insert_query + str(op[key]) + ","

            sqlite_insert_query = sqlite_insert_query[:-1] + ")"

            cursor.execute(sqlite_insert_query)
            connection.commit()

        connection.close()

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = PathWindow()
    sys.exit(app.exec_())
