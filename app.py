import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QAction, QTableWidget,QTableWidgetItem,QVBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, QMessageBox, QLineEdit, QGroupBox, QVBoxLayout, QGridLayout, QHBoxLayout
from PyQt5.QtGui import QIcon, QRegExpValidator
from PyQt5.QtCore import pyqtSlot, QRegExp, Qt
import pandas as pd
import datetime
import os
from matplotlib import pyplot as plt
import seaborn as sns
import matplotlib
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from dateutil.relativedelta import relativedelta
import calendar
import xlsxwriter

class App(QWidget):
 
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def dataPath(self):
        return os.path.abspath('')+'\\data.csv'
    
    def arhivPath(self):
        now = datetime.datetime.now()
        return os.path.abspath('')+'\\arhiv.csv'
    
    def CreateFiles(self):
        dataname = 'data.csv'
        if dataname not in os.listdir(os.path.abspath('')):
            df = pd.DataFrame(columns=['ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                       'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                       'OTHERS', 'DOHOD', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'])
            df.to_csv(os.path.abspath('')+'\\'+dataname, index=False)
        arhivname = 'arhiv.csv'
        if arhivname not in os.listdir(os.path.abspath('')):
            df = pd.DataFrame(columns=['OPER_DATE', 'ID', 'VAGON', 'ADD_DATE', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                       'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                       'OTHERS', 'DOHOD', 'COEF', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'])
            df.to_csv(os.path.abspath('')+'\\'+arhivname, index=False)
    
    def Header(self):
        return ['1. ID номер', '2. Дата добавления', '3. № Вагона', '4. Станция отправления', '5. Станция назначения', '6. Дата отправления', '7. Дата прихода', '8. Расстояние  в одну сторону (км)','9. Кол-во дней в одну сторону', 
                '10. Кол-во дней в обратную сторону', '11. Кол-во дней погрузки', '12. Кол-во дней выгрузки', '13. Дата погрузки', '14. Дата выгрузки', '15. Дата входа не ремонт', '16. Дата выхода из ремонта',
                '17. Кол-во дней в ремонте', '18. Оплата аренды','19. Затраты на ремонт', '20. ППС', '21. Груженый тариф', '22. По рожней 1 территории', '23. По рожней 2 территории', '24. За ускорение 1 территории', 
                '25. За ускорение 2 территории', '26. Телеграммы', '27. Прочие расходы', '28. Приходы', '29. Кол-во дней в пути', '30. Общий расход', '31. Общий доход', '32. Сальдо']
 
    def initUI(self):
        self.CreateFiles()
        self.setGeometry(200, 200, 1500, 800)
        self.data = pd.read_csv(self.dataPath())
        self.arhiv = pd.read_csv(self.arhivPath())
        
        self.nd = NewDialog(self)
        self.nd.setGeometry(20, 20, 1460, 500)
        self.nd.tableWidget = QTableWidget()
        self.nd.tableWidget.setRowCount(len(self.data.index))
        self.nd.tableWidget.setColumnCount(len(self.data.columns))
        for i in range(len(self.data.index)):
            for j in range(len(self.data.columns)):
                self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
        self.nd.tableWidget.move(20,20)
        header = self.Header()
        self.nd.tableWidget.setHorizontalHeaderLabels(header)
        self.nd.tableWidget.resizeColumnsToContents()
        self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
        self.nd.layout = QVBoxLayout()
        self.nd.layout.addWidget(self.nd.tableWidget) 
        self.nd.setLayout(self.nd.layout)
        self.nd.show()
        
        self.group_data = QGroupBox('Данные', self)
        self.group_data.setFixedSize(400, 460)
        self.group_data.setAlignment(Qt.AlignCenter)
        self.grid_data = QGridLayout(self)
        
        self.group_oborot = QGroupBox('Доходы и расходы', self)
        self.group_oborot.setFixedSize(400, 460)
        self.group_oborot.setAlignment(Qt.AlignCenter)
        self.grid_oborot = QGridLayout(self)
        
        self.lbl_vagon = QLabel('№ Вагона:', self)
        self.lbl_vagon.setFixedSize(200, 20)
        
        self.line_vagon = QLineEdit(self)
        self.line_vagon.setFixedSize(120, 20)
        
        self.lbl_path = QLabel('Станция отправления:', self)
        self.lbl_path.setFixedSize(200, 20)
        
        self.line_path = QLineEdit(self)
        self.line_path.setFixedSize(120, 20)
                
        self.lbl_dest = QLabel('Станция назначения:', self)
        self.lbl_dest.setFixedSize(200, 20)
        
        self.line_dest = QLineEdit(self)
        self.line_dest.setFixedSize(120, 20)
        
        self.lbl_date_path = QLabel('Дата отправления:', self)
        self.lbl_date_path.setFixedSize(200, 20)
        
        self.line_date_path = QLineEdit(self)
        self.line_date_path.setFixedSize(120, 20)
        self.line_date_path.setInputMask('99.99.9999')
        
        self.lbl_date_dest = QLabel('Дата прихода:', self)
        self.lbl_date_dest.setFixedSize(200, 20)
        
        self.line_date_dest = QLineEdit(self)
        self.line_date_dest.setFixedSize(120, 20)
        self.line_date_dest.setInputMask('99.99.9999')
        
        self.lbl_km_path = QLabel('Расстояние  в одну сторону (км):', self)
        self.lbl_km_path.setFixedSize(200, 20)
        
        self.line_km_path = QLineEdit(self)
        self.line_km_path.setFixedSize(120, 20)
        self.line_km_path.setText('0')
        regexp = QRegExp('^[0-9-.]+')
        validator = QRegExpValidator(regexp)
        self.line_km_path.setValidator(validator)
        
        self.lbl_days_path = QLabel('Кол-во дней в одну сторону:', self)
        self.lbl_days_path.setFixedSize(200, 20)
        
        self.line_days_path = QLineEdit(self)
        self.line_days_path.setFixedSize(120, 20)
        self.line_days_path.setText('0')
        self.line_days_path.setValidator(validator)
        
        self.lbl_days_dest = QLabel('Кол-во дней в обратную сторону:', self)
        self.lbl_days_dest.setFixedSize(200, 20)
        
        self.line_days_dest = QLineEdit(self)
        self.line_days_dest.setFixedSize(120, 20)
        self.line_days_dest.setText('0')
        self.line_days_dest.setValidator(validator)
        
        self.lbl_days_load = QLabel('Кол-во дней погрузки:', self)
        self.lbl_days_load.setFixedSize(200, 20)
        
        self.line_days_load = QLineEdit(self)
        self.line_days_load.setFixedSize(120, 20)
        self.line_days_load.setText('0')
        self.line_days_load.setValidator(validator)
        
        self.lbl_days_unload = QLabel('Кол-во дней выгрузки:', self)
        self.lbl_days_unload.setFixedSize(200, 20)
        
        self.line_days_unload = QLineEdit(self)
        self.line_days_unload.setFixedSize(120, 20)
        self.line_days_unload.setText('0')
        self.line_days_unload.setValidator(validator)
        
        self.lbl_date_load = QLabel('Дата погрузки:', self)
        self.lbl_date_load.setFixedSize(200, 20)
        
        self.line_date_load = QLineEdit(self)
        self.line_date_load.setFixedSize(120, 20)
        self.line_date_load.setInputMask('99.99.9999')
        
        self.lbl_date_unload = QLabel('Дата выгрузки:', self)
        self.lbl_date_unload.setFixedSize(200, 20)

        self.line_date_unload = QLineEdit(self)
        self.line_date_unload.setFixedSize(120, 20)
        self.line_date_unload.setInputMask('99.99.9999')
        
        self.lbl_date_to_remont = QLabel('Дата входа на ремонт:', self)
        self.lbl_date_to_remont.setFixedSize(200, 20)
        
        self.line_date_to_remont = QLineEdit(self)
        self.line_date_to_remont.setFixedSize(120, 20)
        self.line_date_to_remont.setInputMask('99.99.9999')
        
        self.lbl_date_from_remont = QLabel('Дата выхода из ремонта:', self)
        self.lbl_date_from_remont.setFixedSize(200, 20)
        
        self.line_date_from_remont = QLineEdit(self)
        self.line_date_from_remont.setFixedSize(120, 20)
        self.line_date_from_remont.setInputMask('99.99.9999')
        
        self.lbl_days_remont = QLabel('Кол-во дней на ремонт:', self)
        self.lbl_days_remont.setFixedSize(200, 20)
        
        self.line_days_remont = QLineEdit(self)
        self.line_days_remont.setFixedSize(120, 20)
        self.line_days_remont.setValidator(validator)
        
        self.grid_data.addWidget(self.lbl_vagon, 0, 0)
        self.grid_data.addWidget(self.line_vagon, 0, 1)
        self.grid_data.addWidget(self.lbl_path, 1, 0)
        self.grid_data.addWidget(self.line_path, 1, 1)
        self.grid_data.addWidget(self.lbl_dest, 2, 0)
        self.grid_data.addWidget(self.line_dest, 2, 1)
        self.grid_data.addWidget(self.lbl_date_path, 3, 0)
        self.grid_data.addWidget(self.line_date_path, 3, 1)
        self.grid_data.addWidget(self.lbl_date_dest, 4, 0)
        self.grid_data.addWidget(self.line_date_dest, 4, 1)
        self.grid_data.addWidget(self.lbl_km_path, 5, 0)
        self.grid_data.addWidget(self.line_km_path, 5, 1)
        self.grid_data.addWidget(self.lbl_days_path, 6, 0)
        self.grid_data.addWidget(self.line_days_path, 6, 1)
        self.grid_data.addWidget(self.lbl_days_dest, 7, 0)
        self.grid_data.addWidget(self.line_days_dest, 7, 1)
        self.grid_data.addWidget(self.lbl_days_load, 8, 0)
        self.grid_data.addWidget(self.line_days_load, 8, 1)
        self.grid_data.addWidget(self.lbl_days_unload, 9, 0)
        self.grid_data.addWidget(self.line_days_unload, 9, 1)
        self.grid_data.addWidget(self.lbl_date_load, 10, 0)
        self.grid_data.addWidget(self.line_date_load, 10, 1)
        self.grid_data.addWidget(self.lbl_date_unload, 11, 0)
        self.grid_data.addWidget(self.line_date_unload, 11, 1)
        self.grid_data.addWidget(self.lbl_date_to_remont, 12, 0)
        self.grid_data.addWidget(self.line_date_to_remont, 12, 1)
        self.grid_data.addWidget(self.lbl_date_from_remont, 13, 0)
        self.grid_data.addWidget(self.line_date_from_remont, 13, 1)
        self.grid_data.addWidget(self.lbl_days_remont, 14, 0)
        self.grid_data.addWidget(self.line_days_remont, 14, 1)
        
        self.group_data.setLayout(self.grid_data)
        self.group_data.move(30, 530)
        
        self.lbl_rent = QLabel('Оплата аренды:', self)
        self.lbl_rent.setFixedSize(200, 20)
        
        self.line_rent = QLineEdit(self)
        self.line_rent.setFixedSize(120, 20)
        self.line_rent.setText('0')
        self.line_rent.setValidator(validator)
        
        self.lbl_remont = QLabel('Затраты на ремонт:', self)
        self.lbl_remont.setFixedSize(200, 20)
        
        self.line_remont = QLineEdit(self)
        self.line_remont.setFixedSize(120, 20)
        self.line_remont.setText('0')
        self.line_remont.setValidator(validator)
        
        self.lbl_pps = QLabel('ППС:', self)
        self.lbl_pps.setFixedSize(200, 20)
        
        self.line_pps = QLineEdit(self)
        self.line_pps.setFixedSize(120, 20)
        self.line_pps.setText('0')
        self.line_pps.setValidator(validator)
        
        self.lbl_gruz = QLabel('Груженый тариф:', self)
        self.lbl_gruz.setFixedSize(200, 20)
        
        self.line_gruz = QLineEdit(self)
        self.line_gruz.setFixedSize(120, 20)
        self.line_gruz.setText('0')
        self.line_gruz.setValidator(validator)
        
        self.lbl_rozn1 = QLabel('По рожней 1 территории:', self)
        self.lbl_rozn1.setFixedSize(200, 20)
        
        self.line_rozn1 = QLineEdit(self)
        self.line_rozn1.setFixedSize(120, 20)
        self.line_rozn1.setText('0')
        self.line_rozn1.setValidator(validator)
        
        self.lbl_rozn2 = QLabel('По рожней 2 территории:', self)
        self.lbl_rozn2.setFixedSize(200, 20)
        
        self.line_rozn2 = QLineEdit(self)
        self.line_rozn2.setFixedSize(120, 20)
        self.line_rozn2.setText('0')
        self.line_rozn2.setValidator(validator)
        
        self.lbl_uskor1 = QLabel('За ускорение 1 территории:', self)
        self.lbl_uskor1.setFixedSize(200, 20)
        
        self.line_uskor1 = QLineEdit(self)
        self.line_uskor1.setFixedSize(120, 20)
        self.line_uskor1.setText('0')
        self.line_uskor1.setValidator(validator)
        
        self.lbl_uskor2 = QLabel('За ускорение 2 территории:', self)
        self.lbl_uskor2.setFixedSize(200, 20)
        
        self.line_uskor2 = QLineEdit(self)
        self.line_uskor2.setFixedSize(120, 20)
        self.line_uskor2.setText('0')
        self.line_uskor2.setValidator(validator)
        
        self.lbl_tel = QLabel('Телеграммы:', self)
        self.lbl_tel.setFixedSize(200, 20)
        
        self.line_tel = QLineEdit(self)
        self.line_tel.setFixedSize(120, 20)
        self.line_tel.setText('0')
        self.line_tel.setValidator(validator)
        
        self.lbl_others = QLabel('Прочие расходы:', self)
        self.lbl_others.setFixedSize(200, 20)
        
        self.line_others = QLineEdit(self)
        self.line_others.setFixedSize(120, 20)
        self.line_others.setText('0')
        self.line_others.setValidator(validator)
        
        self.lbl_dohod = QLabel('Приходы:', self)
        self.lbl_dohod.setFixedSize(200, 20)
        
        self.line_dohod = QLineEdit(self)
        self.line_dohod.setFixedSize(120, 20)
        self.line_dohod.setText('0')
        self.line_dohod.setValidator(validator)
        
        self.grid_oborot.addWidget(self.lbl_rent, 0, 0)
        self.grid_oborot.addWidget(self.line_rent, 0, 1)
        self.grid_oborot.addWidget(self.lbl_remont, 1, 0)
        self.grid_oborot.addWidget(self.line_remont, 1, 1)
        self.grid_oborot.addWidget(self.lbl_pps, 2, 0)
        self.grid_oborot.addWidget(self.line_pps, 2, 1)
        self.grid_oborot.addWidget(self.lbl_gruz, 3, 0)
        self.grid_oborot.addWidget(self.line_gruz, 3, 1)
        self.grid_oborot.addWidget(self.lbl_rozn1, 4, 0)
        self.grid_oborot.addWidget(self.line_rozn1, 4, 1)
        self.grid_oborot.addWidget(self.lbl_rozn2, 5, 0)
        self.grid_oborot.addWidget(self.line_rozn2, 5, 1)
        self.grid_oborot.addWidget(self.lbl_uskor1, 6, 0)
        self.grid_oborot.addWidget(self.line_uskor1, 6, 1)
        self.grid_oborot.addWidget(self.lbl_uskor2, 7, 0)
        self.grid_oborot.addWidget(self.line_uskor2, 7, 1)
        self.grid_oborot.addWidget(self.lbl_tel, 8, 0)
        self.grid_oborot.addWidget(self.line_tel, 8, 1)
        self.grid_oborot.addWidget(self.lbl_others, 9, 0)
        self.grid_oborot.addWidget(self.line_others, 9, 1)
        self.grid_oborot.addWidget(self.lbl_dohod, 10, 0)
        self.grid_oborot.addWidget(self.line_dohod, 10, 1)
        
        self.group_oborot.setLayout(self.grid_oborot)
        self.group_oborot.move(460, 530)
        
        self.btn_adddates = QPushButton('Добавить', self)
        self.btn_adddates.setFixedSize(100, 20)
        self.btn_adddates.move(890, 950)
        self.btn_adddates.clicked.connect(self.Add_dates)
        
        self.group_delete = QGroupBox('Удалить строку №', self)
        self.group_delete.setFixedSize(320, 60)
        self.group_delete.setCheckable(True)
        self.group_delete.setChecked(False)
        self.group_delete.setAlignment(Qt.AlignCenter)
        self.grid_delete = QGridLayout(self)
 
        self.line_delete = QLineEdit(self)
        self.line_delete.setFixedSize(40, 20)
        self.line_delete.setValidator(validator)
        
        self.btn_delete = QPushButton('Удалить', self)
        self.btn_delete.setFixedSize(100, 20)
        self.btn_delete.clicked.connect(self.DeleteRow)
        
        self.grid_delete.addWidget(self.line_delete, 0, 0)
        self.grid_delete.addWidget(self.btn_delete, 0, 1)
        self.group_delete.setLayout(self.grid_delete)
        self.group_delete.move(1520, 30)
        
        self.group_choose = QGroupBox('Выбранная ячейка', self)
        self.group_choose.setFixedSize(320, 100)
        self.group_choose.setAlignment(Qt.AlignCenter)
        self.grid_choose = QGridLayout(self)
        
        self.lbl_cell = QLabel('Значение ячейки:', self)
        self.lbl_cell.setFixedSize(200, 20)
        
        self.lbl_cell_val = QLabel(self)
        self.lbl_cell_val.setFixedSize(300, 20)
        
        self.lbl_cell_r = QLabel('Строка выб. ячейки:', self)
        self.lbl_cell_r.setFixedSize(200, 20)
        
        self.lbl_cell_rt = QLabel(self)
        self.lbl_cell_rt.setFixedSize(100, 20)
        
        self.lbl_cell_c = QLabel('Столбец выб. ячейки:', self)
        self.lbl_cell_c.setFixedSize(200, 20)
        
        self.lbl_cell_ct = QLabel(self)
        self.lbl_cell_ct.setFixedSize(100, 20)
        
        self.grid_choose.addWidget(self.lbl_cell, 0, 0)
        self.grid_choose.addWidget(self.lbl_cell_val, 0, 1)
        self.grid_choose.addWidget(self.lbl_cell_r, 1, 0)
        self.grid_choose.addWidget(self.lbl_cell_rt, 1, 1)
        self.grid_choose.addWidget(self.lbl_cell_c, 2, 0)
        self.grid_choose.addWidget(self.lbl_cell_ct, 2, 1)
        self.group_choose.setLayout(self.grid_choose)
        self.group_choose.move(1520, 120)
        
        self.group_change = QGroupBox('Изменить выб. ячейку на:', self)
        self.group_change.setFixedSize(320, 60)
        self.group_change.setAlignment(Qt.AlignCenter)
        self.group_change.setCheckable(True)
        self.group_change.setChecked(False)
        self.grid_change = QGridLayout(self)
        
        self.line_change = QLineEdit(self)
        self.line_change.setFixedSize(100, 20)
        
        self.btn_change = QPushButton('Изменить', self)
        self.btn_change.setFixedSize(100, 20)
        self.btn_change.clicked.connect(self.ChangeValue)
        
        self.grid_change.addWidget(self.line_change, 0, 0)
        self.grid_change.addWidget(self.btn_change, 0, 1)
        self.group_change.setLayout(self.grid_change)
        self.group_change.move(1520, 250)
        
        self.group_report = QGroupBox('Отчеты', self)
        self.group_report.setFixedSize(320, 150)
        self.group_report.setAlignment(Qt.AlignCenter)
        self.group_report.setCheckable(True)
        self.group_report.setChecked(False)
        self.grid_report = QGridLayout(self)
        
        self.btn_month_rep = QPushButton('Ежемесячный', self)
        self.btn_month_rep.setFixedSize(100, 20)
        self.btn_month_rep.clicked.connect(self.MonthReport)
        
        self.btn_month_arhiv = QPushButton('Архивация', self)
        self.btn_month_arhiv.setFixedSize(100, 20)
        self.btn_month_arhiv.clicked.connect(self.MonthArhiv)
        
        self.btn_quar_rep = QPushButton('Ежеквартальный', self)
        self.btn_quar_rep.setFixedSize(100, 20)
        self.btn_quar_rep.clicked.connect(self.QuarterReport)
        
        self.btn_quar_arhiv = QPushButton('Архивация', self)
        self.btn_quar_arhiv.setFixedSize(100, 20)
        self.btn_quar_arhiv.clicked.connect(self.QuarterArhiv)
        
        self.btn_year_rep = QPushButton('Ежегодный', self)
        self.btn_year_rep.setFixedSize(100, 20)
        self.btn_year_rep.clicked.connect(self.YearlyReport)
        
        self.btn_year_arhiv = QPushButton('Архивация', self)
        self.btn_year_arhiv.setFixedSize(100, 20)
        self.btn_year_arhiv.clicked.connect(self.YearArhiv)
        
        self.grid_report.addWidget(self.btn_month_rep, 0, 0)
        self.grid_report.addWidget(self.btn_month_arhiv, 0, 1)
        self.grid_report.addWidget(self.btn_quar_rep, 1, 0)
        self.grid_report.addWidget(self.btn_quar_arhiv, 1, 1)
        self.grid_report.addWidget(self.btn_year_rep, 2, 0)
        self.grid_report.addWidget(self.btn_year_arhiv, 2, 1)
        self.group_report.setLayout(self.grid_report)
        self.group_report.move(1520, 340)
        
        self.group_graphs = QGroupBox('Графики', self)
        self.group_graphs.setFixedSize(155, 200)
        self.group_graphs.setAlignment(Qt.AlignCenter)
        self.vbox_graphs = QVBoxLayout(self)
        
        self.btn_graph1 = QPushButton('Доходы и расходы', self)
        self.btn_graph1.setFixedHeight(20)
        self.btn_graph1.clicked.connect(self.FirstReport)
        
        self.btn_graph2 = QPushButton('Доходы и расходы, %', self)
        self.btn_graph2.setFixedHeight(20)
        self.btn_graph2.clicked.connect(self.SecondReport)
        
        self.btn_graph3 = QPushButton('Доходы', self)
        self.btn_graph3.setFixedHeight(20)
        self.btn_graph3.clicked.connect(self.ThirdReport)
        
        self.btn_graph4 = QPushButton('Расходы', self)
        self.btn_graph4.setFixedHeight(20)
        self.btn_graph4.clicked.connect(self.FourthReport)
        
        self.btn_graph5 = QPushButton('Сальдо', self)
        self.btn_graph5.setFixedHeight(20)
        self.btn_graph5.clicked.connect(self.FifthReport)
        
        self.vbox_graphs.addWidget(self.btn_graph1)
        self.vbox_graphs.addWidget(self.btn_graph2)
        self.vbox_graphs.addWidget(self.btn_graph3)
        self.vbox_graphs.addWidget(self.btn_graph4)
        self.vbox_graphs.addWidget(self.btn_graph5)
        self.group_graphs.setLayout(self.vbox_graphs)
        self.group_graphs.move(890, 610)
        
        self.group_analiz = QGroupBox('Анализировать', self)
        self.group_analiz.setFixedSize(155, 50)
        self.group_analiz.setAlignment(Qt.AlignCenter)
        self.hbox_analiz = QHBoxLayout(self)
        
        self.lbl_from = QLabel('с:', self)
        self.line_from = QLineEdit(self)
        self.line_from.setFixedSize(40, 20)
        self.line_from.setValidator(validator)
        
        self.lbl_till = QLabel('по:', self)
        self.line_till = QLineEdit(self)
        self.line_till.setFixedSize(40, 20)
        self.line_till.setValidator(validator)
        
        self.hbox_analiz.addWidget(self.lbl_from)
        self.hbox_analiz.addWidget(self.line_from)
        self.hbox_analiz.addWidget(self.lbl_till)
        self.hbox_analiz.addWidget(self.line_till)
        self.group_analiz.setLayout(self.hbox_analiz)
        self.group_analiz.move(890, 530)
        
        self.show()
        
    def Add_dates(self):
        vagon = self.line_vagon.text()
        path = self.line_path.text()
        dest = self.line_dest.text()
        date_path = self.line_date_path.text()
        date_dest = self.line_date_dest.text()
        kms = self.line_km_path.text()
        days_path = self.line_days_path.text()
        days_dest = self.line_days_dest.text()
        days_load = self.line_days_load.text()
        days_unload = self.line_days_unload.text()
        date_load = self.line_date_load.text()
        date_unload = self.line_date_unload.text()
        date_to_rem = self.line_date_to_remont.text()
        date_from_rem = self.line_date_from_remont.text()
        days_rem = self.line_days_remont.text()
        rent = self.line_rent.text()
        remont = self.line_remont.text()
        pps = self.line_pps.text()
        gruz = self.line_gruz.text()
        rozn1 = self.line_rozn1.text()
        rozn2 = self.line_rozn2.text()
        uskor1 = self.line_uskor1.text()
        uskor2 = self.line_uskor2.text()
        tels = self.line_tel.text()
        others = self.line_others.text()
        dohod = self.line_dohod.text()
        
        isAddedData = False
        isCorrData = False
        count = 0
        if (vagon != '')&(path != '')&(dest != '')&(kms != '')&(days_path != '')&(days_dest != '')&(days_load != '')&(days_unload != '')&(rent != '')&(remont != '')&(pps != '')&(gruz != '')&(rozn1 != '')&(rozn2 != '')&(uskor1 != '')&(uskor2 != '')&(tels != '')&(others != '')&(days_rem != ''):
            isAddedData = True
        if isAddedData == False:
            QMessageBox.information(self, 'Ошибка', 'Необходимо заполнить все поля')
        if (len(date_path) == 10)&(len(date_dest) == 10)&(len(date_load) == 10)&(len(date_unload) == 10)&(len(date_to_rem) == 10)&(len(date_from_rem) == 10): 
            dates = [date_path, date_dest, date_load, date_unload, date_to_rem, date_from_rem]
            for i in range(len(dates)):
                checkDate = dates[i]
                checkDay = int(checkDate[0:2])
                checkMonth = int(checkDate[3:5])
                checkYear = int(checkDate[6:10])
                
                months30 = [4, 6, 9, 11]
                months31 = [1, 3, 5, 7, 8, 10, 12]
                months28 = [2]
                if (checkDay >= 1)&(checkDay <= 30)&(checkMonth in months30):
                    isCorrData = True
                if (checkDay >= 1)&(checkDay <= 31)&(checkMonth in months31):
                    isCorrData = True
                if (checkDay >= 1)&(checkDay <= 28)&(checkMonth in months28)&(checkYear%4 > 0):
                    isCorrData = True
                if (checkDay >= 1)&(checkDay <= 29)&(checkMonth in months28)&(checkYear%4 == 0):
                    isCorrData = True
                if (isCorrData):
                    count = count + 1
            if (isCorrData == False)&(count < 6):
                    QMessageBox.information(self, 'Ошибка', 'Некорректно заполнены поля')
        
        if (isAddedData)&(count == 6)&(isCorrData):
            now = datetime.datetime.now()
            now_year = now.year
            now_month = now.month
            now_day = now.day
            now_hour = now.hour
            now_min = now.minute
            now_sec = now.second
            if now_month <= 9:
                now_month_s = '0' + str(now_month)
            if now_month > 9:
                now_month_s = str(now_month)
            if now_day <= 9:
                now_day_s = '0' + str(now_day)
            if now_day > 9:
                now_day_s = str(now_day)
            if now_hour <= 9:
                now_hour_s = '0' + str(now_hour)
            if now_hour > 9:
                now_hour_s = str(now_hour)
            if now_min <= 9:
                now_min_s = '0' + str(now_min)
            if now_min > 9:
                now_min_s = str(now_min)
            if (now_sec <= 9):
                now_sec_s = '0' + str(now_sec)
            if now_sec > 9:
                now_sec_s = str(now_sec)
                
            ids = '#'+str(now_year)+now_month_s+now_day_s+now_hour_s+now_min_s+now_sec_s
            add_date = now_day_s+'.'+now_month_s+'.'+str(now_year)
            
            tab = []
            tab_arhiv = []
            total_rashod = int(rent)+int(remont)+int(gruz)+int(rozn1)+int(rozn2)+int(uskor1)+int(uskor2)+int(tels)+int(others)
            total_dohod = int(dohod) + int(pps)
            total_days = int(days_path) + int(days_dest) + int(days_load) + int(days_unload) + int(days_rem)
            total_days2 = total_days
            date_start = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
            date_end = date_start + datetime.timedelta(days=total_days)
            months_between = (date_end.year - date_start.year)*12 + (date_end.month - date_start.month)
            saldo = total_dohod - total_rashod
            date_path_to_add = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
            date_dest_to_add = datetime.date(int(date_dest[6:10]), int(date_dest[3:5]), int(date_dest[0:2]))
            days_in_path = (date_dest_to_add - date_path_to_add).days
            tab.append({'ID':ids, 'ADD_DATE':add_date, 'VAGON':vagon, 'PATH':path, 'DEST':dest, 'DATE_PATH':date_path, 'DATE_DEST':date_dest, 'KMS':kms, 'DAYS_PATH':days_path, 'DAYS_DEST':days_dest, 'DAYS_LOAD':days_load, 'DAYS_UNLOAD':days_unload,
                       'DATE_LOAD':date_load, 'DATE_UNLOAD':date_unload, 'DATE_TO_REM':date_to_rem, 'DATE_FROM_REM':date_from_rem, 'DAYS_REM':days_rem, 'RENT':rent, 'REMONT':remont, 'PPS':pps,'GRUZ':gruz, 
                        'ROZN1':rozn1, 'ROZN2':rozn2, 'USKOR1':uskor1, 'USKOR2':uskor2, 'TELS':tels, 'OTHERS':others, 'DOHOD':dohod, 'DAYS_IN_PATH':str(days_in_path), 'TOTAL_RASHOD':str(total_rashod), 'TOTAL_DOHOD':str(total_dohod), 'SALDO':str(saldo)})

            data2 = pd.DataFrame(tab)
            self.data = pd.concat([self.data, data2], axis=0)
            self.data = self.data.reindex(['ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 'DATE_UNLOAD', 'DATE_TO_REM',
                                            'DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 'OTHERS', 'DOHOD', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'], axis=1)
            
            for i in range(months_between + 1):
                next_month = date_start + relativedelta(months=i+1)
                oper_date = datetime.date(next_month.year, next_month.month, 1)
                if (i == 0)&(date_start + datetime.timedelta(days=total_days) <= datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                    print(oper_date, total_days/total_days2)
                    coef = total_days/total_days2
                if (i == 0)&(date_start + datetime.timedelta(days=total_days) > datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                    print(oper_date, (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2)
                    coef = (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2
                    total_days = total_days - (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)
                if (i > 0)&(i < months_between):
                    oper_date2 = oper_date + datetime.timedelta(days=-1)
                    print(oper_date, (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2)
                    coef = (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2
                    total_days = total_days - calendar.monthrange(oper_date2.year, oper_date2.month)[1]
                if (i > 0)&(i >= months_between):
                    print(oper_date, total_days/total_days2)
                    coef = total_days/total_days2
                tab_arhiv.append({'OPER_DATE':oper_date, 'ID':ids, 'VAGON':vagon, 'ADD_DATE':add_date, 'PATH':path, 'DEST':dest, 'DATE_PATH':date_path, 'DATE_DEST':date_dest, 'KMS':kms, 'DAYS_PATH':days_path, 
                                      'DAYS_DEST':days_dest, 'DAYS_LOAD':days_load, 'DAYS_UNLOAD':days_unload, 'DATE_LOAD':date_load, 'DATE_UNLOAD':date_unload, 'DATE_TO_REM':date_to_rem,'DATE_FROM_REM':date_from_rem, 
                                      'DAYS_REM':days_rem, 'RENT':rent, 'REMONT':remont, 'PPS':pps, 'GRUZ':gruz, 'ROZN1':rozn1, 'ROZN2':rozn2, 'USKOR1':uskor1, 'USKOR2':uskor2, 'TELS':tels, 
                                       'OTHERS':others, 'DOHOD':dohod, 'COEF':coef, 'DAYS_IN_PATH':str(int(days_in_path*coef)), 'TOTAL_RASHOD':str(total_rashod*coef), 'TOTAL_DOHOD':str(total_dohod*coef), 'SALDO':str(saldo*coef)})
                        
            data_arhiv = pd.DataFrame(tab_arhiv)
            self.arhiv = pd.read_csv(self.arhivPath())
            self.arhiv = pd.concat([self.arhiv, data_arhiv], axis=0)
            self.arhiv = self.arhiv.reindex(['OPER_DATE', 'ID', 'VAGON', 'ADD_DATE', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                               'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                               'OTHERS', 'DOHOD', 'COEF', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'], axis=1)
            
            self.arhiv.to_csv(self.arhivPath(), index=False)
            self.data.to_csv(self.dataPath(), index=False)

            self.data = pd.read_csv(self.dataPath())

            self.nd = NewDialog(self)
            self.nd.setGeometry(20, 20, 1460, 500)
            self.nd.tableWidget = QTableWidget()
            self.nd.tableWidget.setRowCount(len(self.data.index))
            self.nd.tableWidget.setColumnCount(len(self.data.columns))
            for i in range(len(self.data.index)):
                for j in range(len(self.data.columns)):
                    self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
            self.nd.tableWidget.move(20,20)
            header = self.Header()
            self.nd.tableWidget.setHorizontalHeaderLabels(header)
            self.nd.tableWidget.resizeColumnsToContents()
            self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
            self.nd.layout = QVBoxLayout()
            self.nd.layout.addWidget(self.nd.tableWidget) 
            self.nd.setLayout(self.nd.layout)
            self.nd.show()

            self.line_vagon.setText('')
            self.line_path.setText('')
            self.line_dest.setText('')
            self.line_date_path.setText('')
            self.line_date_dest.setText('')
            self.line_km_path.setText('0')
            self.line_days_path.setText('0')
            self.line_days_dest.setText('0')
            self.line_days_load.setText('0')
            self.line_days_unload.setText('0')
            self.line_date_load.setText('')
            self.line_date_unload.setText('')
            self.line_date_to_remont.setText('')
            self.line_date_from_remont.setText('')
            self.line_rent.setText('0')
            self.line_remont.setText('0')
            self.line_pps.setText('0')
            self.line_gruz.setText('0')
            self.line_rozn1.setText('0')
            self.line_rozn2.setText('0')
            self.line_uskor1.setText('0')
            self.line_uskor2.setText('0')
            self.line_tel.setText('0')
            self.line_others.setText('0')
            self.line_dohod.setText('0')
            self.line_days_remont.setText('0')
            
    def DeleteRow(self):
        self.data = pd.read_csv(self.dataPath())
        self.arhiv = pd.read_csv(self.arhivPath())
        if (self.line_delete.text() == ''):
            QMessageBox.information(self, 'Ошибка', 'Введите номер строки')
        if (self.line_delete.text() != '')&(((int(self.line_delete.text()) - 1) > (len(self.data) - 1))|((int(self.line_delete.text()) - 1) < 0)):
            QMessageBox.information(self, 'Ошибка', 'Данная строка не является а пределах таблицы')
        else:
            confirm = QMessageBox.question(self, '','Вы точно хотите удалить данную строку?', QMessageBox.Yes | QMessageBox.No)
            if confirm == QMessageBox.Yes:
                row = int(self.line_delete.text())
                arhiv_id = self.data.iloc[row-1]['ID']
                self.data = self.data.drop(row-1)
                self.data = self.data.reset_index(drop=True)
                self.data.to_csv(self.dataPath(), index=False)
                
                self.arhiv = self.arhiv[self.arhiv.ID != arhiv_id]
                self.arhiv = self.arhiv.reset_index(drop=True)
                self.arhiv.to_csv(self.arhivPath(), index=False)

                self.data = pd.read_csv(self.dataPath())

                self.nd = NewDialog(self)
                self.nd.setGeometry(20, 20, 1460, 500)
                self.nd.tableWidget = QTableWidget()
                self.nd.tableWidget.setRowCount(len(self.data.index))
                self.nd.tableWidget.setColumnCount(len(self.data.columns))
                for i in range(len(self.data.index)):
                    for j in range(len(self.data.columns)):
                        self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                self.nd.tableWidget.move(20,20)
                header = self.Header()
                self.nd.tableWidget.setHorizontalHeaderLabels(header)
                self.nd.tableWidget.resizeColumnsToContents()
                self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                self.nd.layout = QVBoxLayout()
                self.nd.layout.addWidget(self.nd.tableWidget) 
                self.nd.setLayout(self.nd.layout)
                self.nd.show()
                
                self.line_delete.setText('')
                
    def ChangeValue(self):
        self.data = pd.read_csv(self.dataPath())
        self.arhiv = pd.read_csv(self.arhivPath())
        if self.lbl_cell_rt.text() != '':
            row = int(self.lbl_cell_rt.text()) - 1
        if self.lbl_cell_ct.text() != '':
            col = int(self.lbl_cell_ct.text()) - 1
        if (self.lbl_cell_rt.text() != '')&(self.lbl_cell_ct.text() != ''):
            if (col < 28)&(col >= 0)&(row >= 0)&(row < len(self.data)):
                arr_digits = [7, 8, 9, 10, 11, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27]
                arr_head = ['ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 'DATE_UNLOAD', 'DATE_TO_REM',
                            'DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 'OTHERS', 'DOHOD', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO']
                arr_dates = [5, 6, 12, 13, 14, 15]
                arr_others = [2, 3, 4]
                tot_rashod = 0
                if col in arr_digits:
                    if (self.line_change.text().isdigit() == True):
                        arhiv_id = self.data.iloc[row]['ID']
                        self.data.at[self.data.index[row], arr_head[col]] = self.line_change.text()
                        if (col >= 17)&(col <= 26)&(col != 19):
                            tot_rashod = int(self.data.iloc[row]['RENT']) + int(self.data.iloc[row]['REMONT']) + int(self.data.iloc[row]['GRUZ']) + int(self.data.iloc[row]['ROZN1']) + int(self.data.iloc[row]['ROZN2']) + int(self.data.iloc[row]['USKOR1']) + int(self.data.iloc[row]['USKOR2']) + int(self.data.iloc[row]['TELS']) + int(self.data.iloc[row]['OTHERS'])
                            self.data.at[self.data.index[row], 'TOTAL_RASHOD'] = tot_rashod
                        if (col == 27)|(col == 19):
                            self.data.at[self.data.index[row], 'TOTAL_DOHOD'] = int(self.data.iloc[row][27]) + int(self.data.iloc[row][19])
                        self.data.at[self.data.index[row], 'SALDO'] = int(self.data.iloc[row]['TOTAL_DOHOD']) - int(self.data.iloc[row]['TOTAL_RASHOD'])
                            
                        self.data.to_csv(self.dataPath(), index=False)
                        self.data = pd.read_csv(self.dataPath())
                        self.nd = NewDialog(self)
                        self.nd.setGeometry(20, 20, 1460, 500)
                        self.nd.tableWidget = QTableWidget()
                        self.nd.tableWidget.setRowCount(len(self.data.index))
                        self.nd.tableWidget.setColumnCount(len(self.data.columns))
                        for i in range(len(self.data.index)):
                            for j in range(len(self.data.columns)):
                                self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                        self.nd.tableWidget.move(20,20)
                        header = self.Header()
                        self.nd.tableWidget.setHorizontalHeaderLabels(header)
                        self.nd.tableWidget.resizeColumnsToContents()
                        self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                        self.nd.layout = QVBoxLayout()
                        self.nd.layout.addWidget(self.nd.tableWidget) 
                        self.nd.setLayout(self.nd.layout)
                        self.nd.show()
                        
                        self.arhiv = self.arhiv[self.arhiv.ID != arhiv_id]
                        tab_arhiv = []
                        for i in range(len(self.data)):
                            if self.data.iloc[i]['ID'] == arhiv_id:
                                add_date = self.data.iloc[i]['ADD_DATE']
                                vagon = self.data.iloc[i]['VAGON']
                                path = self.data.iloc[i]['PATH']
                                dest = self.data.iloc[i]['DEST']
                                date_path = self.data.iloc[i]['DATE_PATH']
                                date_dest = self.data.iloc[i]['DATE_DEST']
                                kms = self.data.iloc[i]['KMS']
                                days_path = self.data.iloc[i]['DAYS_PATH']
                                days_dest = self.data.iloc[i]['DAYS_DEST']
                                days_load = self.data.iloc[i]['DAYS_LOAD']
                                days_unload = self.data.iloc[i]['DAYS_UNLOAD']
                                date_load = self.data.iloc[i]['DATE_LOAD']
                                date_unload = self.data.iloc[i]['DATE_UNLOAD']
                                date_to_rem = self.data.iloc[i]['DATE_TO_REM']
                                date_from_rem = self.data.iloc[i]['DATE_FROM_REM']
                                days_rem = self.data.iloc[i]['DAYS_REM']
                                rent = self.data.iloc[i]['RENT']
                                remont = self.data.iloc[i]['REMONT']
                                pps = self.data.iloc[i]['PPS']
                                gruz = self.data.iloc[i]['GRUZ']
                                rozn1 = self.data.iloc[i]['ROZN1']
                                rozn2 = self.data.iloc[i]['ROZN2']
                                uskor1 = self.data.iloc[i]['USKOR1']
                                uskor2 = self.data.iloc[i]['USKOR2']
                                tels = self.data.iloc[i]['TELS']
                                others = self.data.iloc[i]['OTHERS']
                                dohod = self.data.iloc[i]['DOHOD']
                                days_in_path = int(self.data.iloc[i]['DAYS_IN_PATH'])
                                total_rashod = int(self.data.iloc[i]['TOTAL_RASHOD'])
                                total_dohod = int(self.data.iloc[i]['TOTAL_DOHOD'])
                                saldo = int(self.data.iloc[i]['SALDO'])
                                
                        total_days = int(days_path) + int(days_dest) + int(days_load) + int(days_unload) + int(days_rem)
                        total_days2 = total_days
                        date_start = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        date_end = date_start + datetime.timedelta(days=total_days)
                        months_between = (date_end.year - date_start.year)*12 + (date_end.month - date_start.month)
                        for i in range(months_between + 1):
                            next_month = date_start + relativedelta(months=i+1)
                            oper_date = datetime.date(next_month.year, next_month.month, 1)
                            if (i == 0)&(date_start + datetime.timedelta(days=total_days) <= datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                print(oper_date, total_days/total_days2)
                                coef = total_days/total_days2
                            if (i == 0)&(date_start + datetime.timedelta(days=total_days) > datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                print(oper_date, (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2)
                                coef = (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2
                                total_days = total_days - (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)
                            if (i > 0)&(i < months_between):
                                oper_date2 = oper_date + datetime.timedelta(days=-1)
                                print(oper_date, (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2)
                                coef = (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2
                                total_days = total_days - calendar.monthrange(oper_date2.year, oper_date2.month)[1]
                            if (i > 0)&(i >= months_between):
                                print(oper_date, total_days/total_days2)
                                coef = total_days/total_days2
                            tab_arhiv.append({'OPER_DATE':oper_date, 'ID':arhiv_id, 'ADD_DATE':add_date, 'VAGON':vagon, 'PATH':path, 'DEST':dest, 'DATE_PATH':date_path, 'DATE_DEST':date_dest, 'KMS':kms, 'DAYS_PATH':days_path, 
                                                  'DAYS_DEST':days_dest, 'DAYS_LOAD':days_load, 'DAYS_UNLOAD':days_unload, 'DATE_LOAD':date_load, 'DATE_UNLOAD':date_unload, 'DATE_TO_REM':date_to_rem,'DATE_FROM_REM':date_from_rem, 
                                                  'DAYS_REM':days_rem, 'RENT':rent, 'REMONT':remont, 'PPS':pps, 'GRUZ':gruz, 'ROZN1':rozn1, 'ROZN2':rozn2, 'USKOR1':uskor1, 'USKOR2':uskor2, 'TELS':tels, 
                                                   'OTHERS':others, 'DOHOD':dohod, 'COEF':coef, 'DAYS_IN_PATH':str(int(coef*days_in_path)), 'TOTAL_RASHOD':str(total_rashod*coef), 'TOTAL_DOHOD':str(total_dohod*coef), 'SALDO':str(saldo*coef)})

                        data_arhiv = pd.DataFrame(tab_arhiv)
                        self.arhiv = pd.concat([self.arhiv, data_arhiv], axis=0)
                        self.arhiv = self.arhiv.reindex(['OPER_DATE', 'ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                                           'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                                           'OTHERS', 'DOHOD', 'COEF', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'], axis=1)

                        self.arhiv.to_csv(self.arhivPath(), index=False)
                        self.data.to_csv(self.dataPath(), index=False)
                        
                    else:
                        QMessageBox.information(self, 'Ошибка', 'Данные в этой ячейке должны быть цифрой')
                if (col in arr_dates)&(len(self.line_change.text()) == 10):
                    if (self.line_change.text()[0:2].isdigit() == True)&(self.line_change.text()[2] == '.')&(self.line_change.text()[3:5].isdigit() == True)&(self.line_change.text()[5] == '.')&(self.line_change.text()[6:10].isdigit() == True):
                        isCorrect = False
                        chDate = self.line_change.text()
                        chDay = int(chDate[0:2])
                        chMonth = int(chDate[3:5])
                        chYear = int(chDate[6:10])
                        months30 = [4, 6, 9, 11]
                        months31 = [1, 3, 5, 7, 8, 10, 12]
                        months28 = [2]
                        if (chDay >= 1)&(chDay <= 30)&(chMonth in months30):
                            isCorrect = True
                        if (chDay >= 1)&(chDay <= 31)&(chMonth in months31):
                            isCorrect = True
                        if (chDay >= 1)&(chDay <= 28)&(chMonth in months28)&(chYear%4 > 0):
                            isCorrect = True
                        if (chDay >= 1)&(chDay <= 29)&(chMonth in months28)&(chYear%4 == 0):
                            isCorrect = True
                        
                        if (isCorrect):
                            arhiv_id = self.data.iloc[row]['ID']
                            self.data.at[self.data.index[row], arr_head[col]] = self.line_change.text()
                            date_path = self.data.iloc[row]['DATE_PATH']
                            date_dest = self.data.iloc[row]['DATE_DEST']
                            date_path_to_add = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                            date_dest_to_add = datetime.date(int(date_dest[6:10]), int(date_dest[3:5]), int(date_dest[0:2]))
                            days_in_path = (date_dest_to_add - date_path_to_add).days
                            self.data.at[self.data.index[row], 'DAYS_IN_PATH'] = str(days_in_path)
                            self.data.to_csv(self.dataPath(), index=False)
                            self.data = pd.read_csv(self.dataPath())
                            self.nd = NewDialog(self)
                            self.nd.setGeometry(20, 20, 1460, 500)
                            self.nd.tableWidget = QTableWidget()
                            self.nd.tableWidget.setRowCount(len(self.data.index))
                            self.nd.tableWidget.setColumnCount(len(self.data.columns))
                            for i in range(len(self.data.index)):
                                for j in range(len(self.data.columns)):
                                    self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                            self.nd.tableWidget.move(20,20)
                            header = self.Header()
                            self.nd.tableWidget.setHorizontalHeaderLabels(header)
                            self.nd.tableWidget.resizeColumnsToContents()
                            self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                            self.nd.layout = QVBoxLayout()
                            self.nd.layout.addWidget(self.nd.tableWidget) 
                            self.nd.setLayout(self.nd.layout)
                            self.nd.show()
                            
                            self.arhiv = self.arhiv[self.arhiv.ID != arhiv_id]
                            tab_arhiv = []
                            for i in range(len(self.data)):
                                if self.data.iloc[i]['ID'] == arhiv_id:
                                    add_date = self.data.iloc[i]['ADD_DATE']
                                    vagon = self.data.iloc[i]['VAGON']
                                    path = self.data.iloc[i]['PATH']
                                    dest = self.data.iloc[i]['DEST']
                                    date_path = self.data.iloc[i]['DATE_PATH']
                                    date_dest = self.data.iloc[i]['DATE_DEST']
                                    kms = self.data.iloc[i]['KMS']
                                    days_path = self.data.iloc[i]['DAYS_PATH']
                                    days_dest = self.data.iloc[i]['DAYS_DEST']
                                    days_load = self.data.iloc[i]['DAYS_LOAD']
                                    days_unload = self.data.iloc[i]['DAYS_UNLOAD']
                                    date_load = self.data.iloc[i]['DATE_LOAD']
                                    date_unload = self.data.iloc[i]['DATE_UNLOAD']
                                    date_to_rem = self.data.iloc[i]['DATE_TO_REM']
                                    date_from_rem = self.data.iloc[i]['DATE_FROM_REM']
                                    days_rem = self.data.iloc[i]['DAYS_REM']
                                    rent = self.data.iloc[i]['RENT']
                                    remont = self.data.iloc[i]['REMONT']
                                    pps = self.data.iloc[i]['PPS']
                                    gruz = self.data.iloc[i]['GRUZ']
                                    rozn1 = self.data.iloc[i]['ROZN1']
                                    rozn2 = self.data.iloc[i]['ROZN2']
                                    uskor1 = self.data.iloc[i]['USKOR1']
                                    uskor2 = self.data.iloc[i]['USKOR2']
                                    tels = self.data.iloc[i]['TELS']
                                    others = self.data.iloc[i]['OTHERS']
                                    dohod = self.data.iloc[i]['DOHOD']
                                    days_in_path = int(self.data.iloc[i]['DAYS_IN_PATH'])
                                    total_rashod = int(self.data.iloc[i]['TOTAL_RASHOD'])
                                    total_dohod = int(self.data.iloc[i]['TOTAL_DOHOD'])
                                    saldo = int(self.data.iloc[i]['SALDO'])

                            total_days = int(days_path) + int(days_dest) + int(days_load) + int(days_unload) + int(days_rem)
                            total_days2 = total_days
                            date_start = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                            date_end = date_start + datetime.timedelta(days=total_days)
                            months_between = (date_end.year - date_start.year)*12 + (date_end.month - date_start.month)
                            for i in range(months_between + 1):
                                next_month = date_start + relativedelta(months=i+1)
                                oper_date = datetime.date(next_month.year, next_month.month, 1)
                                if (i == 0)&(date_start + datetime.timedelta(days=total_days) <= datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                    print(oper_date, total_days/total_days2)
                                    coef = total_days/total_days2
                                if (i == 0)&(date_start + datetime.timedelta(days=total_days) > datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                    print(oper_date, (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2)
                                    coef = (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2
                                    total_days = total_days - (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)
                                if (i > 0)&(i < months_between):
                                    oper_date2 = oper_date + datetime.timedelta(days=-1)
                                    print(oper_date, (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2)
                                    coef = (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2
                                    total_days = total_days - calendar.monthrange(oper_date2.year, oper_date2.month)[1]
                                if (i > 0)&(i >= months_between):
                                    print(oper_date, total_days/total_days2)
                                    coef = total_days/total_days2
                                tab_arhiv.append({'OPER_DATE':oper_date, 'ID':arhiv_id, 'VAGON':vagon, 'ADD_DATE':add_date, 'PATH':path, 'DEST':dest, 'DATE_PATH':date_path, 'DATE_DEST':date_dest, 'KMS':kms, 'DAYS_PATH':days_path, 
                                                      'DAYS_DEST':days_dest, 'DAYS_LOAD':days_load, 'DAYS_UNLOAD':days_unload, 'DATE_LOAD':date_load, 'DATE_UNLOAD':date_unload, 'DATE_TO_REM':date_to_rem,'DATE_FROM_REM':date_from_rem, 
                                                      'DAYS_REM':days_rem, 'RENT':rent, 'REMONT':remont, 'PPS':pps, 'GRUZ':gruz, 'ROZN1':rozn1, 'ROZN2':rozn2, 'USKOR1':uskor1, 'USKOR2':uskor2, 'TELS':tels, 
                                                       'OTHERS':others, 'DOHOD':dohod, 'COEF':coef, 'DAYS_IN_PATH':str(int(coef*days_in_path)), 'TOTAL_RASHOD':str(total_rashod*coef), 'TOTAL_DOHOD':str(total_dohod*coef), 'SALDO':str(saldo*coef)})

                            data_arhiv = pd.DataFrame(tab_arhiv)
                            self.arhiv = pd.concat([self.arhiv, data_arhiv], axis=0)
                            self.arhiv = self.arhiv.reindex(['OPER_DATE', 'ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                                               'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                                               'OTHERS', 'DOHOD', 'COEF', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'], axis=1)

                            self.arhiv.to_csv(self.arhivPath(), index=False)
                            self.data.to_csv(self.dataPath(), index=False)

                        if (isCorrect == False):
                            QMessageBox.information(self, 'Ошибка', 'Некорректная дата')
                    else:
                        QMessageBox.information(self, 'Ошибка', 'Данные в этой ячейке должны соответствовать формату ДД.ММ.ГГГГ.')
                        
                if col in arr_others:
                    if (self.line_change.text() != ''):
                        arhiv_id = self.data.iloc[row]['ID']
                        self.data.at[self.data.index[row], arr_head[col]] = self.line_change.text()
                        self.data.to_csv(self.dataPath(), index=False)
                        self.data = pd.read_csv(self.dataPath())
                        self.nd = NewDialog(self)
                        self.nd.setGeometry(20, 20, 1460, 500)
                        self.nd.tableWidget = QTableWidget()
                        self.nd.tableWidget.setRowCount(len(self.data.index))
                        self.nd.tableWidget.setColumnCount(len(self.data.columns))
                        for i in range(len(self.data.index)):
                            for j in range(len(self.data.columns)):
                                self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                        self.nd.tableWidget.move(20,20)
                        header = self.Header()
                        self.nd.tableWidget.setHorizontalHeaderLabels(header)
                        self.nd.tableWidget.resizeColumnsToContents()
                        self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                        self.nd.layout = QVBoxLayout()
                        self.nd.layout.addWidget(self.nd.tableWidget) 
                        self.nd.setLayout(self.nd.layout)
                        self.nd.show()
                            
                        self.arhiv = self.arhiv[self.arhiv.ID != arhiv_id]
                        tab_arhiv = []
                        for i in range(len(self.data)):
                            if self.data.iloc[i]['ID'] == arhiv_id:
                                add_date = self.data.iloc[i]['ADD_DATE']
                                vagon = self.data.iloc[i]['VAGON']
                                path = self.data.iloc[i]['PATH']
                                dest = self.data.iloc[i]['DEST']
                                date_path = self.data.iloc[i]['DATE_PATH']
                                date_dest = self.data.iloc[i]['DATE_DEST']
                                kms = self.data.iloc[i]['KMS']
                                days_path = self.data.iloc[i]['DAYS_PATH']
                                days_dest = self.data.iloc[i]['DAYS_DEST']
                                days_load = self.data.iloc[i]['DAYS_LOAD']
                                days_unload = self.data.iloc[i]['DAYS_UNLOAD']
                                date_load = self.data.iloc[i]['DATE_LOAD']
                                date_unload = self.data.iloc[i]['DATE_UNLOAD']
                                date_to_rem = self.data.iloc[i]['DATE_TO_REM']
                                date_from_rem = self.data.iloc[i]['DATE_FROM_REM']
                                days_rem = self.data.iloc[i]['DAYS_REM']
                                rent = self.data.iloc[i]['RENT']
                                remont = self.data.iloc[i]['REMONT']
                                pps = self.data.iloc[i]['PPS']
                                gruz = self.data.iloc[i]['GRUZ']
                                rozn1 = self.data.iloc[i]['ROZN1']
                                rozn2 = self.data.iloc[i]['ROZN2']
                                uskor1 = self.data.iloc[i]['USKOR1']
                                uskor2 = self.data.iloc[i]['USKOR2']
                                tels = self.data.iloc[i]['TELS']
                                others = self.data.iloc[i]['OTHERS']
                                dohod = self.data.iloc[i]['DOHOD']
                                days_in_path = int(self.data.iloc[i]['DAYS_IN_PATH'])
                                total_rashod = int(self.data.iloc[i]['TOTAL_RASHOD'])
                                total_dohod = int(self.data.iloc[i]['TOTAL_DOHOD'])
                                saldo = int(self.data.iloc[i]['SALDO'])

                        total_days = int(days_path) + int(days_dest) + int(days_load) + int(days_unload) + int(days_rem)
                        total_days2 = total_days
                        date_start = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        date_end = date_start + datetime.timedelta(days=total_days)
                        months_between = (date_end.year - date_start.year)*12 + (date_end.month - date_start.month)
                        for i in range(months_between + 1):
                            next_month = date_start + relativedelta(months=i+1)
                            oper_date = datetime.date(next_month.year, next_month.month, 1)
                            if (i == 0)&(date_start + datetime.timedelta(days=total_days) <= datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                print(oper_date, total_days/total_days2)
                                coef = total_days/total_days2
                            if (i == 0)&(date_start + datetime.timedelta(days=total_days) > datetime.date(date_start.year, date_start.month, calendar.monthrange(date_start.year, date_start.month)[1])):
                                print(oper_date, (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2)
                                coef = (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)/total_days2
                                total_days = total_days - (calendar.monthrange(date_start.year, date_start.month)[1] - date_start.day)
                            if (i > 0)&(i < months_between):
                                oper_date2 = oper_date + datetime.timedelta(days=-1)
                                print(oper_date, (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2)
                                coef = (calendar.monthrange(oper_date2.year, oper_date2.month)[1])/total_days2
                                total_days = total_days - calendar.monthrange(oper_date2.year, oper_date2.month)[1]
                            if (i > 0)&(i >= months_between):
                                print(oper_date, total_days/total_days2)
                                coef = total_days/total_days2
                            tab_arhiv.append({'OPER_DATE':oper_date, 'ID':arhiv_id, 'VAGON':vagon, 'ADD_DATE':add_date, 'PATH':path, 'DEST':dest, 'DATE_PATH':date_path, 'DATE_DEST':date_dest, 'KMS':kms, 'DAYS_PATH':days_path, 
                                                'DAYS_DEST':days_dest, 'DAYS_LOAD':days_load, 'DAYS_UNLOAD':days_unload, 'DATE_LOAD':date_load, 'DATE_UNLOAD':date_unload, 'DATE_TO_REM':date_to_rem,'DATE_FROM_REM':date_from_rem, 
                                                'DAYS_REM':days_rem, 'RENT':rent, 'REMONT':remont, 'PPS':pps, 'GRUZ':gruz, 'ROZN1':rozn1, 'ROZN2':rozn2, 'USKOR1':uskor1, 'USKOR2':uskor2, 'TELS':tels, 
                                                'OTHERS':others, 'DOHOD':dohod, 'COEF':coef, 'DAYS_IN_PATH':str(int(coef*days_in_path)), 'TOTAL_RASHOD':str(total_rashod*coef), 'TOTAL_DOHOD':str(total_dohod*coef), 'SALDO':str(saldo*coef)})

                        data_arhiv = pd.DataFrame(tab_arhiv)
                        self.arhiv = pd.concat([self.arhiv, data_arhiv], axis=0)
                        self.arhiv = self.arhiv.reindex(['OPER_DATE', 'ID', 'ADD_DATE', 'VAGON', 'PATH', 'DEST', 'DATE_PATH', 'DATE_DEST', 'KMS', 'DAYS_PATH', 'DAYS_DEST', 'DAYS_LOAD', 'DAYS_UNLOAD', 'DATE_LOAD', 
                                                            'DATE_UNLOAD', 'DATE_TO_REM','DATE_FROM_REM', 'DAYS_REM', 'RENT', 'REMONT', 'PPS', 'GRUZ', 'ROZN1', 'ROZN2', 'USKOR1', 'USKOR2', 'TELS', 
                                                            'OTHERS', 'DOHOD', 'COEF', 'DAYS_IN_PATH', 'TOTAL_RASHOD', 'TOTAL_DOHOD', 'SALDO'], axis=1)

                        self.arhiv.to_csv(self.arhivPath(), index=False)
                        self.data.to_csv(self.dataPath(), index=False)

                        
                
                self.line_change.setText('')
                    
    def ClickedCell(self):
        itemText = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), self.nd.tableWidget.currentColumn()).text()
        itemRow = str(self.nd.tableWidget.currentItem().row() + 1)
        itemCol = str(self.nd.tableWidget.currentItem().column() + 1)
        self.lbl_cell_val.setText(itemText)
        self.lbl_cell_rt.setText(itemRow)
        self.lbl_cell_ct.setText(itemCol)
        
    def MonthReport(self):
        today = datetime.date.today()
        last_month = today + relativedelta(months=-1)
        if last_month.month == 1:
            last_month_s = 'Январь'
        if last_month.month == 2:
            last_month_s = 'Февраль'
        if last_month.month == 3:
            last_month_s = 'Март'
        if last_month.month == 4:
            last_month_s = 'Апрель'
        if last_month.month == 5:
            last_month_s = 'Май'
        if last_month.month == 6:
            last_month_s = 'Июнь'
        if last_month.month == 7:
            last_month_s = 'Июль'
        if last_month.month == 8:
            last_month_s = 'Август'
        if last_month.month == 9:
            last_month_s = 'Сентябрь'
        if last_month.month == 10:
            last_month_s = 'Октябрь'
        if last_month.month == 11:
            last_month_s = 'Ноябрь'
        if last_month.month == 12:
            last_month_s = 'Декабрь'
            
        msg = QMessageBox.question(self, '', 'Вы точно хотите загрузить ежемесячный отчет на '+ last_month_s + ' ' + str(last_month.year) + '?', QMessageBox.Yes|QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.arhiv = pd.read_csv(self.arhivPath())
            for i in range(len(self.arhiv)):
                self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
            monthly_rep = self.arhiv[(self.arhiv.OPER_DATE == datetime.date(today.year, today.month, 1))]
            if len(monthly_rep) <= 0:
                QMessageBox.information(self, 'Информация', 'В данном месяце нет данных')
            if len(monthly_rep) > 0:
                monthly_rep_month = QFileDialog.getSaveFileName(self, 'Сохранить ежемесячный отчет', 'Ежемесячный отчет_'+last_month_s + '_' + str(last_month.year)+'.xlsx', '*xlsx')[0]
                if len(monthly_rep_month) != 0:
                    wb = xlsxwriter.Workbook(monthly_rep_month)
                    ws = wb.add_worksheet('Ежемесячный отчет')
                    ws2 = wb.add_worksheet('Анализ')
                    column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона','Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                    column2 = ['Доходы', 'Расходы', 'Сальдо']
                    column3 = ['Доля дохода', 'Доля расхода']
                    column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                    vagons = []
                    tab_vagon = []
                    monthly_rep_vagons = monthly_rep
                    monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                    for i in range(len(monthly_rep_vagons)):
                        vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                    for i in range(len(vagons)):
                        vagon_rep = monthly_rep[monthly_rep.VAGON == vagons[i]]
                        vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                        vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                        vagon_saldo = sum(vagon_rep.SALDO)
                        tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                    df_vagon = pd.DataFrame(tab_vagon)
                    df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                    df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)
                    
                    dohod = 0
                    rashod = 0
                    saldo = 0
                    for i in range(len(monthly_rep)):
                        dohod = dohod + int(monthly_rep.iloc[i]['TOTAL_DOHOD'])
                        rashod = rashod + int(monthly_rep.iloc[i]['TOTAL_RASHOD'])
                        saldo = saldo + int(monthly_rep.iloc[i]['SALDO'])
                        
                    arr_summ = [dohod, rashod, saldo]
                    arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]
                    
                    bold_form = wb.add_format()
                    bold_form.set_bold()
                    bold_form.set_align('center')
                    
                    date_form = wb.add_format()
                    date_form.set_num_format('dd.mm.yyyy')
                    
                    money_form  = wb.add_format()
                    money_form.set_num_format('# ##0')
                    
                    percent_form = wb.add_format()
                    percent_form.set_num_format('0.00%')

                    for i in range(len(column)):
                        ws.write(0, i, column[i], bold_form)
                        
                    for i in range(len(monthly_rep.index)):
                        for j in range(len(monthly_rep.columns)):
                            ws.write(i+1, j, monthly_rep.iloc[i][j])
                            
                    ws.set_column('A:A', 15, date_form)
                    ws.set_column('S:AC', 20, money_form)
                    ws.set_column('AF:AH', 20, money_form)
                    ws.set_column('AD:AD', 10, percent_form)
                    
                    for i in range(len(column)):
                        ws.set_column(0, i, 25)
                    
                    for i in range(len(column2)):
                        ws2.write(0, i, column2[i], bold_form)
                        
                    for i in range(len(arr_summ)):
                        ws2.write(1, i, arr_summ[i], money_form)
                        
                    for i in range(len(column3)):
                        ws2.write(3, i, column3[i], bold_form)
                    
                    for i in range(len(arr_share)):
                        ws2.write(4, i, arr_share[i], percent_form)
                        
                    for i in range(len(column4)):
                        ws2.write(6, i, column4[i], bold_form)
                        
                    for i in range(len(df_vagon.index)):
                        for j in range(len(df_vagon.columns)):
                            ws2.write(i+7, j, df_vagon.iloc[i][j])
                    
                    pie_chart = wb.add_chart({'type':'pie'})
                    pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                    pie_chart.set_title({'name':'Доля дохода и расхода'})
                    pie_chart.set_style(10)
                    ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})
                    
                    bar_chart = wb.add_chart({'type':'bar'})
                    bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                    bar_chart.set_title({'name':'Доход Расход Сальдо'})
                    bar_chart.set_style(10)
                    ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                    
                    bar_chart2 = wb.add_chart({'type':'bar'})
                    bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                    bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                    bar_chart2.set_style(10)
                    ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})
                    
                    wb.close()
        
    def QuarterReport(self):
        today = datetime.date.today()
        if (today.month >= 4)&(today.month <= 6):
            msg = QMessageBox.question(self, '', 'Вы точно хотите загрузить отчет на 1 квартал ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if (msg == QMessageBox.Yes):
                self.arhiv = pd.read_csv(self.arhivPath())
                for i in range(len(self.arhiv)):
                    self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
                last3month_oper = datetime.date(today.year, 2, 1)
                last2month_oper = datetime.date(today.year, 3, 1)
                last1month_oper = datetime.date(today.year, 4, 1)
                quarter_rep = self.arhiv[(self.arhiv.OPER_DATE == last3month_oper)|(self.arhiv.OPER_DATE == last2month_oper)|(self.arhiv.OPER_DATE == last1month_oper)]
                if (len(quarter_rep)) <= 0:
                    QMessageBox.information(self, 'Информация', 'В данном отрезке данных нет')
                if (len(quarter_rep)) > 0:
                    quarter_rep_path = QFileDialog.getSaveFileName(self, 'Сохранить ежеквартальный отчет', 'Ежеквартальный отчет_'+ 'за 1 квартал ' + str(today.year) +' года.xlsx', '*xlsx')[0]
                    if len(quarter_rep_path) != 0:
                        wb = xlsxwriter.Workbook(quarter_rep_path)
                        ws = wb.add_worksheet('Ежеквартальный отчет')
                        ws2 = wb.add_worksheet('Анализ')
                        column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона', 'Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                    'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                    'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                    'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                        column2 = ['Доходы', 'Расходы', 'Сальдо']
                        column3 = ['Доля дохода', 'Доля расхода']
                        column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                        vagons = []
                        tab_vagon = []
                        monthly_rep_vagons = quarter_rep
                        monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                        for i in range(len(monthly_rep_vagons)):
                            vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                        for i in range(len(vagons)):
                            vagon_rep = quarter_rep[quarter_rep.VAGON == vagons[i]]
                            vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                            vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                            vagon_saldo = sum(vagon_rep.SALDO)
                            tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                        df_vagon = pd.DataFrame(tab_vagon)
                        df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                        df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)

                        dohod = 0
                        rashod = 0
                        saldo = 0
                        for i in range(len(quarter_rep)):
                            dohod = dohod + int(quarter_rep.iloc[i]['TOTAL_DOHOD'])
                            rashod = rashod + int(quarter_rep.iloc[i]['TOTAL_RASHOD'])
                            saldo = saldo + int(quarter_rep.iloc[i]['SALDO'])

                        arr_summ = [dohod, rashod, saldo]
                        arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]

                        
                        bold_form = wb.add_format()
                        bold_form.set_bold()
                        bold_form.set_align('center')

                        date_form = wb.add_format()
                        date_form.set_num_format('dd.mm.yyyy')

                        money_form  = wb.add_format()
                        money_form.set_num_format('# ##0')

                        percent_form = wb.add_format()
                        percent_form.set_num_format('0.00%')

                        for i in range(len(column)):
                            ws.write(0, i, column[i], bold_form)

                        for i in range(len(quarter_rep.index)):
                            for j in range(len(quarter_rep.columns)):
                                ws.write(i+1, j, quarter_rep.iloc[i][j])

                        ws.set_column('A:A', 15, date_form)
                        ws.set_column('S:AC', 20, money_form)
                        ws.set_column('AF:AH', 20, money_form)
                        ws.set_column('AD:AD', 10, percent_form)

                        for i in range(len(column)):
                            ws.set_column(0, i, 25)

                        for i in range(len(column2)):
                            ws2.write(0, i, column2[i], bold_form)

                        for i in range(len(arr_summ)):
                            ws2.write(1, i, arr_summ[i], money_form)

                        for i in range(len(column3)):
                            ws2.write(3, i, column3[i], bold_form)

                        for i in range(len(arr_share)):
                            ws2.write(4, i, arr_share[i], percent_form)
                            
                        for i in range(len(column4)):
                            ws2.write(6, i, column4[i], bold_form)
                        
                        for i in range(len(df_vagon.index)):
                            for j in range(len(df_vagon.columns)):
                                ws2.write(i+7, j, df_vagon.iloc[i][j])
                    

                        pie_chart = wb.add_chart({'type':'pie'})
                        pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                        pie_chart.set_title({'name':'Доля дохода и расхода'})
                        pie_chart.set_style(10)
                        ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})

                        bar_chart = wb.add_chart({'type':'bar'})
                        bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                        bar_chart.set_title({'name':'Доход Расход Сальдо'})
                        bar_chart.set_style(10)
                        ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                        
                        bar_chart2 = wb.add_chart({'type':'bar'})
                        bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                        bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                        bar_chart2.set_style(10)
                        ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})


                        wb.close()
                        
                    
        if (today.month >= 7)&(today.month <= 9):
            msg = QMessageBox.question(self, '', 'Вы точно хотите загрузить отчет на 2 квартал ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if (msg == QMessageBox.Yes):
                self.arhiv = pd.read_csv(self.arhivPath())
                for i in range(len(self.arhiv)):
                    self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
                last3month_oper = datetime.date(today.year, 5, 1)
                last2month_oper = datetime.date(today.year, 6, 1)
                last1month_oper = datetime.date(today.year, 7, 1)
                quarter_rep = self.arhiv[(self.arhiv.OPER_DATE == last3month_oper)|(self.arhiv.OPER_DATE == last2month_oper)|(self.arhiv.OPER_DATE == last1month_oper)]
                if (len(quarter_rep)) <= 0:
                    QMessageBox.information(self, 'Информация', 'В данном отрезке данных нет')
                if (len(quarter_rep)) > 0:
                    quarter_rep_path = QFileDialog.getSaveFileName(self, 'Сохранить ежеквартальный отчет', 'Ежеквартальный отчет_'+ 'за 2 квартал ' + str(today.year) +' года.xlsx', '*xlsx')[0]
                    if len(quarter_rep_path) != 0:
                        wb = xlsxwriter.Workbook(quarter_rep_path)
                        ws = wb.add_worksheet('Ежеквартальный отчет')
                        ws2 = wb.add_worksheet('Анализ')
                        column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона', 'Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                    'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                    'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                    'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                        column2 = ['Доходы', 'Расходы', 'Сальдо']
                        column3 = ['Доля дохода', 'Доля расхода']
                        column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                        vagons = []
                        tab_vagon = []
                        monthly_rep_vagons = quarter_rep
                        monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                        for i in range(len(monthly_rep_vagons)):
                            vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                        for i in range(len(vagons)):
                            vagon_rep = quarter_rep[quarter_rep.VAGON == vagons[i]]
                            vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                            vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                            vagon_saldo = sum(vagon_rep.SALDO)
                            tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                        df_vagon = pd.DataFrame(tab_vagon)
                        df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                        df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)

                        dohod = 0
                        rashod = 0
                        saldo = 0
                        for i in range(len(quarter_rep)):
                            dohod = dohod + int(quarter_rep.iloc[i]['TOTAL_DOHOD'])
                            rashod = rashod + int(quarter_rep.iloc[i]['TOTAL_RASHOD'])
                            saldo = saldo + int(quarter_rep.iloc[i]['SALDO'])

                        arr_summ = [dohod, rashod, saldo]
                        arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]

                        bold_form = wb.add_format()
                        bold_form.set_bold()
                        bold_form.set_align('center')

                        date_form = wb.add_format()
                        date_form.set_num_format('dd.mm.yyyy')

                        money_form  = wb.add_format()
                        money_form.set_num_format('# ##0')

                        percent_form = wb.add_format()
                        percent_form.set_num_format('0.00%')

                        for i in range(len(column)):
                            ws.write(0, i, column[i], bold_form)

                        for i in range(len(quarter_rep.index)):
                            for j in range(len(quarter_rep.columns)):
                                ws.write(i+1, j, quarter_rep.iloc[i][j])

                        ws.set_column('A:A', 15, date_form)
                        ws.set_column('S:AC', 20, money_form)
                        ws.set_column('AF:AH', 20, money_form)
                        ws.set_column('AD:AD', 10, percent_form)

                        for i in range(len(column)):
                            ws.set_column(0, i, 25)

                        for i in range(len(column2)):
                            ws2.write(0, i, column2[i], bold_form)

                        for i in range(len(arr_summ)):
                            ws2.write(1, i, arr_summ[i], money_form)

                        for i in range(len(column3)):
                            ws2.write(3, i, column3[i], bold_form)

                        for i in range(len(arr_share)):
                            ws2.write(4, i, arr_share[i], percent_form)
                            
                        for i in range(len(column4)):
                            ws2.write(6, i, column4[i], bold_form)
                        
                        for i in range(len(df_vagon.index)):
                            for j in range(len(df_vagon.columns)):
                                ws2.write(i+7, j, df_vagon.iloc[i][j])
                    

                        pie_chart = wb.add_chart({'type':'pie'})
                        pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                        pie_chart.set_title({'name':'Доля дохода и расхода'})
                        pie_chart.set_style(10)
                        ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})

                        bar_chart = wb.add_chart({'type':'bar'})
                        bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                        bar_chart.set_title({'name':'Доход Расход Сальдо'})
                        bar_chart.set_style(10)
                        ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                        
                        bar_chart2 = wb.add_chart({'type':'bar'})
                        bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                        bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                        bar_chart2.set_style(10)
                        ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})

                        wb.close()
                        
            
        if (today.month >= 10)&(today.month <= 12):
            msg = QMessageBox.question(self, '', 'Вы точно хотите загрузить отчет на 3 квартал ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if (msg == QMessageBox.Yes):
                self.arhiv = pd.read_csv(self.arhivPath())
                for i in range(len(self.arhiv)):
                    self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
                last3month_oper = datetime.date(today.year, 8, 1)
                last2month_oper = datetime.date(today.year, 9, 1)
                last1month_oper = datetime.date(today.year, 10, 1)
                quarter_rep = self.arhiv[(self.arhiv.OPER_DATE == last3month_oper)|(self.arhiv.OPER_DATE == last2month_oper)|(self.arhiv.OPER_DATE == last1month_oper)]
                if (len(quarter_rep)) <= 0:
                    QMessageBox.information(self, 'Информация', 'В данном отрезке данных нет')
                if (len(quarter_rep)) > 0:
                    quarter_rep_path = QFileDialog.getSaveFileName(self, 'Сохранить ежеквартальный отчет', 'Ежеквартальный отчет_'+ 'за 3 квартал ' + str(today.year) +' года.xlsx', '*xlsx')[0]
                    if len(quarter_rep_path) != 0:
                        wb = xlsxwriter.Workbook(quarter_rep_path)
                        ws = wb.add_worksheet('Ежеквартальный отчет')
                        ws2 = wb.add_worksheet('Анализ')
                        column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона', 'Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                    'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                    'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                    'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                        column2 = ['Доходы', 'Расходы', 'Сальдо']
                        column3 = ['Доля дохода', 'Доля расхода']
                        column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                        vagons = []
                        tab_vagon = []
                        monthly_rep_vagons = quarter_rep
                        monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                        for i in range(len(monthly_rep_vagons)):
                            vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                        for i in range(len(vagons)):
                            vagon_rep = quarter_rep[quarter_rep.VAGON == vagons[i]]
                            vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                            vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                            vagon_saldo = sum(vagon_rep.SALDO)
                            tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                        df_vagon = pd.DataFrame(tab_vagon)
                        df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                        df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)

                        dohod = 0
                        rashod = 0
                        saldo = 0
                        for i in range(len(quarter_rep)):
                            dohod = dohod + int(quarter_rep.iloc[i]['TOTAL_DOHOD'])
                            rashod = rashod + int(quarter_rep.iloc[i]['TOTAL_RASHOD'])
                            saldo = saldo + int(quarter_rep.iloc[i]['SALDO'])

                        arr_summ = [dohod, rashod, saldo]
                        arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]

                        bold_form = wb.add_format()
                        bold_form.set_bold()
                        bold_form.set_align('center')

                        date_form = wb.add_format()
                        date_form.set_num_format('dd.mm.yyyy')

                        money_form  = wb.add_format()
                        money_form.set_num_format('# ##0')

                        percent_form = wb.add_format()
                        percent_form.set_num_format('0.00%')

                        for i in range(len(column)):
                            ws.write(0, i, column[i], bold_form)

                        for i in range(len(quarter_rep.index)):
                            for j in range(len(quarter_rep.columns)):
                                ws.write(i+1, j, quarter_rep.iloc[i][j])
                                
                        ws.set_column('A:A', 15, date_form)
                        ws.set_column('S:AC', 20, money_form)
                        ws.set_column('AF:AH', 20, money_form)
                        ws.set_column('AD:AD', 10, percent_form)

                        for i in range(len(column)):
                            ws.set_column(0, i, 25)

                        for i in range(len(column2)):
                            ws2.write(0, i, column2[i], bold_form)

                        for i in range(len(arr_summ)):
                            ws2.write(1, i, arr_summ[i], money_form)

                        for i in range(len(column3)):
                            ws2.write(3, i, column3[i], bold_form)

                        for i in range(len(arr_share)):
                            ws2.write(4, i, arr_share[i], percent_form)
                            
                        for i in range(len(column4)):
                            ws2.write(6, i, column4[i], bold_form)
                        
                        for i in range(len(df_vagon.index)):
                            for j in range(len(df_vagon.columns)):
                                ws2.write(i+7, j, df_vagon.iloc[i][j])
                    
                        pie_chart = wb.add_chart({'type':'pie'})
                        pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                        pie_chart.set_title({'name':'Доля дохода и расхода'})
                        pie_chart.set_style(10)
                        ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})

                        bar_chart = wb.add_chart({'type':'bar'})
                        bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                        bar_chart.set_title({'name':'Доход Расход Сальдо'})
                        bar_chart.set_style(10)
                        ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                        
                        bar_chart2 = wb.add_chart({'type':'bar'})
                        bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                        bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                        bar_chart2.set_style(10)
                        ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})

                        wb.close()
                        
                        
        if (today.month >= 1)&(today.month <= 3):
            msg = QMessageBox.question(self, '', 'Вы точно хотите загрузить отчет на 4 квартал ' + str(today.year-1) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if (msg == QMessageBox.Yes):
                self.arhiv = pd.read_csv(self.arhivPath())
                for i in range(len(self.arhiv)):
                    self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
                last3month_oper = datetime.date(today.year-1, 11, 1)
                last2month_oper = datetime.date(today.year-1, 12, 1)
                last1month_oper = datetime.date(today.year, 1, 1)
                quarter_rep = self.arhiv[(self.arhiv.OPER_DATE == last3month_oper)|(self.arhiv.OPER_DATE == last2month_oper)|(self.arhiv.OPER_DATE == last1month_oper)]
                if (len(quarter_rep)) <= 0:
                    QMessageBox.information(self, 'Информация', 'В данном отрезке данных нет')
                if (len(quarter_rep)) > 0:
                    quarter_rep_path = QFileDialog.getSaveFileName(self, 'Сохранить ежеквартальный отчет', 'Ежеквартальный отчет_'+ 'за 4 квартал ' + str(today.year-1) +' года.xlsx', '*xlsx')[0]
                    if len(quarter_rep_path) != 0:
                        wb = xlsxwriter.Workbook(quarter_rep_path)
                        ws = wb.add_worksheet('Ежеквартальный отчет')
                        ws2 = wb.add_worksheet('Анализ')
                        column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона', 'Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                    'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                    'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                    'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                        column2 = ['Доходы', 'Расходы', 'Сальдо']
                        column3 = ['Доля дохода', 'Доля расхода']
                        column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                        vagons = []
                        tab_vagon = []
                        monthly_rep_vagons = quarter_rep
                        monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                        for i in range(len(monthly_rep_vagons)):
                            vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                        for i in range(len(vagons)):
                            vagon_rep = quarter_rep[quarter_rep.VAGON == vagons[i]]
                            vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                            vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                            vagon_saldo = sum(vagon_rep.SALDO)
                            tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                        df_vagon = pd.DataFrame(tab_vagon)
                        df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                        df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)
                        
                        dohod = 0
                        rashod = 0
                        saldo = 0
                        for i in range(len(quarter_rep)):
                            dohod = dohod + int(quarter_rep.iloc[i]['TOTAL_DOHOD'])
                            rashod = rashod + int(quarter_rep.iloc[i]['TOTAL_RASHOD'])
                            saldo = saldo + int(quarter_rep.iloc[i]['SALDO'])

                        arr_summ = [dohod, rashod, saldo]
                        arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]

                        bold_form = wb.add_format()
                        bold_form.set_bold()
                        bold_form.set_align('center')

                        date_form = wb.add_format()
                        date_form.set_num_format('dd.mm.yyyy')

                        money_form  = wb.add_format()
                        money_form.set_num_format('# ##0')

                        percent_form = wb.add_format()
                        percent_form.set_num_format('0.00%')

                        for i in range(len(column)):
                            ws.write(0, i, column[i], bold_form)

                        for i in range(len(quarter_rep.index)):
                            for j in range(len(quarter_rep.columns)):
                                ws.write(i+1, j, quarter_rep.iloc[i][j])

                        ws.set_column('A:A', 15, date_form)
                        ws.set_column('S:AC', 20, money_form)
                        ws.set_column('AF:AH', 20, money_form)
                        ws.set_column('AD:AD', 10, percent_form)

                        for i in range(len(column)):
                            ws.set_column(0, i, 25)

                        for i in range(len(column2)):
                            ws2.write(0, i, column2[i], bold_form)

                        for i in range(len(arr_summ)):
                            ws2.write(1, i, arr_summ[i], money_form)

                        for i in range(len(column3)):
                            ws2.write(3, i, column3[i], bold_form)

                        for i in range(len(arr_share)):
                            ws2.write(4, i, arr_share[i], percent_form)
                            
                        for i in range(len(column4)):
                            ws2.write(6, i, column4[i], bold_form)
                        
                        for i in range(len(df_vagon.index)):
                            for j in range(len(df_vagon.columns)):
                                ws2.write(i+7, j, df_vagon.iloc[i][j])

                        pie_chart = wb.add_chart({'type':'pie'})
                        pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                        pie_chart.set_title({'name':'Доля дохода и расхода'})
                        pie_chart.set_style(10)
                        ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})

                        bar_chart = wb.add_chart({'type':'bar'})
                        bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                        bar_chart.set_title({'name':'Доход Расход Сальдо'})
                        bar_chart.set_style(10)
                        ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                        
                        bar_chart2 = wb.add_chart({'type':'bar'})
                        bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                        bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                        bar_chart2.set_style(10)
                        ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})

                        wb.close()
                        
    def YearlyReport(self):
        today = datetime.date.today()
        self.arhiv = pd.read_csv(self.arhivPath())
        for i in range(len(self.arhiv)):
            self.arhiv.at[self.arhiv.index[i], 'OPER_DATE'] = datetime.date(int(self.arhiv.iloc[i]['OPER_DATE'][0:4]), int(self.arhiv.iloc[i]['OPER_DATE'][5:7]), int(self.arhiv.iloc[i]['OPER_DATE'][8:10]))
        year_rep = self.arhiv[(self.arhiv.OPER_DATE >= datetime.date(today.year-1, 2, 1))&(self.arhiv.OPER_DATE <= datetime.date(today.year, 1, 1))]
        if len(year_rep) > 0:
            msg = QMessageBox.question(self, 'Вопрос', 'Вы точно хотите загрузить ежегодный отчет за ' + str(today.year-1) + ' год?', QMessageBox.Yes|QMessageBox.No)
            if (msg == QMessageBox.Yes):
                year_rep_path = QFileDialog.getSaveFileName(self, 'Сохранить ежегодный отчет', 'Ежегодный отчет ' + str(today.year-1) + ' год.xlsx', '*.xlsx')[0]
                if len(year_rep_path) != 0:
                    wb = xlsxwriter.Workbook(year_rep_path)
                    ws = wb.add_worksheet('Ежегодный отчет')
                    ws2 = wb.add_worksheet('Анализ')
                    column = ['Дата отчета', 'ID номер', 'Дата добавления', '№ Вагона', 'Пункт отправления', 'Пункт назначения', 'Дата отправления', 'Дата назначения', 'Расстояние (км)', 'Кол-во дней в одну сторону', 
                                    'Кол-во дней в обратную сторону', 'Кол-во дней погрузки', 'Кол-во дней выгрузки', 'Дата погрузки', 'Дата выгрузки', 'Дата входа не ремонт', 'Дата выхода из ремонта',
                                    'Кол-во дней в ремонте', 'Оплата аренды','Затраты на ремонт', 'ППС', 'Груженый тариф', 'По розней 1 территории', 'По розней 2 территории', 'За ускорение 1 территории', 
                                    'За ускорение 2 территории', 'Телеграммы', 'Прочие расходы', 'Приходы', 'Коэффициент', 'Кол-во дней в пути', 'Общий расход', 'Общий доход', 'Сальдо']
                    column2 = ['Доходы', 'Расходы', 'Сальдо']
                    column3 = ['Доля дохода', 'Доля расхода']
                    column4 = ['№ Вагона', 'Доходы', 'Расходы', 'Сальдо']
                    vagons = []
                    tab_vagon = []
                    monthly_rep_vagons = year_rep
                    monthly_rep_vagons = monthly_rep_vagons.drop_duplicates('VAGON', keep='last')
                    for i in range(len(monthly_rep_vagons)):
                        vagons.append(monthly_rep_vagons.iloc[i]['VAGON'])
                    for i in range(len(vagons)):
                        vagon_rep = year_rep[year_rep.VAGON == vagons[i]]
                        vagon_dohod = sum(vagon_rep.TOTAL_DOHOD)
                        vagon_rashod = sum(vagon_rep.TOTAL_RASHOD)
                        vagon_saldo = sum(vagon_rep.SALDO)
                        tab_vagon.append({'Vagon':vagons[i], 'Dohod':vagon_dohod, 'Rashod':vagon_rashod, 'Saldo':vagon_saldo})
                    df_vagon = pd.DataFrame(tab_vagon)
                    df_vagon = df_vagon.reindex(['Vagon', 'Dohod', 'Rashod', 'Saldo'], axis=1)
                    df_vagon = df_vagon.sort_values(by='Saldo', ascending=False)

                    dohod = 0
                    rashod = 0
                    saldo = 0
                    for i in range(len(year_rep)):
                        dohod = dohod + int(year_rep.iloc[i]['TOTAL_DOHOD'])
                        rashod = rashod + int(year_rep.iloc[i]['TOTAL_RASHOD'])
                        saldo = saldo + int(year_rep.iloc[i]['SALDO'])

                    arr_summ = [dohod, rashod, saldo]
                    arr_share = [dohod/(dohod + rashod), rashod/(dohod + rashod)]

                    bold_form = wb.add_format()
                    bold_form.set_bold()
                    bold_form.set_align('center')

                    date_form = wb.add_format()
                    date_form.set_num_format('dd.mm.yyyy')

                    money_form  = wb.add_format()
                    money_form.set_num_format('# ##0')

                    percent_form = wb.add_format()
                    percent_form.set_num_format('0.00%')

                    for i in range(len(column)):
                        ws.write(0, i, column[i], bold_form)

                    for i in range(len(year_rep.index)):
                        for j in range(len(year_rep.columns)):
                            ws.write(i+1, j, year_rep.iloc[i][j])

                    ws.set_column('A:A', 15, date_form)
                    ws.set_column('S:AC', 20, money_form)
                    ws.set_column('AF:AH', 20, money_form)
                    ws.set_column('AD:AD', 10, percent_form)

                    for i in range(len(column)):
                        ws.set_column(0, i, 25)

                    for i in range(len(column2)):
                        ws2.write(0, i, column2[i], bold_form)

                    for i in range(len(arr_summ)):
                        ws2.write(1, i, arr_summ[i], money_form)

                    for i in range(len(column3)):
                        ws2.write(3, i, column3[i], bold_form)

                    for i in range(len(arr_share)):
                        ws2.write(4, i, arr_share[i], percent_form)
                        
                    for i in range(len(column4)):
                        ws2.write(6, i, column4[i], bold_form)
                        
                    for i in range(len(df_vagon.index)):
                        for j in range(len(df_vagon.columns)):
                            ws2.write(i+7, j, df_vagon.iloc[i][j])

                    pie_chart = wb.add_chart({'type':'pie'})
                    pie_chart.add_series({'categories':['Анализ', 3, 0, 3, 1], 'values':['Анализ', 4, 0, 4, 1]})
                    pie_chart.set_title({'name':'Доля дохода и расхода'})
                    pie_chart.set_style(10)
                    ws2.insert_chart('E1', pie_chart, {'x_offset':25, 'y_offset':10})

                    bar_chart = wb.add_chart({'type':'bar'})
                    bar_chart.add_series({'categories':['Анализ', 0, 0, 0, 2], 'values':['Анализ', 1, 0, 1, 2]})
                    bar_chart.set_title({'name':'Доход Расход Сальдо'})
                    bar_chart.set_style(10)
                    ws2.insert_chart('M1', bar_chart, {'x_offset':25, 'y_offset':10})
                    
                    bar_chart2 = wb.add_chart({'type':'bar'})
                    bar_chart2.add_series({'categories':['Анализ', 7, 0, 6 + len(df_vagon), 0], 'values':['Анализ', 7, 3, 6 + len(df_vagon), 3]})
                    bar_chart2.set_title({'name':'Сальдо по Вагонам'})
                    bar_chart2.set_style(10)
                    ws2.insert_chart('E17', bar_chart2, {'x_offset':25, 'y_offset':10})

                    wb.close()
        else:
            QMessageBox.information(self, 'Информация', 'Нет данных за ' + str(today.year-1) + ' год.')
       
    
    def MonthArhiv(self):
        today = datetime.date.today()
        last_month = today + relativedelta(months=-1)
        if last_month.month == 1:
            last_month_s = 'Январь'
        if last_month.month == 2:
            last_month_s = 'Февраль'
        if last_month.month == 3:
            last_month_s = 'Март'
        if last_month.month == 4:
            last_month_s = 'Апрель'
        if last_month.month == 5:
            last_month_s = 'Май'
        if last_month.month == 6:
            last_month_s = 'Июнь'
        if last_month.month == 7:
            last_month_s = 'Июль'
        if last_month.month == 8:
            last_month_s = 'Август'
        if last_month.month == 9:
            last_month_s = 'Сентябрь'
        if last_month.month == 10:
            last_month_s = 'Октябрь'
        if last_month.month == 11:
            last_month_s = 'Ноябрь'
        if last_month.month == 12:
            last_month_s = 'Декабрь'
            
        msg = QMessageBox.question(self, '', 'Вы точно хотите архивировать данные за '+ last_month_s + ' ' + str(last_month.year) + '?', QMessageBox.Yes|QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.data = pd.read_csv(self.dataPath())
            ids = []
            if len(self.data) <= 0:
                QMessageBox.information(self, 'Информация', 'Монитор пуст')
            else:
                for i in range(len(self.data)):
                    date_path = self.data.iloc[i]['DATE_PATH']
                    date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                    if (date_path_date >= datetime.date(last_month.year, last_month.month, 1))&(date_path_date <= datetime.date(last_month.year, last_month.month, calendar.monthrange(last_month.year, last_month.month)[1])):
                        ids.append(self.data.iloc[i]['ID'])
                        
                self.data = self.data[~self.data.ID.isin(ids)]
                self.data = self.data.reset_index(drop=True)
                
                self.data.to_csv(self.dataPath(), index=False)
                self.data = pd.read_csv(self.dataPath())
                self.nd = NewDialog(self)
                self.nd.setGeometry(20, 20, 1460, 500)
                self.nd.tableWidget = QTableWidget()
                self.nd.tableWidget.setRowCount(len(self.data.index))
                self.nd.tableWidget.setColumnCount(len(self.data.columns))
                for i in range(len(self.data.index)):
                    for j in range(len(self.data.columns)):
                        self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                self.nd.tableWidget.move(20,20)
                header = self.Header()
                self.nd.tableWidget.setHorizontalHeaderLabels(header)
                self.nd.tableWidget.resizeColumnsToContents()
                self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                self.nd.layout = QVBoxLayout()
                self.nd.layout.addWidget(self.nd.tableWidget) 
                self.nd.setLayout(self.nd.layout)
                self.nd.show()
                
                QMessageBox.information(self, 'Информация', 'Архивация завершена')
                
    def QuarterArhiv(self):
        today = datetime.date.today()
        if (today.month >= 1)&(today.month <= 3):
            msg = QMessageBox.question(self, 'Вопрос', 'Вы точно хотите архивировать данные 4 квартала ' + str(today.year-1) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if msg == QMessageBox.Yes:
                self.data = pd.read_csv(self.dataPath())
                ids = []
                if len(self.data) <= 0:
                    QMessageBox.information(self, 'Информация', 'Монитор пуст')
                else:
                    for i in range(len(self.data)):
                        date_path = self.data.iloc[i]['DATE_PATH']
                        date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        last_1month_min = datetime.date(today.year-1, 12, 1)
                        last_1month_max = datetime.date(today.year-1, 12, calendar.monthrange(today.year-1, 12)[1])
                        last_2month_min = datetime.date(today.year-1, 11, 1)
                        last_2month_max = datetime.date(today.year-1, 11, calendar.monthrange(today.year-1, 11)[1])
                        last_3month_min = datetime.date(today.year-1, 10, 1)
                        last_3month_max = datetime.date(today.year-1, 10, calendar.monthrange(today.year-1, 10)[1])
                        if ((date_path_date >= last_1month_min)&(date_path_date <= last_1month_max))|((date_path_date >= last_2month_min)&(date_path_date <= last_2month_max))|((date_path_date >= last_3month_min)&(date_path_date <= last_3month_max)):
                            ids.append(self.data.iloc[i]['ID'])
                            
                    self.data = self.data[~self.data.ID.isin(ids)]
                    self.data = self.data.reset_index(drop=True)
                
                    self.data.to_csv(self.dataPath(), index=False)
                    self.data = pd.read_csv(self.dataPath())
                    self.nd = NewDialog(self)
                    self.nd.setGeometry(20, 20, 1460, 500)
                    self.nd.tableWidget = QTableWidget()
                    self.nd.tableWidget.setRowCount(len(self.data.index))
                    self.nd.tableWidget.setColumnCount(len(self.data.columns))
                    for i in range(len(self.data.index)):
                        for j in range(len(self.data.columns)):
                            self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                    self.nd.tableWidget.move(20,20)
                    header = self.Header()
                    self.nd.tableWidget.setHorizontalHeaderLabels(header)
                    self.nd.tableWidget.resizeColumnsToContents()
                    self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                    self.nd.layout = QVBoxLayout()
                    self.nd.layout.addWidget(self.nd.tableWidget) 
                    self.nd.setLayout(self.nd.layout)
                    self.nd.show()

                    QMessageBox.information(self, 'Информация', 'Архивация завершена')
                    
        if (today.month >= 4)&(today.month <= 6):
            msg = QMessageBox.question(self, 'Вопрос', 'Вы точно хотите архивировать данные 1 квартала ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if msg == QMessageBox.Yes:
                self.data = pd.read_csv(self.dataPath())
                ids = []
                if len(self.data) <= 0:
                    QMessageBox.information(self, 'Информация', 'Монитор пуст')
                else:
                    for i in range(len(self.data)):
                        date_path = self.data.iloc[i]['DATE_PATH']
                        date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        last_1month_min = datetime.date(today.year, 3, 1)
                        last_1month_max = datetime.date(today.year, 3, calendar.monthrange(today.year, 3)[1])
                        last_2month_min = datetime.date(today.year, 2, 1)
                        last_2month_max = datetime.date(today.year, 2, calendar.monthrange(today.year, 2)[1])
                        last_3month_min = datetime.date(today.year, 1, 1)
                        last_3month_max = datetime.date(today.year, 1, calendar.monthrange(today.year, 1)[1])
                        if ((date_path_date >= last_1month_min)&(date_path_date <= last_1month_max))|((date_path_date >= last_2month_min)&(date_path_date <= last_2month_max))|((date_path_date >= last_3month_min)&(date_path_date <= last_3month_max)):
                            ids.append(self.data.iloc[i]['ID'])
                            
                    self.data = self.data[~self.data.ID.isin(ids)]
                    self.data = self.data.reset_index(drop=True)
                
                    self.data.to_csv(self.dataPath(), index=False)
                    self.data = pd.read_csv(self.dataPath())
                    self.nd = NewDialog(self)
                    self.nd.setGeometry(20, 20, 1460, 500)
                    self.nd.tableWidget = QTableWidget()
                    self.nd.tableWidget.setRowCount(len(self.data.index))
                    self.nd.tableWidget.setColumnCount(len(self.data.columns))
                    for i in range(len(self.data.index)):
                        for j in range(len(self.data.columns)):
                            self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                    self.nd.tableWidget.move(20,20)
                    header = self.Header()
                    self.nd.tableWidget.setHorizontalHeaderLabels(header)
                    self.nd.tableWidget.resizeColumnsToContents()
                    self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                    self.nd.layout = QVBoxLayout()
                    self.nd.layout.addWidget(self.nd.tableWidget) 
                    self.nd.setLayout(self.nd.layout)
                    self.nd.show()

                    QMessageBox.information(self, 'Информация', 'Архивация завершена')
                    
        if (today.month >= 7)&(today.month <= 9):
            msg = QMessageBox.question(self, 'Вопрос', 'Вы точно хотите архивировать данные 2 квартала ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if msg == QMessageBox.Yes:
                self.data = pd.read_csv(self.dataPath())
                ids = []
                if len(self.data) <= 0:
                    QMessageBox.information(self, 'Информация', 'Монитор пуст')
                else:
                    for i in range(len(self.data)):
                        date_path = self.data.iloc[i]['DATE_PATH']
                        date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        last_1month_min = datetime.date(today.year, 6, 1)
                        last_1month_max = datetime.date(today.year, 6, calendar.monthrange(today.year, 6)[1])
                        last_2month_min = datetime.date(today.year, 5, 1)
                        last_2month_max = datetime.date(today.year, 5, calendar.monthrange(today.year, 5)[1])
                        last_3month_min = datetime.date(today.year, 4, 1)
                        last_3month_max = datetime.date(today.year, 4, calendar.monthrange(today.year, 4)[1])
                        if ((date_path_date >= last_1month_min)&(date_path_date <= last_1month_max))|((date_path_date >= last_2month_min)&(date_path_date <= last_2month_max))|((date_path_date >= last_3month_min)&(date_path_date <= last_3month_max)):
                            ids.append(self.data.iloc[i]['ID'])
                            
                    self.data = self.data[~self.data.ID.isin(ids)]
                    self.data = self.data.reset_index(drop=True)
                
                    self.data.to_csv(self.dataPath(), index=False)
                    self.data = pd.read_csv(self.dataPath())
                    self.nd = NewDialog(self)
                    self.nd.setGeometry(20, 20, 1460, 500)
                    self.nd.tableWidget = QTableWidget()
                    self.nd.tableWidget.setRowCount(len(self.data.index))
                    self.nd.tableWidget.setColumnCount(len(self.data.columns))
                    for i in range(len(self.data.index)):
                        for j in range(len(self.data.columns)):
                            self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                    self.nd.tableWidget.move(20,20)
                    header = self.Header()
                    self.nd.tableWidget.setHorizontalHeaderLabels(header)
                    self.nd.tableWidget.resizeColumnsToContents()
                    self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                    self.nd.layout = QVBoxLayout()
                    self.nd.layout.addWidget(self.nd.tableWidget) 
                    self.nd.setLayout(self.nd.layout)
                    self.nd.show()

                    QMessageBox.information(self, 'Информация', 'Архивация завершена')
                    
        if (today.month >= 10)&(today.month <= 12):
            msg = QMessageBox.question(self, 'Вопрос', 'Вы точно хотите архивировать данные 3 квартала ' + str(today.year) + ' года?', QMessageBox.Yes|QMessageBox.No)
            if msg == QMessageBox.Yes:
                self.data = pd.read_csv(self.dataPath())
                ids = []
                if len(self.data) <= 0:
                    QMessageBox.information(self, 'Информация', 'Монитор пуст')
                else:
                    for i in range(len(self.data)):
                        date_path = self.data.iloc[i]['DATE_PATH']
                        date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                        last_1month_min = datetime.date(today.year, 9, 1)
                        last_1month_max = datetime.date(today.year, 9, calendar.monthrange(today.year, 9)[1])
                        last_2month_min = datetime.date(today.year, 8, 1)
                        last_2month_max = datetime.date(today.year, 8, calendar.monthrange(today.year, 8)[1])
                        last_3month_min = datetime.date(today.year, 7, 1)
                        last_3month_max = datetime.date(today.year, 7, calendar.monthrange(today.year, 7)[1])
                        if ((date_path_date >= last_1month_min)&(date_path_date <= last_1month_max))|((date_path_date >= last_2month_min)&(date_path_date <= last_2month_max))|((date_path_date >= last_3month_min)&(date_path_date <= last_3month_max)):
                            ids.append(self.data.iloc[i]['ID'])
                            
                    self.data = self.data[~self.data.ID.isin(ids)]
                    self.data = self.data.reset_index(drop=True)
                
                    self.data.to_csv(self.dataPath(), index=False)
                    self.data = pd.read_csv(self.dataPath())
                    self.nd = NewDialog(self)
                    self.nd.setGeometry(20, 20, 1460, 500)
                    self.nd.tableWidget = QTableWidget()
                    self.nd.tableWidget.setRowCount(len(self.data.index))
                    self.nd.tableWidget.setColumnCount(len(self.data.columns))
                    for i in range(len(self.data.index)):
                        for j in range(len(self.data.columns)):
                            self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                    self.nd.tableWidget.move(20,20)
                    header = self.Header()
                    self.nd.tableWidget.setHorizontalHeaderLabels(header)
                    self.nd.tableWidget.resizeColumnsToContents()
                    self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                    self.nd.layout = QVBoxLayout()
                    self.nd.layout.addWidget(self.nd.tableWidget) 
                    self.nd.setLayout(self.nd.layout)
                    self.nd.show()

                    QMessageBox.information(self, 'Информация', 'Архивация завершена')
                    
    def YearArhiv(self):
        today = datetime.date.today()
        msg = QMessageBox.question(self, 'Вопрос', 'Вы точно архивировать данные за ' + str(today.year-1) + ' года?', QMessageBox.Yes|QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.data = pd.read_csv(self.dataPath())
            ids = []
            if len(self.data) <= 0:
                QMessageBox.information(self, 'Информация', 'Монитор пуст')
            else:
                for i in range(len(self.data)):
                    date_path = self.data.iloc[i]['DATE_PATH']
                    date_path_date = datetime.date(int(date_path[6:10]), int(date_path[3:5]), int(date_path[0:2]))
                    min_date = datetime.date(today.year-1, 1, 1)
                    max_date = datetime.date(today.year-1, 12, 31)
                    if (date_path_date >= min_date)&(date_path_date <= max_date):
                        ids.append(self.data.iloc[i]['ID'])
                
                self.data = self.data[~self.data.ID.isin(ids)]
                self.data = self.data.reset_index(drop=True)
                
                self.data.to_csv(self.dataPath(), index=False)
                self.data = pd.read_csv(self.dataPath())
                self.nd = NewDialog(self)
                self.nd.setGeometry(20, 20, 1460, 500)
                self.nd.tableWidget = QTableWidget()
                self.nd.tableWidget.setRowCount(len(self.data.index))
                self.nd.tableWidget.setColumnCount(len(self.data.columns))
                for i in range(len(self.data.index)):
                    for j in range(len(self.data.columns)):
                        self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))
                self.nd.tableWidget.move(20,20)
                header = self.Header()
                self.nd.tableWidget.setHorizontalHeaderLabels(header)
                self.nd.tableWidget.resizeColumnsToContents()
                self.nd.tableWidget.cellClicked.connect(self.ClickedCell)
                self.nd.layout = QVBoxLayout()
                self.nd.layout.addWidget(self.nd.tableWidget) 
                self.nd.setLayout(self.nd.layout)
                self.nd.show()

                QMessageBox.information(self, 'Информация', 'Архивация завершена')
        
    def FirstReport(self):
        self.data = pd.read_csv(self.dataPath())
        self.fd = FigDialog(self)
        self.fd.setGeometry(1070, 530, 800, 450)
        tab = []
        isAvail = False
        if (self.line_from.text() == '')&(self.line_till.text() == ''):
            self.data = self.data
            isAvail = True
            
        if (self.line_from.text() != '')&(self.line_till.text() == ''):
            int_from = int(self.line_from.text()) - 1
            if (int_from <= len(self.data))&(int_from >= 0):
                self.data = self.data[self.data.index >= int_from]
                isAvail = True
                
        if (self.line_from.text() == '')&(self.line_till.text() != ''):
            int_till = int(self.line_till.text()) - 1
            if (int_till <= len(self.data))&(int_till >= 0):
                self.data = self.data[self.data.index <= int_till]
                isAvail = True
                
        if (self.line_from.text() != '')&(self.line_till.text() != ''):
            int_from = int(self.line_from.text()) - 1
            int_till = int(self.line_till.text()) - 1
            if (int_from >= 0)&(int_from <= len(self.data))&(int_till <= len(self.data))&(int_till >= 0)&(int_from <= int_till):
                self.data = self.data[(self.data.index >= int_from)&(self.data.index <= int_till)]
                isAvail = True
        
        if (len(self.data) > 0)&(isAvail):
            for i in range(len(self.data)):
                tab.append({'Точка А-Б':self.data.index[i] + 1, 'Расходы':int(self.data.iloc[i]['TOTAL_RASHOD']), 'Доходы':int(self.data.iloc[i]['DOHOD']) + int(self.data.iloc[i]['PPS']), 'Сальдо':int(self.data.iloc[i]['SALDO'])})

            df_tab = pd.DataFrame(tab)
            f, ax = plt.subplots()
            plt.tick_params(labelsize=8)
            ax=sns.pointplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Расходы'],data=df_tab,color='red', scale=0.3)
            ax=sns.pointplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Доходы'],data=df_tab,color='blue', scale=0.3)
            ax = sns.pointplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Сальдо'],data=df_tab,color='green', scale=0.3)
            ax.set_xlabel('')
            ax.set_ylabel('')
            ax.legend(handles=ax.lines[::len(df_tab)+1], labels=["Расходы", "Доходы", "Сальдо"], ncol=3, loc=9, fontsize=8)
            figure = plt.figure()
            canvas = FigureCanvas(f)
            self.fd.lay = QVBoxLayout()
            self.fd.lay.addWidget(canvas)
            self.fd.setLayout(self.fd.lay)
            self.fd.show()
            
        else:
            QMessageBox.information(self, 'Ошибка', 'Данные пусты или некорректный анализируемый период')
    
    def SecondReport(self):
        self.data = pd.read_csv(self.dataPath())
        self.fd = FigDialog(self)
        self.fd.setGeometry(1070, 530, 800, 450)
        tab = []
        
        isAvail = False
        if (self.line_from.text() == '')&(self.line_till.text() == ''):
            self.data = self.data
            isAvail = True
            
        if (self.line_from.text() != '')&(self.line_till.text() == ''):
            int_from = int(self.line_from.text()) - 1
            if (int_from <= len(self.data))&(int_from >= 0):
                self.data = self.data[self.data.index >= int_from]
                isAvail = True
                
        if (self.line_from.text() == '')&(self.line_till.text() != ''):
            int_till = int(self.line_till.text()) - 1
            if (int_till <= len(self.data))&(int_till >= 0):
                self.data = self.data[self.data.index <= int_till]
                isAvail = True
                
        if (self.line_from.text() != '')&(self.line_till.text() != ''):
            int_from = int(self.line_from.text()) - 1
            int_till = int(self.line_till.text()) - 1
            if (int_from >= 0)&(int_from <= len(self.data))&(int_till <= len(self.data))&(int_till >= 0)&(int_from <= int_till):
                self.data = self.data[(self.data.index >= int_from)&(self.data.index <= int_till)]
                isAvail = True
                
        if (len(self.data) > 0)&(isAvail):
            for i in range(len(self.data)):
                doh = int(self.data.iloc[i]['DOHOD']) + int(self.data.iloc[i]['PPS'])
                ras = int(self.data.iloc[i]['TOTAL_RASHOD'])
                tab.append({'Точка А-Б':self.data.index[i] + 1, 'Расходы':round((ras/(ras + doh))*100, 1), 'Доходы':round((doh/(ras + doh))*100, 1)})

            df_tab = pd.DataFrame(tab)
            f, ax = plt.subplots()
            plt.tick_params(labelsize=8)
            ax=sns.pointplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Расходы'],data=df_tab,color='red', scale=0.3)
            ax=sns.pointplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Доходы'],data=df_tab,color='blue', scale=0.3)
            ax.set_xlabel('')
            ax.set_ylabel('')
            ax.legend(handles=ax.lines[::len(df_tab)+1], labels=["Расходы", "Доходы"], ncol=2, loc=9, fontsize=8)
            figure = plt.figure()
            canvas = FigureCanvas(f)
            self.fd.lay = QVBoxLayout()
            self.fd.lay.addWidget(canvas)
            self.fd.setLayout(self.fd.lay)
            self.fd.show()
        else:
            QMessageBox.information(self, 'Ошибка', 'Данные пусты или некорректный анализируемый период')
            
    def ThirdReport(self):
        self.data = pd.read_csv(self.dataPath())
        self.fd = FigDialog(self)
        self.fd.setGeometry(1070, 530, 800, 450)
        tab = []
        
        isAvail = False
        if (self.line_from.text() == '')&(self.line_till.text() == ''):
            self.data = self.data
            isAvail = True
            
        if (self.line_from.text() != '')&(self.line_till.text() == ''):
            int_from = int(self.line_from.text()) - 1
            if (int_from <= len(self.data))&(int_from >= 0):
                self.data = self.data[self.data.index >= int_from]
                isAvail = True
                
        if (self.line_from.text() == '')&(self.line_till.text() != ''):
            int_till = int(self.line_till.text()) - 1
            if (int_till <= len(self.data))&(int_till >= 0):
                self.data = self.data[self.data.index <= int_till]
                isAvail = True
                
        if (self.line_from.text() != '')&(self.line_till.text() != ''):
            int_from = int(self.line_from.text()) - 1
            int_till = int(self.line_till.text()) - 1
            if (int_from >= 0)&(int_from <= len(self.data))&(int_till <= len(self.data))&(int_till >= 0)&(int_from <= int_till):
                self.data = self.data[(self.data.index >= int_from)&(self.data.index <= int_till)]
                isAvail = True
        
        if (len(self.data) > 0)&(isAvail):
            for i in range(len(self.data)):
                doh = int(self.data.iloc[i]['DOHOD']) + int(self.data.iloc[i]['PPS'])
                tab.append({'Точка А-Б':self.data.index[i]+1, 'Доходы':doh})
                
            df_tab = pd.DataFrame(tab)
            f, ax = plt.subplots()
            plt.tick_params(labelsize=8)
            ax=sns.barplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Доходы'],data=df_tab,color='blue')
            ax.set_xlabel('')
            ax.set_ylabel('')
            #ax.legend(handles=ax.lines[::len(df_tab)+1], labels=["Доходы"], ncol=1, loc=9, fontsize=8)
            figure = plt.figure()
            canvas = FigureCanvas(f)
            self.fd.lay = QVBoxLayout()
            self.fd.lay.addWidget(canvas)
            self.fd.setLayout(self.fd.lay)
            self.fd.show()
        else:
            QMessageBox.information(self, 'Ошибка', 'Данные пусты или некорректный анализируемый период')
            
    def FourthReport(self):
        self.data = pd.read_csv(self.dataPath())
        self.fd = FigDialog(self)
        self.fd.setGeometry(1070, 530, 800, 450)
        tab = []
        isAvail = False
        if (self.line_from.text() == '')&(self.line_till.text() == ''):
            self.data = self.data
            isAvail = True
            
        if (self.line_from.text() != '')&(self.line_till.text() == ''):
            int_from = int(self.line_from.text()) - 1
            if (int_from <= len(self.data))&(int_from >= 0):
                self.data = self.data[self.data.index >= int_from]
                isAvail = True
                
        if (self.line_from.text() == '')&(self.line_till.text() != ''):
            int_till = int(self.line_till.text()) - 1
            if (int_till <= len(self.data))&(int_till >= 0):
                self.data = self.data[self.data.index <= int_till]
                isAvail = True
                
        if (self.line_from.text() != '')&(self.line_till.text() != ''):
            int_from = int(self.line_from.text()) - 1
            int_till = int(self.line_till.text()) - 1
            if (int_from >= 0)&(int_from <= len(self.data))&(int_till <= len(self.data))&(int_till >= 0)&(int_from <= int_till):
                self.data = self.data[(self.data.index >= int_from)&(self.data.index <= int_till)]
                isAvail = True
        
        if (len(self.data) > 0)&(isAvail):
            for i in range(len(self.data)):
                ras = int(self.data.iloc[i]['TOTAL_RASHOD'])
                tab.append({'Точка А-Б':self.data.index[i] + 1, 'Расходы':ras})
                
            df_tab = pd.DataFrame(tab)
            f, ax = plt.subplots()
            plt.tick_params(labelsize=8)
            ax=sns.barplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Расходы'],data=df_tab,color='red')
            ax.set_xlabel('')
            ax.set_ylabel('')
            #ax.legend(handles=ax.lines[::len(df_tab)+1], labels=["Расходы"], ncol=1, loc=9, fontsize=8)
            figure = plt.figure()
            canvas = FigureCanvas(f)
            self.fd.lay = QVBoxLayout()
            self.fd.lay.addWidget(canvas)
            self.fd.setLayout(self.fd.lay)
            self.fd.show()
        else:
            QMessageBox.information(self, 'Ошибка', 'Данные пусты или некорректный анализируемый период')
            
    def FifthReport(self):
        self.data = pd.read_csv(self.dataPath())
        self.fd = FigDialog(self)
        self.fd.setGeometry(1070, 530, 800, 450)
        tab = []
        isAvail = False
        if (self.line_from.text() == '')&(self.line_till.text() == ''):
            self.data = self.data
            isAvail = True
            
        if (self.line_from.text() != '')&(self.line_till.text() == ''):
            int_from = int(self.line_from.text()) - 1
            if (int_from <= len(self.data))&(int_from >= 0):
                self.data = self.data[self.data.index >= int_from]
                isAvail = True
                
        if (self.line_from.text() == '')&(self.line_till.text() != ''):
            int_till = int(self.line_till.text()) - 1
            if (int_till <= len(self.data))&(int_till >= 0):
                self.data = self.data[self.data.index <= int_till]
                isAvail = True
                
        if (self.line_from.text() != '')&(self.line_till.text() != ''):
            int_from = int(self.line_from.text()) - 1
            int_till = int(self.line_till.text()) - 1
            if (int_from >= 0)&(int_from <= len(self.data))&(int_till <= len(self.data))&(int_till >= 0)&(int_from <= int_till):
                self.data = self.data[(self.data.index >= int_from)&(self.data.index <= int_till)]
                isAvail = True
        
        if (len(self.data) > 0)&(isAvail):
            for i in range(len(self.data)):
                ras = int(self.data.iloc[i]['SALDO'])
                tab.append({'Точка А-Б':self.data.index[i] + 1, 'Сальдо':ras})
                
            df_tab = pd.DataFrame(tab)
            f, ax = plt.subplots()
            plt.tick_params(labelsize=8)
            ax=sns.barplot(ax=ax,x=df_tab['Точка А-Б'],y=df_tab['Сальдо'],data=df_tab,color='green')
            ax.set_xlabel('')
            ax.set_ylabel('')
            #ax.legend(handles=ax.lines[::len(df_tab)+1], labels=["Расходы"], ncol=1, loc=9, fontsize=8)
            figure = plt.figure()
            canvas = FigureCanvas(f)
            self.fd.lay = QVBoxLayout()
            self.fd.lay.addWidget(canvas)
            self.fd.setLayout(self.fd.lay)
            self.fd.show()
        else:
            QMessageBox.information(self, 'Ошибка', 'Данные пусты или некорректный анализируемый период')
            
class NewDialog(QWidget):
    
    def __init__(self, parent):
        super(NewDialog, self).__init__(parent)
        
class FigDialog(QWidget):
        
    def __init__(self, parent):
        super(FigDialog, self).__init__(parent)
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())