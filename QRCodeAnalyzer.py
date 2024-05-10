# coding=utf-8
import datetime
import os
import sys

import xlsxwriter  # Это модуль Python для записи файлов в формате Excel 2007+ XLSX

import PyQt5
from PyQt5 import QtGui, uic, QtCore
from PyQt5.QtCore import QDate, QTime
from PyQt5.QtWidgets import QApplication, QDialog, QTableWidgetItem, QMessageBox

import data_processing_engine


class AlignDelegate(PyQt5.QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = QtCore.Qt.AlignCenter


class STARTER(QDialog):
    def __init__(self):
        super(STARTER, self).__init__()
        try:
            uic.loadUi('qrcodeanalyzer.ui', self)
            self.setWindowIcon(QtGui.QIcon('logo.png'))
            self.qRadioButtonLong.toggled.connect(self.qRadioButtonLongToggled)
            self.qRadioButtonShort.toggled.connect(self.qRadioButtonShortToggled)
            self.qRadioButtonShort_2.toggled.connect(self.qRadioButtonShortToggled)
            self.qDateTimeEditBegin.dateChanged.connect(self.qDateTimeEditBeginDateChanged)
            self.qDateTimeEditEnd.dateChanged.connect(self.qDateTimeEditEndDateChanged)
            self.qDateTimeEditBegin.timeChanged.connect(self.qDateTimeEditBeginTimeChanged)
            self.qDateTimeEditEnd.timeChanged.connect(self.qDateTimeEditEndTimeChanged)

            self.qTableWidgetLong.setColumnWidth(0, 80)  # Дата
            self.qTableWidgetLong.setColumnWidth(1, 80)  # Время
            self.qTableWidgetLong.setColumnWidth(2, 660)  # ФИО
            delegate = AlignDelegate(self.qTableWidgetLong)
            self.qTableWidgetLong.setItemDelegateForColumn(0, delegate)
            self.qTableWidgetLong.setItemDelegateForColumn(1, delegate)
            # self.qTableWidgetLong.setItemDelegateForColumn(2, delegate)

            self.qTableWidgetShort.setColumnWidth(0, 660)  # ФИО
            self.qTableWidgetShort.setColumnWidth(1, 160)  # Количество посещений

            self.qComboBoxLesson.currentIndexChanged.connect(self.qComboBoxLessonCurrentIndexChanged)
            self.qComboBoxGroup.currentIndexChanged.connect(self.qComboBoxGroupCurrentIndexChanged)

            self.qPushButtonExport.clicked.connect(self.qPushButtonExportClicked)
            self.qLineEditFamFilter.textChanged.connect(self.qLineEditFamFilterTextChanged)

            self.qTableWidgetShort.cellDoubleClicked.connect(self.qTableWidgetShortCellDoubleClicked)
            self.qTableWidgetLong.cellDoubleClicked.connect(self.qTableWidgetLongCellDoubleClicked)

            self.qPushButtonFamFilterClear.clicked.connect(self.qPushButtonFamFilterClearClicked)

            self.qPushButtonExit.clicked.connect(self.close)

            # print(datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S'))

            dtts = data_processing_engine.getMinMaxDate()
            d1 = dtts[0][0]
            d2 = dtts[0][1]
            t1 = dtts[0][2]
            t2 = dtts[0][3]

            minYear = int(d1.split("-")[0])
            minMonth = int(d1.split("-")[1])
            minDay = int(d1.split("-")[2])

            maxYear = int(d2.split("-")[0])
            maxMonth = int(d2.split("-")[1])
            maxDay = int(d2.split("-")[2])

            minHour = int(t1.split(":")[0])
            minMin = int(t1.split(":")[1])
            minSec = int(t1.split(":")[2])

            maxHour = int(t2.split(":")[0])
            maxMin = int(t2.split(":")[1])
            maxSec = int(t2.split(":")[2])

            self.qDateTimeEditBegin.setDate(QDate(minYear, minMonth, minDay))  # год, месяц, число
            self.qDateTimeEditEnd.setDate(QDate(maxYear, maxMonth, maxDay))  # год, месяц, число
            self.qDateTimeEditBegin.setTime(QTime(minHour, minMin, minSec))  # часы, минуты, секунды
            self.qDateTimeEditEnd.setTime(QTime(maxHour, maxMin, maxSec))  # часы, минуты, секунды

            diss = data_processing_engine.diss
            self.qComboBoxLesson.clear()
            for d in diss:
                self.qComboBoxLesson.addItems(d)
        except:
            pass

    def qComboBoxLessonCurrentIndexChanged(self):
        cdis = str(self.qComboBoxLesson.currentText())
        gr = data_processing_engine.getGroupsForDisc([cdis])
        self.qComboBoxGroup.clear()
        for g in gr:
            self.qComboBoxGroup.addItems(g)

    def qComboBoxGroupCurrentIndexChanged(self):
        dis = str(self.qComboBoxLesson.currentText())
        gr = str(self.qComboBoxGroup.currentText())
        d1 = str(self.qDateTimeEditBegin.date().toString("yyyy-MM-dd"))
        d2 = str(self.qDateTimeEditEnd.date().toString("yyyy-MM-dd"))
        t1 = str(self.qDateTimeEditBegin.time().toString("hh:mm:ss"))
        t2 = str(self.qDateTimeEditEnd.time().toString("hh:mm:ss"))

        while (self.qTableWidgetLong.rowCount() > 0):
            self.qTableWidgetLong.removeRow(0)

        while (self.qTableWidgetShort.rowCount() > 0):
            self.qTableWidgetShort.removeRow(0)

        famFilter = self.qLineEditFamFilter.text().strip(" ").upper()

        att = data_processing_engine.getAttendance(dis, gr, d1, d2, t1, t2, famFilter)
        for i, a in enumerate(att):
            self.qTableWidgetLong.insertRow(self.qTableWidgetLong.rowCount())
            self.qTableWidgetLong.setItem(i, 0, QTableWidgetItem(a[0]))
            self.qTableWidgetLong.setItem(i, 1, QTableWidgetItem(str(a[1])))
            self.qTableWidgetLong.setItem(i, 2, QTableWidgetItem(str(a[2])))

        att = data_processing_engine.getNumAttendance(dis, gr, d1, d2, t1, t2,
                                                      self.qRadioButtonShort_2.isChecked(), famFilter)

        for i, a in enumerate(att):
            self.qTableWidgetShort.insertRow(self.qTableWidgetShort.rowCount())
            self.qTableWidgetShort.setItem(i, 0, QTableWidgetItem(a[0]))
            self.qTableWidgetShort.setItem(i, 1, QTableWidgetItem(str(a[1])))

    def qRadioButtonLongToggled(self):
        if (self.qRadioButtonLong.isChecked()):
            self.qTableWidgetLong.show()
            self.qTableWidgetShort.hide()

    def qRadioButtonShortToggled(self):
        self.qComboBoxGroupCurrentIndexChanged()
        if (self.qRadioButtonShort.isChecked() or self.qRadioButtonShort_2.isChecked()):
            self.qTableWidgetShort.show()
            self.qTableWidgetLong.hide()

    def qDateTimeEditBeginDateChanged(self):
        self.qDateTimeEditBegin
        self.qComboBoxGroupCurrentIndexChanged()

    def qDateTimeEditEndDateChanged(self):
        self.qComboBoxGroupCurrentIndexChanged()

    def qDateTimeEditBeginTimeChanged(self):
        self.qComboBoxGroupCurrentIndexChanged()

    def qDateTimeEditEndTimeChanged(self):
        self.qComboBoxGroupCurrentIndexChanged()

    def qPushButtonExportClicked(self):
        self.exportToExcel()

    def qLineEditFamFilterTextChanged(self):
        self.qComboBoxGroupCurrentIndexChanged()

    def qTableWidgetShortCellDoubleClicked(self, row, col):
        text = self.qTableWidgetShort.item(row, 0).text()
        n1 = text.find("(")
        if n1 != -1:
            fio = text[:n1].strip(" ").upper()
            self.qLineEditFamFilter.setText(fio)

    def qTableWidgetLongCellDoubleClicked(self, row, col):
        text = self.qTableWidgetLong.item(row, 2).text()
        n1 = text.find("(")
        if n1 != -1:
            fio = text[:n1].strip(" ").upper()
            self.qLineEditFamFilter.setText(fio)

    def qPushButtonFamFilterClearClicked(self):
        self.qLineEditFamFilter.setText("")

    def exportToExcel(self):
        dis = str(self.qComboBoxLesson.currentText())
        gr = str(self.qComboBoxGroup.currentText())
        d1 = str(self.qDateTimeEditBegin.date().toString("yyyy-MM-dd"))
        d2 = str(self.qDateTimeEditEnd.date().toString("yyyy-MM-dd"))
        t1 = str(self.qDateTimeEditBegin.time().toString("hh:mm:ss"))
        t2 = str(self.qDateTimeEditEnd.time().toString("hh:mm:ss"))

        s = "???"
        if (self.qRadioButtonShort.isChecked()):
            s = self.qRadioButtonShort.text()
        if (self.qRadioButtonShort_2.isChecked()):
            s = self.qRadioButtonShort_2.text()
        if (self.qRadioButtonLong.isChecked()):
            s = self.qRadioButtonLong.text()

        famFilter = self.qLineEditFamFilter.text().strip(" ").upper()

        if (self.qRadioButtonShort.isChecked() or self.qRadioButtonShort_2.isChecked()):
            att = data_processing_engine.getNumAttendance(dis, gr, d1, d2, t1, t2,
                                                          self.qRadioButtonShort_2.isChecked(), famFilter)
        else:
            att = data_processing_engine.getAttendance(dis, gr, d1, d2, t1, t2, famFilter)

        try:
            my_file = 'report.xlsx'

            book = xlsxwriter.Workbook(my_file)
            sheet = book.add_worksheet()
            bold = book.add_format({'bold': True})
            boldRight = book.add_format({'bold': True, 'align': 'right'})
            boldLeft = book.add_format({'bold': True, 'align': 'left'})
            boldCenter = book.add_format({'bold': True, 'align': 'center'})
            center = book.add_format({'align': 'center'})
            sheet.set_column('A:A', 5, center)

            sheet.merge_range('B1:C1', "ОТЧЁТ ПО ПОСЕЩАЕМОСТИ", boldCenter)
            sheet.write('B2', s, boldCenter)
            sheet.write('C2', datetime.datetime.now().strftime('от %d.%m.%Y %H:%M:%S'), bold)
            sheet.write('B4', "по дисциплине: ", boldCenter)
            sheet.write('C4', dis, bold)
            sheet.write('B5', "группа: ", boldCenter)
            sheet.write('C5', gr, bold)

            sheet.write('A10', "№", boldCenter)

            if (self.qRadioButtonShort.isChecked() or self.qRadioButtonShort_2.isChecked()):
                sheet.write('B10', "ФИО", boldCenter)
                sheet.write('C10', "Количество посещений", boldLeft)
                sheet.set_column('B:B', 56)
                sheet.set_column('C:C', 14, center)
                sheet.set_column('D:D', 40)
                sheet.set_column('E:E', 25)
                sheet.write('C7', d1 + " " + t1, bold)
                sheet.write('C8', d2 + " " + t2, bold)

                if famFilter != "":
                    sheet.write('D5', "Фильтр по фамилии: " + famFilter, boldLeft)

                sheet.write('B7', "Дата и время начала анализа: ", boldCenter)
                sheet.write('B8', "Дата и время конца анализа: ", boldCenter)
            else:
                sheet.set_column('B:B', 20, center)
                sheet.set_column('C:C', 20, center)
                sheet.set_column('D:D', 55)
                sheet.set_column('E:E', 30)
                sheet.write('B10', "Дата", boldCenter)
                sheet.write('C10', "Время", boldCenter)

                if famFilter != "":
                    sheet.write('D5', "Фильтр по фамилии: " + famFilter, boldLeft)

                sheet.write('D10', "ФИО", boldCenter)
                sheet.write('D7', d1 + " " + t1, bold)
                sheet.write('D8', d2 + " " + t2, bold)
                sheet.merge_range('B7:C7', "Дата и время начала анализа: ", boldCenter)
                sheet.merge_range('B8:C8', "Дата и время конца анализа: ", boldCenter)

            sm = 10
            for i, a in enumerate(att):
                sheet.write(sm + i, 0, i + 1)
                for j, b in enumerate(a):
                    sheet.write(sm + i, 1 + j, b)

            book.close()

            sf = os.system("start " + my_file)
            if sf != 0:
                raise Exception()
        except:
            QMessageBox.about(self, "Ошибка", "Не могу создать отчёт !\nВозможно файл уже открыт")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = STARTER()
    window.show()
    sys.exit(app.exec_())
