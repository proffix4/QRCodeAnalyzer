#!/usr/bin/env python3

import csv
import locale
import os
import sqlite3

from autocorrection_records import autocorrection_visit_record, autocorrection_group_record

diss = []  # Глобальный список дисциплин
pers = []  # Глобальный список посещений

current_dir = os.path.abspath(os.curdir)  # Текущая директория
csv_directory = os.path.join(current_dir, "Data")  # Директория с файлами csv


def parseData():
    """Парсинг данных из файлов csv"""
    files = os.listdir(csv_directory)  # Получаем список файлов в директории
    csv_files = filter(lambda x: x.endswith(".csv"), files)  # Получаем список файлов с расширением .csv
    os.chdir(csv_directory)  # Переходим в директорию с файлами csv
    for csv_file in csv_files:  # Перебор файлов csv
        with (((((open(csv_file, encoding="utf-8")))))):  # Открываем файл csv
            csv_reader = csv.reader(
                open(csv_file, encoding="utf-8", newline=""), delimiter=","
            )  # Создаем объект csv_reader для чтения файла csv
            dis = ""  # Теущая дисциплина при парсинге
            for line in csv_reader:  # Перебор строк файла csv
                date = line[0]  # Дата посещения
                time = line[1]  # Время посещения
                text = line[4]  # Текст посещения
                text = autocorrection_visit_record(text)  # Автокоррекция записи посещения
                ns = text.find("*")  # Поиск символа *
                if ns == -1:  # Если не встретился символ *, то это запись посещения
                    n1 = text.find("(")  # Поиск символа (
                    if n1 != -1:  # Если не встретился символ (
                        n2 = text.rfind(")")  # Поиск символа )
                        fio = text[:n1].strip(" ").upper()  # ФИО студента
                        group = text[n1 + 1: n2].strip(" ")  # Группа студента
                        group = autocorrection_group_record(group)  # Автокоррекция записи группы
                        pers.append([date, time, fio, group, dis])  # Добавляем запись в список посещений
                    else:
                        if text != "text":
                            print(
                                "Ошибка в файле: " + csv_file + " в строке: " + str(csv_reader.line_num) + " - " + text)
                else:  # Если встретился символ *, то это дисциплина
                    dis = text[ns + 2:].strip("*").strip(" ")  # Дисциплина
                    if [dis] not in diss:  # Если дисциплина еще не встречалась
                        diss.append([dis])  # Добавляем дисциплину в список дисциплин
    os.chdir(current_dir)  # Переходим в директорию с файлами csv


def createDB(pers):
    """Создание базы данных из списка посещений"""
    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    cur.execute("DROP TABLE IF EXISTS pers")  # Удаляем таблицу pers, если она существует
    cur.execute("CREATE TABLE pers(dt DATA, tm TIME, fio TEXT, gr TEXT, dis TEXT)")  # Создаем таблицу pers
    cur.executemany("INSERT INTO pers VALUES(?, ?, ?, ?, ?)", pers)  # Добавляем записи в таблицу pers из списка
    con.commit()  # Сохраняем изменения
    con.close()  # Закрываем базу данных


def getGroupsForDisc(dis):
    """Получение списка групп для указанной дисциплины"""
    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    sql = "SELECT gr FROM pers WHERE dis = '" + dis[0] + "' GROUP by gr"  # SQL запрос
    res = cur.execute(sql).fetchall()  # Выполняем запрос
    con.close()  # Закрываем базу данных
    return res  # Возвращаем список групп


def getNumAttendance(dis, gr, d1, d2, t1, t2, sortByName=True, famFilter=""):
    """Получение количества посещений студентов для указанной дисциплины, группы, дат и времени"""
    mma = getMaxNumAttendance(dis, gr, d1, d2, t1, t2)  # Максимальное количество посещений
    if mma == 0:  # Если максимальное количество посещений равно 0, то не считаем проценты
        sql1 = "SELECT fio || ' (' ||  gr || ')', count(dis) as c "
    else:  # Иначе считаем проценты
        sql1 = "SELECT fio || ' (' ||  gr || ')', count(dis)  ||  ' (' || (100 * count(dis) / " + str(
            mma) + ")  || '%)'  as c "

    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    sql = (
            sql1 +
            "FROM pers "
            "WHERE dis = '" + dis + "'"
                                    " AND gr = '" + gr + "'"
                                                         " AND dt BETWEEN '" + d1 + "' AND '" + d2 + "'"
                                                                                                     " AND tm BETWEEN '" + t1 + "' AND '" + t2 + "'"
                                                                                                                                                 " AND fio like '" + famFilter + "%'"
                                                                                                                                                                                 " GROUP by fio, gr, dis"
    )
    if not sortByName:
        sql = sql + " ORDER by c DESC, fio"
    else:
        sql = sql + " ORDER by fio"
    res = cur.execute(sql).fetchall()  # Выполняем запрос
    con.close()  # Закрываем базу данных
    return res  # Возвращаем количество посещений


def getAllNumAttendance(dis, gr, sortByName=True):
    """Получение количества посещений студентов для указанной дисциплины и группы за все время"""
    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    sql = (
            "SELECT fio || ' (' ||  gr || ')', count(dis) as c "
            "FROM pers "
            "WHERE dis = '" + dis + "'"
                                    " AND gr = '" + gr + "'"
                                                         " GROUP by fio, gr, dis"
    )
    if not sortByName:
        sql = sql + " ORDER by c DESC, fio"
    else:
        sql = sql + " ORDER by fio"
    res = cur.execute(sql).fetchall()  # Выполняем запрос
    con.close()  # Закрываем базу данных
    return res  # Возвращаем количество посещений


def getAttendance(dis, gr, d1, d2, t1, t2, famFilter):
    """Получение списка посещений для указанной дисциплины, группы, дат и времени"""
    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    sql = (
            "SELECT dt, tm, fio || ' (' ||  gr || ')'"
            "FROM pers "
            "WHERE dis = '" + dis + "'"
                                    " AND gr = '" + gr + "'"
                                                         " AND dt BETWEEN '" + d1 + "' AND '" + d2 + "'"
                                                                                                     " AND tm BETWEEN '" + t1 + "' AND '" + t2 + "'"
                                                                                                                                                 " AND fio like '" + famFilter + "%'"
                                                                                                                                                                                 " ORDER by dt, tm, fio"
    )
    res = cur.execute(sql).fetchall()  # Выполняем запрос
    con.close()  # Закрываем базу данных
    return res  # Возвращаем список посещений


def getMaxNumAttendance(dis, gr, d1, d2, t1, t2):
    """Получение максимального количества посещений студентом для указанной дисциплины, группы, дат и времени"""
    con = sqlite3.connect("data.db")  # Открываем базу данных
    cur = con.cursor()  # Создаем объект курсора
    sql = (
            "SELECT count(dis) as c "
            "FROM pers "
            "WHERE dis = '" + dis + "'"
                                    " AND gr = '" + gr + "'"
                                                         " AND dt BETWEEN '" + d1 + "' AND '" + d2 + "'"
                                                                                                     " AND tm BETWEEN '" + t1 + "' AND '" + t2 + "'"
                                                                                                                                                 " GROUP by fio, gr, dis"
    )
    sql = sql + " ORDER by c DESC, fio"
    res = cur.execute(sql).fetchall()  # Выполняем запрос
    con.close()  # Закрываем базу данных
    if len(res) > 0:  # Возвращаем количество посещений
        return res[0][0]
    else:
        return 0


def getMinMaxDate():
    con = sqlite3.connect("data.db")
    cur = con.cursor()
    sql = "SELECT MIN(dt), MAX(dt), MIN(tm), MAX(tm) FROM pers"
    res = cur.execute(sql).fetchall()
    con.close()
    return res


# Основная часть модуля, автозапуск при подключении модуля
locale.setlocale(locale.LC_ALL, "ru")  # Устанавливаем русскую локаль для правильных названий и форматов
parseData()  # Парсинг данных
createDB(pers)  # Создание БД
