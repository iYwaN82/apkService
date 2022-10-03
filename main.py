from datetime import datetime
import os
import shutil
import time
import fdb
import schedule
import configparser
import pandas as pd
import win32com.client as win32
import xlwt

#commit

global ini
config_read_timer = 10
dir = os.path.abspath(os.curdir)


def del_tmp():
    try:
        os.remove(dir + "\\tmp.fdb")
        log_print("[+] Удаляем временную БД")
    except:
        log_print("[!] Не могу удалить временную базу, файл отсутсвует")


def log_print(txt: str):
    dt = datetime.now()
    log_time = f"[{str(dt.now())[0:19]}]"
    log = open(dir + "\\log.txt", "a+")
    log.write(log_time + ' ' + txt + "\n")
    print("LOG:" + txt)
    log.close()


def read_ini():
    global ini
    ini = configparser.ConfigParser()
    ini.read(dir + "\\settings.ini")


def change_ini():
    try:
        old_time = ini["main"]["time"]
    except:
        old_time = ""

    read_ini()
    cur_time = ini['main']['time']
    # print(f"\nПуть к базе: {ini['main']['base']}\nПуть выгрузки: {ini['main']['out_path']}\nВремя выгрузки: {ini['main']['time']}")

    if old_time != cur_time:
        if old_time != "":
            log_print(f"[+] Время запуска изменено с {old_time} на {ini['main']['time']}")
        schedule.clear('main')
        schedule.every().day.at(str(ini['main']['time'])).do(write_srv).tag('main')

    log_print(f"{schedule.get_jobs('main')}")


def start_srv():
    todo: "Запуск сервера"


def write_srv():
    # --------- Копируем базу с корневой каталог
    # global dbpath
    global con, dbpath, df
    if os.path.exists(dir + "\\tmp.fdb"):
        del_tmp()
    try:
        shutil.copyfile(str(ini['main']['base']), dir + "\\tmp.fdb")
        dbpath = dir + "\\tmp.fdb"
        log_print("[+] База скопирована в корневую директорию")
    except:
        log_print("[!] Не удалось скопировать локальную базу.")

    # ------- Открываем БД
    dbport = 3356
    dbhost = '127.0.0.1'

    try:
        con = fdb.connect(host=dbhost,
                          port=dbport,
                          database=dbpath,
                          user='sysdba',
                          password='masterkey',
                          charset='UTF8',
                          fb_library_name=str(dir + '\\fbclient.dll'))
        log_print("[+] Открываем БД")
    except:
        log_print("[!] Не могу открыть файл БД")
    # Create a Cursor object that operates in the context of Connection con:
    cur = con.cursor()

    sql_req = """SELECT 
	CAST(DS_BATCH_DELIVERY.DELIVERED_TIME AS DATE) AS SDATE,
	DS_GROUP_TYPE.NAME AS GT, 
	DS_GROUP.DISPLAY_NAME AS NAME, 
	DS_GROUP.DESCRIPTION AS DIS,
	ROUND (SUM(DS_BATCH_DELIVERY.DELIVERED_WEIGHT),0) as DW,
	ROUND (SUM(DS_BATCH_DELIVERY.DELIVERED_WEIGHT * DS_BATCH_DELIVERY.ACTUALDM_PERC/100),0) AS DWS,
	ROUND (SUM(DS_BATCH_DELIVERY.WEIGHBACK_AMOUNT), 0) as DWWB
	FROM DS_BATCH
	INNER JOIN DS_BATCH_DELIVERY ON	DS_BATCH_DELIVERY.BATCH_ID = DS_BATCH.ID 
	INNER JOIN DS_GROUP ON	DS_GROUP.ID = DS_BATCH_DELIVERY.GROUP_ID
	INNER JOIN DS_GROUP_TYPE ON DS_GROUP_TYPE.ID = DS_GROUP.GROUP_TYPE
	WHERE DS_BATCH_DELIVERY.DELIVERED_TIME >= CURRENT_DATE - 7 
	GROUP BY SDATE,GT, NAME, DIS"""

    # Execute the SELECT statement:
    cur.execute(sql_req)

    try:
        df = pd.DataFrame(cur.fetchall())
        print("Формируем DF")
    except:
        log_print("[!] Не понятная хйня в отчете")

    df.rename(
        columns={0: 'Дата', 1: 'Тип группы', 2: 'Имя на дисплее', 3: 'Описание', 4: 'Фактический вес',
                 5: 'Фактический сухой вес', 6: 'Вес остатков корма'},
        inplace=True)

    print(df)

    dt = datetime.now()
    xlsx_date = str(dt.date()).replace("-", "")
    xlsx_fname = xlsx_date + '_1000 ПРК - APK' + '.xlsx'
    log_print("[i] Файл для выгрузки в формате XLSX: " + xlsx_fname)
    log_print("[i] Страница в XLSX: " + xlsx_fname[:-5])
    df_sheets = {xlsx_fname[:-5]: df}
    writer = pd.ExcelWriter(dir + "\\data\\" + xlsx_fname, engine='xlsxwriter')
    log_print("[i] Сохраняем файл: " + dir + "\\data\\" + xlsx_fname)

    for sheet_name in df_sheets.keys():
        df_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    try:
        writer.save()
        #writer.close()
        log_print("[i] Файл сохранен: " + dir + "\\data\\" + xlsx_fname)
    except PermissionError:
        log_print("[!] Файл таблицы XLSX занят")


    log_print ("[+] Конвертируем XLSX-XLS")
    try:
        filename = dir + '\\data\\' + xlsx_fname
        pd.read_excel(filename).to_excel(filename[:-1])
        shutil.copyfile(str(dir + "\\data\\" + xlsx_fname)[:-1], ini['main']['out_path'] + xlsx_fname[:-1])
        log_print ("[+] Копируем " + str(dir + '\\data\\' + xlsx_fname)[:-1] + " + в " + ini['main']['out_path'] + xlsx_fname[:-1])
    except:
        log_print(
            "[-] Ошибка конвертации и копирования: " + str(dir + '\\data\\' + xlsx_fname)[:-1] + " + в " + ini['main']['out_path'] + xlsx_fname[
                                                                                                            :-1])

    """
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    log_print("[-1-]")
    excel.DisplayAlerts = False
    log_print("[-2-]")
    log_print("[+] Открываем XLSX:" + dir + "\\data\\" + xlsx_fname)
    wb = excel.Workbooks.Open(dir + "\\data\\" + xlsx_fname)
    log_print("[-3-]")
    try:
        log_print("[-4-]")
        log_print("[+] Сохраняем в XLS файл формата Excel 97/2003")
        wb.SaveAs(str(dir + "\\data\\" + xlsx_fname)[:-1], 56)
        #wb.Close()
        excel.Application.Quit()
        shutil.copyfile(str(dir + "\\data\\" + xlsx_fname)[:-1], ini['main']['out_path'] + xlsx_fname[:-1])
        log_print(
            "[+] Копируем " + str(dir + '\\data\\' + xlsx_fname)[:-1] + " + в " + ini['main']['out_path'] + xlsx_fname[
                                                                                                            :-1])
    except:
        log_print("[!] Нет доступа к XLS файлу, либо MS Excel не установлен")
    """
    # cur.close()

    # Удаляем временную базу
    del_tmp()


def main():
    change_ini()
    schedule.every(config_read_timer).seconds.do(change_ini)
    # time.sleep(config_read_timer + 2)
    # write_srv()
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()
