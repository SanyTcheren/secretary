#! /usr/bin/env python3
"""Helper for create report and send email to boss."""

from pathlib import Path
from datetime import datetime, timedelta
import shelve
import logging
import openpyxl
import ezgmail

class Secretary:
    """My beautiful girl."""

    oil_path = './blank/oil.xlsx' # файл с образцом отчета по ДЭС
    work_path = './blank/work.xlsx' # файл с образцом отчета оп ППР
    storage_path = './blank/storage.xlsx' # файл с перечнем ППР
    init_path = './init.data' # файл с исходными данными

    def __init__(self):
        """Загружает файл с настройками, при отсутствии файла - создает его."""
        logging.info('Create secretary.')
        if not Path(Secretary.init_path).is_file():
            Secretary.init_data()
        init_file = shelve.open(Secretary.init_path)
        self.boss_email = init_file['boss_email']
        self.my_email = init_file['my_email']
        self.brigade = init_file['brigade']
        init_file.close()
        Secretary.check_gmail(self.my_email)

    def work(self):
        """Выполнение основных обязаностей."""
        today = datetime.now().date() # День отчета
        yesterday = (datetime.now() - timedelta(days=1)).date() # День работ

        storage_wb = openpyxl.load_workbook(Secretary.storage_path)
        storage_sheet = storage_wb.active
        job = storage_sheet.cell(row=yesterday.day, column=3).value
        device = storage_sheet.cell(row=yesterday.day, column=2).value
        place = storage_sheet.cell(row=yesterday.day, column=1).value

        work_wb = openpyxl.load_workbook(Secretary.work_path)
        work_sheet = work_wb.active
        work_sheet['A3'] = today
        work_sheet['D4'] = today
        work_sheet['B4'] = yesterday
        work_sheet['A7'] = f"{work_sheet['A7'].value} {place}, {device}"
        work_sheet['C6'] = job
        work_report = f'reports/{self.brigade}бр_Суточный_отчет_ППР_{today}.xlsx'
        work_wb.save(work_report)

        oil_wb = openpyxl.load_workbook(Secretary.oil_path)
        oil_sheet = oil_wb.active
        oil_sheet['C12'] = today
        oil_sheet['C19'] = yesterday
        oil_report = f'reports/{self.brigade}бр_Суточный_отчет_ДЭС_{today}.xlsx'
        oil_wb.save(oil_report)

        ezgmail.send(self.boss_email,
                     f"{self.brigade}бр Суточный отчет",
                     '',
                     [work_report, oil_report])
        logging.info('Work completed.')

    @staticmethod
    def init_data():
        """Создание файла с начальными данными."""
        logging.info('Запись начальных данных')
        init_file = shelve.open(Secretary.init_path)
        init_file['boss_email'] = input('boss email:')
        init_file['my_email'] = input('my email:')
        init_file['brigade'] = input('brigade:')
        init_file.close()

    @staticmethod
    def check_gmail(gmail):
        """Проверка связи с gmail."""
        try:
            ezgmail.init()
            if gmail != ezgmail.EMAIL_ADDRESS:
                logging.error('Несовпадает Ваш email c credentials.json')
                raise Exception
            logging.info('Check_gmail successfully completed.')
        except Exception as ex:
            logging.error('Неудалось связаться с gmail, проверьте токен в credentials.json')
            raise ex


def main():
    """Work."""
    day = datetime.now().date()
    logging.basicConfig(filename=f"./logs/{day}",
                        level=logging.INFO,
                        format="%(asctime)s%(levelname)s[%(name)s] %(message)s")
    secretary = Secretary()
    secretary.work()


if __name__ == '__main__':
    main()
