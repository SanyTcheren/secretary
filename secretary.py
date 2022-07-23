#! /usr/bin/env python3
"""Helper for create report and send email to boss."""

from pathlib import Path
import shelve
import logging
import ezgmail

class Secretary:
    """My beautiful girl."""

    init_path = './init.data' # файл с исходными данными

    def __init__(self):
        """Загружает файл с настройками, при отсутствии файла - создает его."""
        logging.info('Создаем секретаршу.')
        if not Path(Secretary.init_path).is_file():
            Secretary.init_data()
        init_file = shelve.open(Secretary.init_path)
        self.boss_email = init_file['boss_email']
        self.my_email = init_file['my_email']
        init_file.close()
        Secretary.check_gmail(self.my_email)

    def work(self):
        """Выполнение основных обязаностей."""
        ezgmail.send(self.boss_email, 'Test secretary', '2 testing successfull.')
        logging.info('work completed.')

    @staticmethod
    def init_data():
        """Создание файла с начальными данными."""
        logging.info('Запись начальных данных')
        init_file = shelve.open(Secretary.init_path)
        init_file['boss_email'] = input('boss email:')
        init_file['my_email'] = input('my email:')
        init_file.close()

    @staticmethod
    def check_gmail(gmail):
        """Проверка связи с gmail."""
        try:
            ezgmail.init()
            if gmail != ezgmail.EMAIL_ADDRESS:
                logging.error('Несовпадает Ваш email c credentials.json')
                raise Exception
            logging.info('check_gmail successful completed.')
        except Exception as ex:
            logging.error('Неудалось связаться с gmail, проверьте токен в credentials.json')
            raise ex


def main():
    """Work."""
    logging.basicConfig(level=logging.INFO)
    secretary = Secretary()
    secretary.work()


if __name__ == '__main__':
    main()
