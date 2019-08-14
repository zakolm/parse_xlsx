#!/usr/bin/env python
# -*- coding: utf-8 -*-

try:
    import configparser
except:
    from six.moves import configparser
import errno
import logging
import logging.config
import os
import smtplib
import ssl
import stat
import pyexcel
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def sent_email(NAME, TO, SUBJECT):
    """
    Send result valid RFS.
    """
    try:
        HOST = config.get('DEFAULT', 'host')
        FROM = config.get('DEFAULT', 'email')
        PASSWORD = config.get('DEFAULT', 'password')

        msg = MIMEMultipart('alternative')
        msg['Subject'] = SUBJECT
        msg['From'] = FROM
        msg['To'] = TO

        message = """Добрый день, %s!
Прошу ознакомиться с новым предложением.
Прайс во вложении.\n\n
С уважением, команда РЫБАКОВ!""" % NAME.encode('utf-8')

        server = smtplib.SMTP_SSL(HOST)
        server.login(FROM, PASSWORD)
        server.sendmail(FROM, TO, message)
    except UnicodeDecodeError as e:
        logging.error('Ошибка кодировки: {}'.format(e))


def main():
    """
    Тут должен быть только один файл(все остальные убраны скриптом start),
        но если что, то будет взят последний отсортированный(по системе).
    """

    try:
        filename = ''
        path = './'
        for file in os.listdir(path):
            if file.endswith('xlsx'):
                filename = file
                break
        logging.debug(filename)

        # Проверка найден ли хоть один файл в директории.
        if filename != '':
            filename = path + filename
            # Get your data in an ordered dictionary of lists
            my_dict = pyexcel.get_dict(file_name=filename, name_columns_by_row=0)

            # Get your data in a dictionary of 2D arrays
            book_dict = pyexcel.get_book_dict(file_name=filename)

            # Get your keys from dictionary
            keys = list(book_dict.keys())

            logging.info('Обработал xlsx-файл')

            for i in range(2, len(book_dict[keys[0]])):
                if book_dict[keys[0]][i][4] == '' and book_dict[keys[0]][i][0] != '':
                    logging.info('Отправляю письмо %s' % (str(book_dict[keys[0]][i][1])))
                    sent_email(book_dict[keys[0]][i][0], book_dict[keys[0]][i][1], book_dict[keys[0]][i][2])
                    logging.info('Отправил письмо %s' % (str(book_dict[keys[0]][i][1])))
        else:
            print('Нет файла!')
    except:
        logging.error('Unknown error')


def settings_log(log_config, log_dir):
    """
    Настройка логгирования
    :param log_config: файл с лог-конфигом
    :param log_dir: папка с логами
    :return: None
    """
    try:
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        logging.config.fileConfig(log_config)
    except OSError as error:
        if error.errno != errno.EEXIST:
            raise


if __name__ == '__main__':
    # Расположение скрипта.
    script_home = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(script_home)

    print(script_home)

    try:
        # Настройки логгирования
        settings_log(log_config=script_home + '/scripts/logging.conf',
                     log_dir=script_home + '/logs/')
        # Работа с конфигом
        os.chmod(script_home + '/scripts/config.ini', stat.S_IREAD | stat.S_IWRITE)
        config = configparser.ConfigParser()
        config.read(script_home + '/scripts/config.ini')
        # Старт скрипта
        logging.info('Начало программы')
        main()
    finally:
        logging.info('Конец программы')

