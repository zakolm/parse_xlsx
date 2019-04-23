#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fileencoding=utf-8

try:
    import configparser
except:
    from six.moves import configparser
import logging
import logging.config
import os
import pyexcel
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def sent_email(NAME, TO, SUBJECT):
    """
    Send result valid RFS.
    """
    try:
        config = configparser.ConfigParser()
        config.read('config.ini')

        HOST = config.get('DEFAULT', 'host')
        FROM = config.get('DEFAULT', 'email')
        PASSWORD = config.get('DEFAULT', 'password')

        msg = MIMEMultipart('alternative')
        msg['Subject'] = SUBJECT
        msg['From'] = FROM
        msg['To'] = TO

        html = '<html><body><p>Добрый день, ' + NAME + '!<br>' + \
               '<br>Прошу ознакомиться с новым предложением.' + \
               '<br>Прайс во вложении.<br><br>С уважением, команда РЫБАКОВ!</p></body></html>'
        part2 = MIMEText(html, 'html')

        msg.attach(part2)

        server = smtplib.SMTP_SSL(HOST)
        server.login(FROM, PASSWORD)

        server.sendmail(FROM, TO, msg.as_string())
        server.quit()
    except UnicodeDecodeError as e:
        logging.error('Ошибка кодировки: {}'.format(e))


def main():
    """
    Тут должен быть только один файл(все остальные убраны скриптом start),
        но если что, то будет взят последний отсортированный(по системе).
    """

    try:
        filename = ''
        path = '../'
        for file in os.listdir(path):
            if file.endswith('xlsx'):
                filename = file
                break

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
                if book_dict[keys[0]][i][4] == '':
                    logging.info('Отправляю письмо')
                    sent_email(book_dict[keys[0]][i][0], book_dict[keys[0]][i][1], book_dict[keys[0]][i][2])
                    logging.info('Отправил письмо')
        else:
            print('Нет файла!')
    except:
        logging.error('error')


if __name__ == '__main__':
    logging.config.fileConfig('logging.conf')
    logging.info('Начало программы')
    try:
        main()
    finally:
        logging.info('Конец программы')
