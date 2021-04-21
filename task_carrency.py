import requests
from datetime import datetime
from lxml import etree
import xlsxwriter

from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header
import os

import pymorphy2

URL = 'https://www.moex.com/export/derivatives/currency-rate.aspx?language=ru&'
USD = 'USD/RUB'
EUR = 'EUR/RUB'
TOP_TABLE = {'A1': 'Дата USD',
             'B1': 'Курс USD',
             'C1': 'Изменение USD',
             'D1': 'Дата EUR',
             'E1': 'Курс EUR',
             'F1': 'Изменение EUR',
             'G1': 'EUR/USD'}

now_date = datetime.today()
end_date = now_date.strftime("%Y-%m-%d")
start_date = f'{now_date.strftime("%Y")}-{now_date.strftime("%m")}-1'


def create_xlsx(data_dollar, data_euro):
    """Create table and filling in the data"""

    name_work_book = f'{end_date}-formula.xlsx'
    workbook = xlsxwriter.Workbook(name_work_book)
    worksheet = workbook.add_worksheet()

    bold_format = workbook.add_format({'bold': True})
    bold_format.set_align('center')
    money_format = workbook.add_format({'num_format': '#,##0.000"₽"'})
    money_format.set_align('center')

    for k, v in TOP_TABLE.items():
        worksheet.write(f'{k}', f'{v}', bold_format)
    for i, data in enumerate(zip(data_dollar, data_euro), start=2):
        data_u, data_e = data
        worksheet.write(f'A{i}', data_u.attrib['moment'])
        worksheet.write(f'B{i}', float(data_u.attrib['value']), money_format)
        worksheet.write(f'D{i}', data_e.attrib['moment'])
        worksheet.write(f'E{i}', float(data_e.attrib['value']), money_format)
        worksheet.write(f'G{i}', f'=B{i}/E{i}', money_format)
        if i - 1 < len(data_dollar):
            worksheet.write(f'C{i}', f'=B{i}-B{i + 1}', money_format)
            worksheet.write(f'F{i}', f'=E{i}-E{i + 1}', money_format)

    worksheet.set_column('A:G', 20)
    workbook.close()

    return name_work_book


def get_xml(currency, start_date, end_date):
    """Get xml from https://www.moex.com"""

    answer = requests.get(f'{URL}&currency={currency}&moment_start={start_date}&moment_end={end_date}')
    tree = etree.fromstring(bytes(answer.text, encoding='windows-1251'))
    return tree.xpath('//rate')


def get_mail():
    """Get mail and password"""

    address = input('Введите mail: ')
    password_mail = input('Введите пароль: ')
    server = input('Введите адрес сервера: ')
    return address, password_mail, server


def send_mail(filepath, address, password_mail, server, body):
    """Send email with an attached file"""

    basename = os.path.basename(filepath)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filepath, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % basename)

    msg = MIMEMultipart()
    msg['Subject'] = Header('таблица', 'utf-8')
    msg['From'] = address
    msg['To'] = address
    msg.attach(MIMEText(body, 'plain'))
    msg.attach(part)

    server = SMTP_SSL(server)
    server.ehlo(address)
    server.login(address, password_mail)
    server.auth_plain()
    server.sendmail(address, address, msg.as_string())
    server.quit()


def main():
    data_usd = get_xml(USD, start_date, end_date)
    data_eur = get_xml(EUR, start_date, end_date)
    count_row = len(data_usd)
    table = create_xlsx(data_usd, data_eur)
    morph = pymorphy2.MorphAnalyzer()
    row = morph.parse('строка')[0]
    body_mail = f'В таблице {count_row} {row.make_agree_with_number(count_row).word}'
    mail, password, server = get_mail()

    send_mail(table, mail, password, server, body_mail)


if __name__ == '__main__':
    main()
