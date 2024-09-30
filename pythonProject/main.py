# This is a sample Python script.
import asyncio
import smtplib
import email
import email.mime.application
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from gspread import Client, Spreadsheet

import pay

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1-2fMMxxau3QKULV9yo_POBMdrehzN2a-ipcGiOaaGVc/edit?resourcekey=&gid=462629804#gid=462629804"
sender_email = "dendrod5@mail.ru"
receiver_email = "dendrod5@mail.ru"
password = "TSpyuJ4ZX1T8CuJvZ1hp"
filename = 'Приглашение Кругосветка 2024.pdf'

msg = MIMEMultipart()
with open('lastID.txt', encoding='utf-8') as f:
    lastID = int(f.read())

def generate_message_text(n, ws, ws1):
    htmlBody = '<p>Поздравляем, Вы успешно зарегистрировались на Кругосветку-2024!<br> Пожалуйста, внимательно ознакомьтесь с информацией ниже.</p>  <h4><u>Регистрационные данные:</u></h4><ol>'
    # ---НОМЕР---
    htmlBody += '<li>Индивидуальный номер и номер карты участника: ' + str(n) + '</li>'
    # ---ДИСТАНЦИЯ---
    if ws.cell(row=n, col=13).value == None:
        htmlBody += '<li>Дистанция: ' + ws.cell(row=n, col=14).value + '</li>'
    else:
        htmlBody += '<li>Дистанция: ' + ws.cell(row=n, col=13).value + '</li>'
    # ---КОЛИЧЕСТВО ЧЕЛОВЕК ---
    if (ws.cell(row=n, col=5).value == 'Групповая'):
        htmlBody += '<li>Количество человек в группе: ' + str(ws.cell(row=n, col=8).value) + '</li>'
    # ---МЕСТО СТАРТА - --
    htmlBody += '<li>Место старта: ' + ws.cell(row=n, col=12).value + '</li>'

        #htmlBody += '<li>Время отправления теплохода: ' + str(values['Регистрация на теплоход (прямой)']) + \
      #              ' <strong>Не опаздывайте!</strong></li>'
    #elif (values['Регистрация на теплоход (обратный)'] != 0):
     #   htmlBody += '<li>Время отправления теплохода: ' + values[
      #  'Регистрация на теплоход (обратный)'] + ' <strong>Не опаздывайте!</strong></li>';

    # ---ЦЕНА - --
    htmlBody += '</ol>'
    allPrice = 100
    if (ws.cell(row=n, col=5).value == 'Групповая'):
        allPrice = 100 * (int(ws.cell(row=n, col=8).value) - int(ws.cell(row=n, col=9).value))
    if (allPrice < 0):
        allPrice = 0
    htmlBody += '<h4><u>Стоимость участия:</u></h4> <p>' + str(allPrice) + ' рублей. (100 рублей с человека, дети до 5 лет бесплатно). </p>'
    htmlBody += '<h4><u>Инструкция по оплате:</u></h4> <p>Оплатить Вы можете двумя способами: </p> <ol> <li> По QR-коду из прикреплённого приглашения </li> <li> В день проведения мероприятия (6 октября) на точке старта Вашего маршрута. </li></ol> </p> <p>Для подтверждения стартового взноса онлайн необходимо пройти онлайн форму по ссылке: https://forms.gle/VXaRP9V3qFJMGKBn8</p>'\

    # ---ТЕПЛОХОД---
    htmlBody += '<p><strong>Для регистрации на маршруты: </p> <ol> <li> Пешеходный маршрут 15 км (Парк им. С.М.Кирова)</li> <li> Пешеходный маршрут 20 км (Парк им. С.М.Кирова) </li> <li> Пешеходный маршрут 30 км (Парк им. С.М.Кирова)</li> </ol> <p> Необходимо оплатить стартовый взнос онлайн в течении суток после регистрации, иначе мы будем вынуждены её отменить</strong></p>'
    htmlBody += '<p>Оставшиеся места на теплоходы для данных маршрутов:</p> <ol> <li> 1 рейс - ' + ws1.cell(row=4, col=3).value + '</li>'
    htmlBody += '<li>2 рейс - ' + ws1.cell(row=5, col=3).value + '</li>'
    htmlBody += '<li>3 рейс - ' + ws1.cell(row=6, col=3).value + '</li>'
    htmlBody += '<li>4 рейс - ' + ws1.cell(row=7, col=3).value + '</li>'
    htmlBody += '<li>5 рейс - ' + ws1.cell(row=8, col=3).value + '</li>'
    htmlBody += '<li>6 рейс - ' + ws1.cell(row=9, col=3).value + '</li></ol>'
    htmlBody += '<p>После оплаты <strong>получить карту маршрута </strong>вы можете в день проведения мероприятия (6 октября) на точке старта Вашего маршрута. </p> <p>По всем вопросам звоните по телефону 8-912-746-12-45, 8-906-819-81-79 - Радионова Наталья Антоновна, президент Федерации спортивного туризма Удмуртии.</p> <p>Желаем Вам приятного путешествия!</p>'
    return htmlBody

async def main():
    # Use a breakpoint in the code line below to debug your script.
    # Press Ctrl+F8 to toggle the breakpoint.
    gs: Client = gspread.service_account("./Serv-Acc.json")
    sh: Spreadsheet = gs.open_by_url(SPREADSHEET_URL)
    ws = sh.sheet1
    ws1 = sh.worksheet(title='Места на теплоход')
    n = lastID
    fp = open(filename, 'rb')
    att = email.mime.application.MIMEApplication(fp.read(), _subtype="pdf")
    fp.close()
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    server = smtplib.SMTP('smtp.mail.ru: 587')
    server.starttls()
    server.login(sender_email, password)
    while True:
        to = ws.cell(row=n, col=4).value
        if to != None:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = to
            msg['Subject'] = "Регистрация на Кругосветку-2024"
            msg.attach(MIMEText(generate_message_text(n, ws, ws1), 'html'))
            msg.attach(att)
            server.sendmail(sender_email, to, msg.as_string())
            print(n)
            n = n + 1
            with open('lastID.txt', 'r+', encoding='utf-8') as f:
                f.seek(0)
                f.truncate(0)
                f.write(str(n))
                f.close()
            await asyncio.sleep(60)
        else:
            print('zzZ')
            await asyncio.sleep(600)


# Press the green button in the gutter to run the script.
loop = asyncio.get_event_loop()
asyncio.ensure_future(main())
asyncio.ensure_future(pay.pay())
loop.run_forever()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
