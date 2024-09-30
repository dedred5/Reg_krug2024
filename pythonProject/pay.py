import smtplib
import asyncio
import gspread
from gspread import Client, Spreadsheet
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

SPREADSHEET_URL1 = "https://docs.google.com/spreadsheets/d/1Aht0ipDovye2nAx_CC0J5UEKHaUL-f0Al0Yu2Z3Go5Y/edit?resourcekey=&gid=1244274127#gid=1244274127"
SPREADSHEET_URL2 = "https://docs.google.com/spreadsheets/d/1-2fMMxxau3QKULV9yo_POBMdrehzN2a-ipcGiOaaGVc/edit?resourcekey=&gid=462629804#gid=462629804"
sender_email = "dendrod5@mail.ru"
receiver_email = "dendrod5@mail.ru"
password = "TSpyuJ4ZX1T8CuJvZ1hp"
gs: Client = gspread.service_account("./Serv-Acc.json")
sh1: Spreadsheet = gs.open_by_url(SPREADSHEET_URL1)
sh2: Spreadsheet = gs.open_by_url(SPREADSHEET_URL2)
ws1 = sh1.sheet1
ws2 = sh2.sheet1
ws3 = sh2.worksheet(title='Места на теплоход')

def generate_text():
    htmlBody = '<p>Поздравляем, вы успешно оплатили стартовый взнос на Кругосветку-2024!</p> <p>Будем ждать вас на старте вашего маршрута</p> <p>Приятных путешествий!</p>'
    return htmlBody

def generate_error():
    htmlBody = '<p>К сожалению, выбранный вами рейс на теплоход полностью заполнен и вам не хватит места, поэтому предлагаю выбрать другой рейс.</p><p>Для этого необходимо ещё раз ответить на форму по ссылке: https://forms.gle/VXaRP9V3qFJMGKBn8</p>'
    return htmlBody

def generate_error2(to, lastteg, ws2, ws1):
    htmlBody = '<p>Ваша оплата уже подтверждена</p><h4><u>Регистрационные данные:</u></h4><ol>'
    # ---НОМЕР---
    htmlBody += '<li>Индивидуальный номер и номер карты участника: ' + str(to) + '</li>'
    # ---ДИСТАНЦИЯ---
    if ws2.cell(row=to, col=13).value == None:
        dist = ws2.cell(row=to, col=14).value
    else:
        dist = ws2.cell(row=to, col=13).value
    htmlBody += '<li>Дистанция: ' + dist + '</li>'
    # ---КОЛИЧЕСТВО ЧЕЛОВЕК ---
    if (ws2.cell(row=to, col=5).value == 'Групповая'):
        htmlBody += '<li>Количество человек в группе: ' + str(ws2.cell(row=to, col=8).value) + '</li>'
    # ---МЕСТО СТАРТА - --
    htmlBody += '<li>Место старта: ' + ws2.cell(row=to, col=12).value + '</li>'
    if ws1.cell(row=lastteg, col=8).value != None or ws1.cell(row=lastteg, col=9).value != None:
        if ws1.cell(row=lastteg, col=8).value != None:
            htmlBody += '<li>Рейс теплохода: ' + ws1.cell(row=lastteg, col=8).value + '</li></ol>'
        else:
            htmlBody += '<li>Рейс теплохода: ' + ws1.cell(row=lastteg, col=9).value + '</li></ol>'
    else:
        htmlBody += '</ol>'
    return htmlBody

def send_error(To, server):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = To
    msg['Subject'] = "Регистрация на Кругосветку-2024"
    msg.attach(MIMEText(generate_error(), 'html'))
    server.sendmail(sender_email, To, msg.as_string())

def send_text(To, server):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = To
    msg['Subject'] = "Регистрация на Кругосветку-2024"
    msg.attach(MIMEText(generate_text(), 'html'))
    server.sendmail(sender_email, To, msg.as_string())

def send_error2(To, lastteg, to, server, ws2, ws1):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = To
    msg['Subject'] = "Регистрация на Кругосветку-2024"
    msg.attach(MIMEText(generate_error2(to, lastteg, ws2, ws1), 'html'))
    server.sendmail(sender_email, To, msg.as_string())
async def pay():
    server = smtplib.SMTP('smtp.mail.ru: 587')
    server.starttls()
    server.login(sender_email, password)
    with open('lastteg.txt', encoding='utf-8') as f:
        lastteg = int(f.read())
    while True:
        to = ws1.cell(row=lastteg, col=2).value
        if to != None:
            to = int(to)
            To = ws2.cell(row=to, col=4).value
            if ws2.cell(row=to, col=15).value == None and To != None:
                if ws1.cell(row=lastteg, col=8).value == None:
                    rey = ws1.cell(row=lastteg, col=9).value
                else:
                    rey = ws1.cell(row=lastteg, col=8).value
                if rey != None:
                    rey = int(rey[0]) + 3
                    if int(ws1.cell(row=lastteg, col=3).value) > int(ws3.cell(row=rey, col=3).value):
                        print(1, lastteg)
                        send_error(To, server)
                        ws1.update_cell(row=lastteg, col=10, value='1')
                    else:
                        if int(ws1.cell(row=lastteg, col=3).value) == 1:
                            kol = int(ws3.cell(row=rey, col=6).value) + 1
                            ws3.update_cell(row=rey, col=6, value=kol)
                        else:
                            kol = int(ws3.cell(row=rey, col=7).value) + int(ws1.cell(row=lastteg, col=3).value)
                            ws3.update_cell(row=rey, col=7, value=kol)
                        send_text(To, server)
                        ws2.update_cell(row=to, col=15, value='2')
                else:
                    print('ok')
                    send_text(To, server)
                    ws2.update_cell(row=to, col=15, value='2')
            else:
                ws1.update_cell(row=lastteg, col=10, value='2')
                send_error2(To, lastteg, to, server, ws2, ws1)
            lastteg += 1
            with open('lastteg.txt', 'r+', encoding='utf-8') as f:
                f.seek(0)
                f.truncate(0)
                f.write(str(lastteg))
                f.close()
            print(lastteg)
            await asyncio.sleep(90)
        else:
            print('zzZp')
            server.close()
            await asyncio.sleep(600)
            server = smtplib.SMTP('smtp.mail.ru: 587')
            server.starttls()
            server.login(sender_email, password)
