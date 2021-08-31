import xlsxwriter
import sqlite3
from DataBase import DData
import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QMessageBox
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


# Отнаследуем наш класс от простейшего графического примитива QWidget
class Example(QWidget):
    def __init__(self):
        # Надо не забыть вызвать инициализатор базового класса
        super().__init__()
        self.filename = ''
        # В метод initUI() будем выносить всю настройку интерфейса,
        # чтобы не перегружать инициализатор
        self.initUI()

    def initUI(self):
        self.start()

    def start(self):
        text, ok = QInputDialog.getText(
            self, 'Сохранение', 'Введите название файла: ')

        if ok:
            self.correct_filename(text)

    def check_this_filename(self):
        files = os.listdir()
        if self.path in files:
            print('я тут')
            QMessageBox.critical(self, "Ошибка ", "Файл с таким именем уже существует!", QMessageBox.Ok)
            self.start()

    def correct_filename(self, text):
        flag = bool(text.count('.xlsx'))
        if not flag:
            self.filename = text + '.xlsx'
        else:
            self.filename = text
        self.save_file()

    def save_file(self):
        con = sqlite3.connect('db/projects.db')
        cur = con.cursor()

        con1 = sqlite3.connect('db/users.db')
        cur1 = con1.cursor()
        a = [job for job in cur.execute("SELECT * FROM projects")]
        b = [job for job in cur1.execute("SELECT * FROM users")]
        cur.close()
        cur1.close()
        l = list()
        for i in a:
            i1 = i
            user_id = int(i1[-1])
            for j in b:
                if j[0] == user_id:
                    i1 += j
                    l.append(i1)
                    break

        workbook = xlsxwriter.Workbook(self.filename)
        worksheet = workbook.add_worksheet()

        abc = ['ID проекта', 'ID прокета в базе TatNeft', 'Ссылка на проект', 'Название проекта', 'Роль в проекте',
               'О проекте',
               'Заслуги в проекте',
               'Связка проект-Юзер', 'ID Пользователя', 'Имя', 'Фамилия', 'Отчество',
               'Табельный номер', 'Почта', 'Мобильный Номер', 'Образование', 'Должность', 'Компания', 'File_Path']

        row, col = 0, 0

        for i in abc:
            worksheet.set_column(row, col, 20)
            worksheet.write(row, col, i)
            col += 1

        for row, (
                id_pr, r_pr_id, pr_url, pr_name, pr_rol, pr_cont, pr_att, user_id, real_user_id, name, surname, father,
                tablenum, email, mobile, obr, prof, company) in enumerate(l):
            path = DData().dirname((name, surname, father, user_id))
            worksheet.write(row + 1, 0, str(id_pr))
            worksheet.write(row + 1, 1, r_pr_id)
            worksheet.write(row + 1, 2, pr_url)
            worksheet.write(row + 1, 3, pr_name)
            worksheet.write(row + 1, 4, pr_rol)
            worksheet.write(row + 1, 5, pr_cont)
            worksheet.write(row + 1, 6, pr_att)
            worksheet.write(row + 1, 7, str(user_id))
            worksheet.write(row + 1, 8, str(real_user_id))
            worksheet.write(row + 1, 9, name)
            worksheet.write(row + 1, 10, surname)
            worksheet.write(row + 1, 11, father)
            worksheet.write(row + 1, 12, str(tablenum))
            worksheet.write(row + 1, 13, email)
            worksheet.write(row + 1, 14, mobile)
            worksheet.write(row + 1, 15, obr)
            worksheet.write(row + 1, 16, prof)
            worksheet.write(row + 1, 17, company)

            worksheet.write_url(row + 1, 18, f'external:{path}')
        workbook.close()
        send_from = # your email
        send_to = # recipient's address
        subject = # topic
        text = f'Выслано {formatdate(localtime=True)}'
        self.send_mail(send_from, send_to, subject, text, 'smtp.gmail.com', 587, password= # your email password)

    def send_mail(self, send_from, send_to, subject, text, server, port, password='', isTls=True):
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(self.filename, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={self.filename}')
        msg.attach(part)

        # context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
        # SSL connection only working on Python 3+
        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(send_from, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()


if __name__ == '__main__':
    # Создадим класс приложения PyQT
    app = QApplication(sys.argv)
    # А теперь создадим и покажем пользователю экземпляр
    # нашего виджета класса Example
    ex = Example()
    # ex.show()
    # Будем ждать, пока пользователь не завершил исполнение QApplication,
    # а потом завершим и нашу программу
    sys.exit()
