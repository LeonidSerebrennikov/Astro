import sys
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QMessageBox, QPushButton, QApplication, QHeaderView
from PyQt5.uic import loadUi
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from docxtpl import DocxTemplate
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np


class AdminInterface(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("MainForm.ui", self)  # Загружаем интерфейс из файла mainform.ui
        self.data = self.load_data_from_google_sheet()
        self.initUI()

        self.setupTable()

    def initUI(self):
        deny_button_style = "background-color: rgba(200, 200, 200); border-radius: 8px; border: 1px solid #0d6efd; padding: 5px 15px; margin-top: 10px; font: 10pt; font-weight: 50; font-weight: normal;"

        self.deny_button.setStyleSheet(deny_button_style)
        self.setFixedSize(1398, 653)
        self.table = self.findChild(QTableWidget, "tableWidget")
        self.table.setColumnCount(len(self.data.columns))
        self.table.setRowCount(len(self.data))
        self.table.setHorizontalHeaderLabels(self.data.columns)

        self.table.cellClicked.connect(self.cell_clicked)  # Связывание сигнала cellClicked с функцией cell_clicked
        self.setupTable()
        self.approve_button.clicked.connect(self.generate_report)  # Привязка функции generate_report() к нажатию кнопки "Утвердить"
        self.deny_button.clicked.connect(self.reject_row)  # Привязка функции reject_row() к нажатию кнопки "Отклонить"
        self.refresh_button.clicked.connect(self.refresh_table)  # Привязка функции refresh_table() к нажатию кнопки "Обновить"

    def load_data_from_google_sheet(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('D:/forms/loyal-bit-424108-u8-d0c1fcab4138.json', scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1TUm4ULKwdQw0-3wGsLqcwFsL_VNmyih8Q2bLrfElgQc')
        self.worksheet = sheet.get_worksheet(0)
        data = self.worksheet.get_all_records()
        return pd.DataFrame(data)

    def setupTable(self):
        for i in range(len(self.data)):
            for j in range(len(self.data.columns)):
                item = QTableWidgetItem(str(self.data.iloc[i, j]))
                self.table.setItem(i, j, item)

    def cell_clicked(self, row, column):
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            item.setSelected(True)

        self.selected_row_index = row


    def generate_report(self):
        if hasattr(self, 'selected_row_index'):
            selected_row_data = list(self.data.iloc[self.selected_row_index])
            printApp(selected_row_data)
            QMessageBox.information(self, "Notification", "Report generated successfully")
        else:
            QMessageBox.warning(self, "Warning", "Select a row before generating the report.")

    def reject_row(self):
        if hasattr(self, 'selected_row_index'):
            user_email = self.data.iloc[self.selected_row_index, 2]  # Получаем адрес электронной почты
            req_date = self.data.iloc[self.selected_row_index, 0]  # Получаем дату заявки
            subject = "Заявка отклонена"
            message = "К сожалению, Ваша заявка на астрономическое наблюдение от " + \
                      str(req_date) + " была отклонена."
            msg = MIMEMultipart()
            msg['From'] = "leonidargentum@gmail.com"  # Мой адрес электронной почты
            msg['To'] = user_email
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'plain'))

            try:
                with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Замените настройки на свои SMTP-сервера
                    server.starttls()
                    server.login("leonidargentum@gmail.com", "xatz rucc hhdd hpmz")  # Логин и пароль для отправки писем
                    text = msg.as_string()
                    server.sendmail("leonidargentum@gmail.com", user_email, text)  # Отправка письма
                    QMessageBox.information(self, "Уведомление", f"Письмо успешно отправлено на {user_email}")
            except Exception as e:
               QMessageBox.warning(self, "Предупреждение", f"Не удалось отправить письмо: {e}")
        else:
            QMessageBox.warning(self, "Предупреждение", "Выберите строку перед отклонением.")



    def refresh_table(self):
        self.data = self.load_data_from_google_sheet()
        self.setupTable()


def printApp(data_row):
    try:
        doc = DocxTemplate("шаблон_заявки.docx")
        current_date = date.today()

        context = {
            'current_date': current_date,
            'name': data_row[3],
            'target': data_row[1],
            'observation_type': data_row[4],
            'object_type': data_row[10],
            'early_date': data_row[5],
            'late_date': data_row[6],
            'duration': data_row[7],
            'redshift': data_row[8],
            'v_magnitude': data_row[9],
            #'number': data_row[8]
            'comment': data_row[11],
            'email': data_row[2]
        }

        doc.render(context)
        doc.save("Заявка " + str(data_row[1]) + " " + str(current_date) + ".docx")

    except Exception as e:
        print(f"Ошибка при генерации отчета: {e}")
        print(data_row)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AdminInterface()

    window.show()
    sys.exit(app.exec_())