import os
import random
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import time
import sys


drivers = {
    "Chrome": webdriver.Chrome,
    "Edge": webdriver.Edge,
    "Firefox": webdriver.Firefox,
    "Safari": webdriver.Safari,
    "Internet Explorer": webdriver.Ie
}

driver = None

for browser_name, browser in drivers.items():
    try:
        driver = browser()
        break
    except WebDriverException:
        pass

if driver is None:
    sys.exit("Не удается запустить драйвер браузера")

driver.get("https://www.moex.com/")

WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
    (By.XPATH, "//button[contains(@class, 'menu-button--mobile')]"))).click()

WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH, "//a[text()='Рынки']"))).click()
WebDriverWait(driver, 20).until(
    expected_conditions.element_to_be_clickable((By.XPATH, "//a[text()='Срочный рынок']"))).click()

try:
    WebDriverWait(driver, 20).until(
        expected_conditions.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Согласен')]"))).click()
except NoSuchElementException:
    pass

try:
    WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
        (By.XPATH, "//div[contains(@class, 'left-menu__mobile-header')]"))).click()
except NoSuchElementException:
    pass

WebDriverWait(driver, 20).until(
    expected_conditions.element_to_be_clickable((By.XPATH, "//span[text()='Индикативные курсы']"))).click()

WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
    (By.XPATH, "//span[contains(@class, 'ui-icon ui-select__icon -arrow')]"))).click()

WebDriverWait(driver, 20).until(expected_conditions.visibility_of_element_located(
    (By.XPATH, "//div[@class='ui-dropdown -opened']")))

xpath_list = ["//a[@href='?currency=CHF_RUB' and contains(text(), 'CHF/RUB - Швейцарский франк к российскому рублю')]",
              "//a[@href='?currency=CNY_RUB' and contains(text(), 'CNY/RUB - Китайский юань к российскому рублю')]",
              "//a[@href='?currency=GBP_RUB' and contains(text(), 'GBP/RUB - Британский фунт к российскому рублю')]",
              "//a[@href='?currency=CHF_RUB' and contains(text(), 'CHF/RUB - Швейцарский франк к российскому рублю')]",
              "//a[@href='?currency=HKD_RUB' and contains(text(), 'HKD/RUB - Гонконгский доллар к российскому рублю')]",
              "//a[@href='?currency=INR_RUB' and contains(text(), 'INR/RUB – Индийская рупия к российскому рублю')]",
              "//a[@href='?currency=JPY_RUB' and contains(text(), 'JPY/RUB - Японская йена к российскому рублю')]",
              "//a[@href='?currency=KZT_RUB' and contains(text(), 'KZT/RUB - Казахский тенге к россискому рублю')]", # Эта опечатка на странице
              "//a[@href='?currency=TRY_RUB' and contains(text(), 'TRY/RUB - Турецкая лира к российскому рублю')]",
              "//a[@href='?currency=TRY_RUB' and contains(text(), 'UAH/RUB - Украинская гривна к российскому рублю')]"]

try:
    WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
        (By.XPATH,
         "//a[@href='?currency=USD_RUB' and contains(text(), 'USD/RUB - Доллар США к российскому рублю')]"))).click()
except NoSuchElementException:
    try:
        WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
            (By.XPATH,
             "//a[@href='?currency=CAD_RUB' and contains(text(), 'CAD/RUB - Канадский доллар к российскому рублю')]"))).click()
    except NoSuchElementException:
        while True:
            proba = random.choice(xpath_list)
            try:
                WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable((By.XPATH, proba))).click()
                break
            except NoSuchElementException:
                print(f"Элемент не найден: {proba}. Пытаемся снова...")  # Сообщение об ошибке

"""Переходы к таблицам данных можно/лучше реализовать через прямые переходы по URL, без переходов по меню,
так как при выборе клиринга открывается уникальная страница для каждой рублевой пары. Предполагаю, что это было бы рациональнее."""

WebDriverWait(driver, 20).until(expected_conditions.presence_of_element_located((By.XPATH, "//tbody")))


data = []
rows = driver.find_elements(By.XPATH, "//tbody/tr")

for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    if len(cells) == 5:
        date_usd = cells[0].text
        clearing_value_usd = cells[1].text
        clearing_time_usd = cells[2].text
        main_clearing_value_usd = cells[3].text
        main_clearing_time_usd = cells[4].text
        data.append(
            [date_usd, clearing_value_usd, clearing_time_usd, main_clearing_value_usd, main_clearing_time_usd, "", "",
             "", "", "", ""])

WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
    (By.XPATH, "//span[contains(@class, 'ui-icon ui-select__icon -arrow')]"))).click()

WebDriverWait(driver, 20).until(expected_conditions.visibility_of_element_located(
    (By.XPATH, "//div[@class='ui-dropdown -opened']")))

try:
    WebDriverWait(driver, 20).until(expected_conditions.element_to_be_clickable(
        (
            By.XPATH,
            "//a[@href='?currency=EUR_RUB' and contains(text(), 'EUR/RUB - Евро к российскому рублю')]"))).click()
except NoSuchElementException:
    pass  # Просто пропускаем, так как в условии не было альтернативного варианта, хотя можно добавить другую валюту.

time.sleep(2) # прогрузка
rows = driver.find_elements(By.XPATH, "//tbody/tr")
for i, row in enumerate(rows):
    cells = row.find_elements(By.TAG_NAME, "td")
    if len(cells) == 5:
        date_eur = cells[0].text
        clearing_value_eur = cells[1].text
        clearing_time_eur = cells[2].text
        main_clearing_value_eur = cells[3].text
        main_clearing_time_eur = cells[4].text
        if i < len(data):
            data[i][6] = date_eur
            data[i][7] = clearing_value_eur
            data[i][8] = clearing_time_eur
            data[i][9] = main_clearing_value_eur
            data[i][10] = main_clearing_time_eur

driver.quit()

workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

fin_format = workbook.add_format({'num_format': '"₽"#,##0.00'})
rez_format = workbook.add_format({'num_format': '#,##0.00'})

worksheet.write('A1', 'Дата')
worksheet.write('B1', 'Значение курса промежуточного клиринга')
worksheet.write('C1', 'Значение курса основного клиринга')

worksheet.write('E1', 'Дата')
worksheet.write('F1', 'Значение курса промежуточного клиринга')
worksheet.write('G1', 'Значение курса основного клиринга')

worksheet.write('H1', 'Изменение')


row_start = 1 # Запись данных, ниже шапки
for i, entry in enumerate(data):
    worksheet.write(row_start + i, 0, entry[0])
    worksheet.write(row_start + i, 1,
                    "₽" + f"{float(entry[1]):.2f}" + " " + entry[2])
    worksheet.write(row_start + i, 2,
                    "₽" + f"{float(entry[3]):.2f}" + " " + entry[4])
    worksheet.write(row_start + i, 3, '')
    worksheet.write(row_start + i, 4, entry[6])
    worksheet.write(row_start + i, 5,
                    "₽" + f"{float(entry[7]):.2f}" + " " + entry[8])
    worksheet.write(row_start + i, 6,
                    "₽" + f"{float(entry[9]):.2f}" + " " + entry[10])

    # Добавляем значение в столбец H (индекс 7)
    if entry[3] and entry[9]:
        value_h = float(entry[9]) / float(entry[3])
        worksheet.write(row_start + i, 7, value_h, rez_format)
    else:
        worksheet.write(row_start + i, 7, '-')

"""
Задание противоречиво в части разделения столбцов, если разделить столбцы клиринга на значение и время, 
то столбец 'Изменение' будет дальше H. Принял решение сохранить 'Изменение' в столбце H в соответствии 
с условием, так как это может быть использовано для последующей интеграции данных с другими системами.
Условие разделения реализовал в столбцах с клирингом, добавив к значениям время через пробел.
"""

worksheet.autofit()

workbook.close()

number_of_rows = len(data) + 1

def get_row_word(rows_count):
    if rows_count % 10 == 1 and rows_count % 100 != 11:
        return "строка"
    elif rows_count % 10 in [2, 3, 4] and not (rows_count % 100 in [12, 13, 14]):
        return "строки"
    else:
        return "строк"

def send_email(file_path, rows_count):
    from_email = "kulikovma@inbox.ru"
    to_email = "kulikovma@inbox.ru"
    subject = "Отчет по курсам"
    row_word = get_row_word(rows_count)
    body = f"В файле {rows_count} {row_word}."

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(file_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
        msg.attach(part)

        with smtplib.SMTP('smtp.mail.ru', 587) as server:
            server.starttls()
            server.login("kulikovma@inbox.ru", "ypRKeAc35fMmTDFw3bhn")
            server.send_message(msg)

send_email('report.xlsx', number_of_rows)
