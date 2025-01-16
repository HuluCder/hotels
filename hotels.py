import rpa as r
import pandas as pd
import time
import os
from datetime import datetime, timedelta
import locale

# Установка русской локали для корректного отображения дат
try:
    locale.setlocale(locale.LC_TIME, 'UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'Russian_Russia.1251')  # Для Windows

# Определение дат заезда и выезда в нужном формате
check_in_datetime = datetime.now() + timedelta(days=1)
check_out_datetime = check_in_datetime + timedelta(days=6)

check_in_date = check_in_datetime.strftime('%d')
check_out_date = check_out_datetime.strftime('%d')
month_genitive_check_in = check_in_datetime.strftime('%B')
month_genitive_check_out = check_out_datetime.strftime('%B')

months_genitive = {
    'январь': 'января', 'февраль': 'февраля', 'март': 'марта',
    'апрель': 'апреля', 'май': 'мая', 'июнь': 'июня',
    'июль': 'июля', 'август': 'августа', 'сентябрь': 'сентября',
    'октябрь': 'октября', 'ноябрь': 'ноября', 'декабрь': 'декабря'
}

month_genitive_check_in = months_genitive.get(month_genitive_check_in.lower(), month_genitive_check_in)
month_genitive_check_out = months_genitive.get(month_genitive_check_out.lower(), month_genitive_check_out)

print("Дата заезда:", check_in_date, month_genitive_check_in)
print("Дата выезда:", check_out_date, month_genitive_check_out)

r.init(visual_automation=True, chrome_browser=True)

# Поиск "Гостиницы Сочи"
r.url('https://www.google.com')
time.sleep(3)
r.type('//*[@name="q"]', 'Гостиницы Сочи')
r.click('//input[@name="btnK"]')
time.sleep(5)

# Нажатие на кнопку "Посмотреть"
if r.present('//span[contains(text(), "Ещё")]'):
    r.click('//span[contains(text(), "Ещё")]')
    time.sleep(5)
else:
    print("[-] Кнопка 'Посмотреть' не найдена.")
    r.close()
    exit()

# Выбор дат заезда и выезда
if r.present('//input[@aria-label="Заезд"]'):
    r.click('//input[@aria-label="Заезд"]')
    time.sleep(2)
    
    # Нажатие на элементы с числом и месяцем
    if r.present(f'//div[contains(@aria-label, "{check_in_date} {month_genitive_check_in}")]'):
        r.click(f'//div[contains(@aria-label, "{check_in_date} {month_genitive_check_in}")]')
    else:
        print(f"[-] Дата заезда не найдена: {check_in_date} {month_genitive_check_in}")
        r.close()
        exit()
    
    if r.present(f'//div[contains(@aria-label, "{check_out_date} {month_genitive_check_out}")]'):
        r.click(f'//div[contains(@aria-label, "{check_out_date} {month_genitive_check_out}")]')
    else:
        print(f"[-] Дата выезда не найдена: {check_out_date} {month_genitive_check_out}")
        r.close()
        exit()

    r.click('//span[contains(text(), "Готово")]')
    time.sleep(5)
else:
    print("[-] Поле выбора даты не найдено.")
    r.close()
    exit()

# Сбор информации о первых 14 отелях
hotel_data = []

for i in range(1, 15):
    xpath_hotel = f'(//h2[@jscontroller="bqejFf" and @jsaction="YcW9n:dDUAne;"])[{i}]'
    if r.present(xpath_hotel):
        r.click(xpath_hotel)

        hotel_name = r.read('//h1[@class="FNkAEc o4k8l"]')
        hotel_address = r.read('//div[@class="K4nuhf"]//span[@class="CFH2De"]')
        # Попытка взять цену из основного элемента, иначе из запасного
        if r.present(f'(//button[contains(@class, "VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ nCP5yc AjY5Oe DuMIQc LQeN7 A9rngd VqBuDc wNqaKc idHpEf")]//span[@class="W9vOvb nDkDDb"])'):
            hotel_price = r.read(f'(//button[contains(@class, "VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ nCP5yc AjY5Oe DuMIQc LQeN7 A9rngd VqBuDc wNqaKc idHpEf")]//span[@class="W9vOvb nDkDDb"])')
        else:
            hotel_price = r.read(f'(//button[contains(@class, "VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ nCP5yc AjY5Oe DuMIQc LQeN7 A9rngd VqBuDc wNqaKc idHpEf")]//span[@class="qQOQpe prxS3d"])')
        hotel_booking_link = r.read('//a[contains(@class, "WpHeLc") and contains(@aria-label, "Посетить сайт партнера")]//@href')

        hotel_data.append({
            'Название': hotel_name,
            'Адрес': hotel_address,
            'Цена': hotel_price,
            'Ссылка на бронь': hotel_booking_link
        })

        # Нажатие на кнопку возврата к списку отелей
        if r.present('//span[@jsname="UDkh2"]'):
            r.click('//span[@jsname="UDkh2"]')
        else:
            print("[-] Кнопка возврата не найдена.")
            r.close()
            exit()
    else:
        print(f"[-] Отель №{i} не найден.")
        break

# Сохранение данных в DataFrame
df = pd.DataFrame(hotel_data)

# Сохранение в Excel
file_path = os.path.join(os.getcwd(), 'гостиницы.xlsx')
df.to_excel(file_path, index=False)

print(f"[+] Данные успешно сохранены в {file_path}")

# Завершение работы
r.close()
