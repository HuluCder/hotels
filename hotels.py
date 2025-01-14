import rpa as r
import pandas as pd
import time
import os

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

#Выбор дат заезда и выезда
if r.present('//input[@aria-label="Заезд"]'):
    r.click('//input[@aria-label="Заезд"]')
    time.sleep(1)
    r.click('//div[@aria-label="вторник, 14 января 2025 г."]')
    r.click('//div[@aria-label="понедельник, 20 января 2025 г."]')
    r.click('//span[contains(text(), "Готово")]')
    time.sleep(5)
else:
    print("[-] Поле выбора даты не найдено.")
    r.close()
    exit()

#Сбор информации о первых 14 отелях
hotel_data = []

for i in range(1, 16):
    xpath_hotel = f'(//h2[@jscontroller="bqejFf" and @jsaction="YcW9n:dDUAne;"])[{i}]'
    if r.present(xpath_hotel):
        r.click(xpath_hotel)
        time.sleep(5)

        hotel_name = r.read('//h1[@class="FNkAEc o4k8l"]')
        hotel_address = r.read('//div[@class="K4nuhf"]//span[@class="CFH2De"]')
        hotel_price = r.read(f'(//h2[@jscontroller="bqejFf" and @jsaction="YcW9n:dDUAne;"])[{i}]//following::span[contains(@class, "qQOQpe") and contains(text(), "₸")][1]')
        hotel_booking_link = r.url()

        hotel_data.append({
            'Название': hotel_name,
            'Адрес': hotel_address,
            'Цена': hotel_price,
            'Ссылка на бронь': hotel_booking_link
        })

        # Нажатие на кнопку возврата к списку отелей
        if r.present('//span[@jsname="UDkh2"]'):
            r.click('//span[@jsname="UDkh2"]')
            time.sleep(5)
        else:
            print("[-] Кнопка возврата не найдена.")
            r.close()
            exit()
    else:
        print(f"[-] Отель №{i} не найден.")
        break

#Сохранение данных в DataFrame
df = pd.DataFrame(hotel_data)

#Сохранение в Excel
file_path = os.path.join(os.getcwd(), 'гостиницы.xlsx')
df.to_excel(file_path, index=False)

print(f"[+] Данные успешно сохранены в {file_path}")

#Завершение работы
r.close()
