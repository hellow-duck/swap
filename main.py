import tkinter as tk
from tkinter import ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Загрузка таблицы из Excel файла
df = pd.read_excel('./updated_data.xlsx')

def countryInterest(country_name):
    options = Options()
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'  # Путь к Firefox
    options.add_argument('--headless')
    service = Service(r'E:\driver\geckodriver.exe')  # Путь к geckodriver
    driver = webdriver.Firefox(service=service, options=options)

    rate = None  # Инициализация переменной

    try:
        country_url = df.loc[df['country'].str.lower() == country_name.lower(), 'page'].values

        if len(country_url) > 0:
            print(f'Ссылка для {country_name}: {country_url[0]}')
            driver.get(country_url[0])

            # Ожидание загрузки элемента с процентной ставкой
            rate_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.table-responsive:nth-child(3) > table:nth-child(1) > tbody:nth-child(2) > tr:nth-child(1) > td:nth-child(2)"))
            )

            rate = rate_element.text
            print(f'Процентная ставка для {country_name}: {rate}')
        else:
            print(f'Страна "{country_name}" не найдена в таблице.')

    except Exception as e:
        print(f'Произошла ошибка: {e}')
        print(driver.page_source)  # Для отладки

    finally:
        driver.quit()  # Закрытие драйвера

    return rate

def on_currency_change(event):
    selected_currency = event.widget.get()
    country_data = countries_df[countries_df['country'] == selected_currency]
    
    if not country_data.empty:
        rate = countryInterest(selected_currency)
        if event.widget == combo_currency1:
            entry_rate1.delete(0, tk.END)
            entry_rate1.insert(0, rate)
        elif event.widget == combo_currency2:
            entry_rate2.delete(0, tk.END)
            entry_rate2.insert(0, rate)

def calculate_swap():
    currency1 = combo_currency1.get()
    currency2 = combo_currency2.get()
    
    rate1 = float(entry_rate1.get())
    rate2 = float(entry_rate2.get())
    
    msp1 = df.loc[df['country'] == currency1, 'MSP'].values[0]
    msp2 = df.loc[df['country'] == currency2, 'MSP'].values[0]
    
    koef1 = rate1 * msp1
    koef2 = rate2 * msp2
    
    swap_buy = koef1 - koef2
    label_swap.config(text=f"Своп: {swap_buy:.2f}")


# Загрузка базы данных стран
countries_df = pd.read_excel('updated_data.xlsx')

# Создаем основное окно
root = tk.Tk()
root.title("Своп Калькулятор")

# Выпадающий список валют 1
label_currency1 = tk.Label(root, text="Выберите первую валюту:")
label_currency1.pack()

combo_currency1 = ttk.Combobox(root, values=countries_df['country'].tolist())
combo_currency1.pack()
combo_currency1.bind("<<ComboboxSelected>>", on_currency_change)

# Выпадающий список валют 2
label_currency2 = tk.Label(root, text="Выберите вторую валюту:")
label_currency2.pack()

combo_currency2 = ttk.Combobox(root, values=countries_df['country'].tolist())
combo_currency2.pack()
combo_currency2.bind("<<ComboboxSelected>>", on_currency_change)

# Поля для отображения процентных ставок
label_rate1 = tk.Label(root, text="Процентная ставка для первой валюты:")
label_rate1.pack()

entry_rate1 = tk.Entry(root)
entry_rate1.pack()

label_rate2 = tk.Label(root, text="Процентная ставка для второй валюты:")
label_rate2.pack()

entry_rate2 = tk.Entry(root)
entry_rate2.pack()

# Кнопка для расчета свопа
btn_calculate = tk.Button(root, text="Рассчитать своп", command=calculate_swap)
btn_calculate.pack()

# Поле для вывода рассчитанного свопа
label_swap = tk.Label(root, text="Своп: ")
label_swap.pack()

# Запуск основного цикла приложения
root.mainloop()
