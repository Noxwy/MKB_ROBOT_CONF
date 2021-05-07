# -*- coding: utf-8 -*-
"""
Created on Fri May  7 22:17:16 2021

@author: Alexey
"""


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import random
import requests

url = 'https://service.nalog.ru/inn.do'# URL Робота

driver = 'chromedriver.exe'# Путь до драйвера

ls_inn = []# Массив ИНН

# Обработка результата из Excel и сохранение результата

def GetValueFromExcel():
    #Чтение файла 
   excel_data_df = pd.read_excel('ExcelParametr.xlsx', sheet_name='Sheet1')
   #print(excel_data_df)
   #Получение информации из столбца ФАМИЛИЯ 
   famil = excel_data_df['Фамилия'].tolist()
   
   #Получение информации из столбца ИМЯ
   name = excel_data_df['Имя'].tolist()
   
   #Получение информации из столбца ОТЧЕСТВО
   otch = excel_data_df['Отчество'].tolist()
   
   #Получение информации из столбца ДАТА РОЖДЕНИЯ
   date_brth = excel_data_df['Дата рождения'].dt.strftime('%d.%m.%Y').tolist()
   
   #Получение информации из столбца СЕРИЯ Паспорта 
   seria = excel_data_df['Серия'].tolist()
   
   #Получение информации из столбца НОМЕР Паспорта
   number = excel_data_df['Номер'].tolist()
   
   #Получение информации из столбца ДАТА ВЫДАЧИ ПАСПОРТА
   date_get = excel_data_df['Дата Выдачи'].dt.strftime('%d.%m.%Y').tolist()
   
   #inn  = excel_data_df['ИНН'].tolist() 
   
   #Цикл получения инн с сайта
   count = 0
   while count < len(famil):
       
      #Запуск браузера (Создание обектв browser)
      browser = webdriver.Chrome(driver)
      #Вызов функции для работы с сайтом
      result = service(browser, famil[count], name[count],
              otch[count], date_brth[count], date_get[count], 
              str(seria[count])+str(number[count]))
      
      #Обработка полученной информации, если нашел то вернет номер ИНН, если нет, то вернет Nan
      if 'Информация об ИНН не найдена' in result or 'Ошибка на сайте' in result:
          ls_inn.append('Nan')
      else:
          ls_inn.append(result.split('ИНН:')[1].strip())
          
      #Закрытие браузера
      browser.close()
      count+=1
      
   #Создание новой таблицы 
   data = {'Фамилия':famil ,
           'Имя': name,
           'Отчество':otch,
           'Дата рождения':date_brth,
           'Серия':seria,
           'Номер':number,
           'Дата Выдачи':date_get,
           'ИНН':ls_inn}
   df = pd.DataFrame(data) 
   
   #Запись в Excel
   df.to_excel("output.xlsx",index=False)
      
def service(browser, surname, name, otch, datebrth, datepassport, passport):
   browser.get(url)
   try:
       # Отправка фамилии 
       senkey(browser, '//*[@id="fam"]', surname)
   except:
       
       # Нажатие клавишц продолжить, а потом робот вводит фамилию
       browser.find_element_by_xpath('//*[@id="unichk_0"]').click()
       browser.find_element_by_xpath('//*[@id="btnContinue"]').click()
       time.sleep(1)
       senkey(browser, '//*[@id="fam"]', surname)   
       
       
   #senkey(browser, '//*[@id="fam"]', surname)
        # Отправка имени
   senkey(browser, '//*[@id="nam"]', name)
    # Отправка отчесвта 
   senkey(browser, '//*[@id="otch"]', otch)
    # Отправка даты рождения
   senkey(browser, '//*[@id="bdate"]', datebrth)
    # Отправка серии и номера
   senkey(browser, '//*[@id="docno"]', passport)
    # Отправка даты выдачи паспорта
   senkey(browser, '//*[@id="docdt"]', datepassport)
    # Отправка фамилии 
   browser.find_element_by_xpath('//*[@id="btn_send"]').click()
   c = 0 
   while c<20: 
       
       try:
          # Поиск ИНН, если нашли, то вренет ИНН, если нет, то ждем 1 секунду и повторяем поиск ИНН
          # Повторяем операцию 20 раз, если после 20 раз нет результата, то возвращаем информацию с сайта 
           result = browser.find_element_by_xpath('//*[@id="result_1"]/div').text
           if result == '':
               time.sleep(1)
               c+=1
           else:    
               print('\nПоложительный Результат: '+result)
               return result
           #browser.quit()
               break
       except:
           time.sleep(1)
           c+=1
           
           
       try:
           result  = browser.find_element_by_xpath('//*[@id="result_0"]/div').text
           if result == '':
               time.sleep(1)
               c+=1
           else:    
               print('\nОтрицательный Результат: '+result)
               return result
           #browser.quit()
               break           
       except:
           time.sleep(1)
           c+=1 
           
   #Проверка на загрузку сайта                  
   reques = check_Web (browser,'/html/body/div[6]')
   if reques == 1:
       print('\nРезультат с сайта не загрузился !')
       return 'Ошибка на сайте'
       
#Проверка на загрузку сайта
def check_Web(br, str_xpath):
    try:
        br.find_element_by_xpath(str_xpath).text
        return 1
    except:
        return 0
#Функция отправки слов на сайт    
def senkey(brw, str_xpath, str_send):
   sen = brw.find_element_by_xpath(str_xpath)
   for i in str_send:
      sen.send_keys(i)
      time.sleep(0.001)


GetValueFromExcel()