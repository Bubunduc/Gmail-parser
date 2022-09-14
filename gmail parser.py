from selenium import webdriver
from selenium.webdriver.common.by import By
import asyncio

import aiohttp
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter
import os
import requests
class Email_parser():
    def __init__(self,url,reqs,names):
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        self.driver = webdriver.Chrome(options = options)
        self.driver.get(url)
        self.url = url
        self.reqs = reqs
        self.names = names
    def click(self,ad):
        line = self.driver.find_element(By.TAG_NAME,'input').send_keys(ad)
        but = self.driver.find_element(By.TAG_NAME,'button').click()
    def entrepreneur(self):
        k = 0


        while True:
            try:
                try:
                    isliquided = self.driver.find_element(By.XPATH,'//div[@class="flex-grow"][1]').text

                    if 'Ликвидировано' in isliquided:
                        entrepreneur = self.driver.find_elements(By.XPATH, '//div[@class="flex-grow"]/a')[2].get_attribute('href')
                        return entrepreneur
                    else:
                        entrepreneur = self.driver.find_element(By.XPATH, f'//div[@class="flex-grow"]/a').get_attribute('href')
                        return entrepreneur
                except:
                    entrepreneur = self.driver.find_element(By.XPATH,f'//div[@class="flex-grow"]/a').get_attribute('href')
                    return entrepreneur

            except:
                if k == 5:
                    return ''
                k+=1
                self.driver.refresh()
                continue


    def business_links(self,url):
        site = requests.get(url).text
        soup = BeautifulSoup(site, 'html.parser')
        links = soup.find('li')
        links = links.find_all('a')
        links = [i['href'] for i in links]
        links = ['https://e-ecolog.ru/'+i for i in links]
        links.append(self.director())



        return links

    async def call_url(self,url):
        async with aiohttp.ClientSession() as session:

            async with session.get(url) as response:
                data = await response.text()
                soup = BeautifulSoup(data, 'html.parser')
                mails = soup.find_all(itemprop="email")
                if mails == []:
                    return ''
                else:
                    mails = [i.string for i in mails]
                    return mails
    def asynced_search(self,links):
        lins = []
        for url in links:
            try:
                if '//person' in url:
                    url = url.replace('/person','entity')
                    lins.append(url)
                else:
                    url = url.replace('person', 'entity')
                    lins.append(url)

            except:
                lins.append(url)
        loop = asyncio.get_event_loop()
        futures = [loop.create_task(self.call_url(url)) for url in lins]
        loop.run_until_complete(asyncio.gather(*futures))
        results = [task.result() for task in futures]
        results =  [x for x in results if x!='']
        return results
    def director(self):
        director_mail = self.driver.find_elements(By.CLASS_NAME,'bl')[1]
        return director_mail.get_attribute('href')

    def x(self):
        all_mails = []
        for req in self.reqs:
            self.click(req)

            url = self.entrepreneur()
            if url == '':
                self.driver.get(self.driver.current_url)
                url = self.driver.current_url

            else:

                self.driver.get(url)

            try:

                links = self.business_links(url)
                main_adress = self.driver.find_elements(By.PARTIAL_LINK_TEXT, '@')[:-1]

                main_adress = [i.text for i in main_adress]
                ready_mails = []

                mails = self.asynced_search(links)
                for ma in mails:
                    for r in ma:
                        ready_mails.append(r)
                for main in main_adress:
                    ready_mails.append(main)
                ready_mails = set(ready_mails)
                ready_mails = list(ready_mails)
                print(ready_mails)
                all_mails.append(', '.join(ready_mails))
                self.driver.get(self.url)
            except:
                all_mails.append(' ')
                self.driver.get(self.url)
        self.driver.close()
        self.driver.quit()
        
        return all_mails





file_name = input('Введите название файла: \n')
cols = [9, 10]
requisits = pd.read_excel(f'{file_name}.xlsx',usecols=cols,header=None)
adresses = requisits[10].tolist()[1:]
names = requisits[9].tolist()[1:]
reqs = []
for i in range(0,len(adresses)):
    reqs.append(f'{names[i]} {adresses[i]}')
f = Email_parser('https://e-ecolog.ru/',reqs,names)


emails = f.x()


os.remove(f'{file_name}.xlsx')
workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(f'J1', 'Компания')
worksheet.write(f'K1', 'Адрес')
worksheet.write(f'L1', 'Электронка')
for i,(name) in enumerate(names,start = 2):
    worksheet.write(f'J{i}', name)
for i,(adress) in enumerate(adresses,start = 2):
    worksheet.write(f'K{i}',str(adress))
for i,(email) in enumerate(emails,start = 2):
    worksheet.write(f'L{i}', email)
workbook.close()

r = input('Парсинг завершен!!!')




