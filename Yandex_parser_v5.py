from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time
import os

driver_path_chrome = r'driver_path'
web_path = 'https://yandex.ru/search/?lr=213&text='

PATH_FROM = 'path_from'
PATH_TO = 'path_to'

refs_count = 2
stop_word = 'ИНН'
search_criteria = ' инн rusprofile.ru'

files_to_scrap = os.listdir(PATH_FROM)
counter = 1
for file in files_to_scrap:
      text_list = []
      with open(os.path.join(PATH_FROM, file), 'r', encoding='UTF-8') as file:
            file_lines = len(file.readlines())
            file.seek(0)
            for i in range(file_lines):
                  company = file.readline().strip()
                  options = webdriver.ChromeOptions()
##                  options.add_argument('headless')
                  driver = webdriver.Chrome(driver_path_chrome, chrome_options=options)
                  driver.get(web_path + (company + search_criteria))

                  for i in range(refs_count):
                        try:
                              drv_obj = driver.find_element_by_xpath(
                                    "//li[@class='serp-item'][@data-cid={0}] \
                                    /div/h2[@class='organic__title-wrapper typo typo_text_l typo_line_m']".format(i))
                              text_value = drv_obj.text
                              if stop_word in text_value.upper():
                                    text_list.append(company + '::' + text_value)
                                    print(company + '::' + text_value)
                                    break
                              else:
                                    pass
                        except NoSuchElementException:
                              print('Selenium error occured!')
                              continue
                        except ConnectionAbortedError:
                              print('Connection error occured!')
                              continue
                        except TimeoutException:
                              print('Timeout error occured!')
                              continue
##                        time.sleep(0.5)
                  driver.quit()
##                  time.sleep(0.5)

            with open(os.path.join(PATH_TO, f'cut_{counter}.txt'), 'w', encoding='UTF-8') as f:
                  for i in text_list:
                        f.write(i + '\n')
            counter += 1
