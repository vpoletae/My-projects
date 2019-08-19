from BIS_API import Sanctions
from Spark_API import Spark_api
import Translator
import time
import re
from pprint import pprint
import openpyxl as opxl
      
def read_config(file):
      with open(file, 'r', encoding='UTF-8') as f:
            login = f.readline().strip().replace(' ','').split('=')[1]
            password = f.readline().strip().replace(' ','').split('=')[1]
            wsdl = f.readline().strip().replace(' ','').split('=')[1]
            key = f.readline().strip().replace(' ','').split('=')[1]
      return login, password, wsdl, key

if __name__=='__main__':
      start_time = time.time()
      login, password, wsdl, key = read_config('config.txt')

      api = Spark_api(login=login, password=password, wsdl=wsdl)

      soap = api.initiate_soap()
      api.close_session(soap)
      api.open_session(soap)

      ## think about capturing data from mail

      # 1. Import xlsx data base (export hold report)
      wb = opxl.load_workbook('INN.xlsx')
      ws = wb.active
      col_C = ws['C']
      # 2. Translate from English to Russian
      
      # Repeat for every hierarchy level
            # 3. Use Spark
            
            # 4. Translate from Russian to English
            
            # 5. Use BIS

      api.close_session(soap)

      print("--- %s seconds ---" % (time.time() - start_time))






























































































      
      
