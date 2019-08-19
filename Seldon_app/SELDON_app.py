from pprint import pprint
import traceback

from get_dates import get_dates
from convert_to_xlsx import convert_to_xlsx, \
     write_to_xlsx, get_files_in_folder
from mailer import send_mail, mailer_split
from seldon_api import Seldon_api

import math
import time
import os

PATH = 'your_path'

TIME_SLEEP = 20

def main(path=PATH):
      date_from, date_to = get_dates()
      seldon_api = Seldon_api()
      ##token_parameter = seldon_api.user_login()
      ##print(token_parameter)
      token_parameter = 'Your_token'
##      seldon_api.check_balance(token_parameter)
      filters_list = seldon_api.get_available_filters(token_parameter)

      pages_threshold = 100 # det by Seldon
      for i in filters_list:
            filter_ = i[1]
            filter_name = i[0]
            print(f'Current filter: {filter_name}\n')
            try:
                  task_id_new = seldon_api.make_query(token_parameter, filter_,
                                                      date_from, date_to)
                  print(f'New Task ID: {task_id_new}\n')
            except TypeError as e:
                  print('Ошибка:\n', traceback.format_exc())
            task_id_update = seldon_api.make_update(token_parameter, filter_,
                                                    date_from, date_to)
            print(f'Updated Task ID: {task_id_update}\n')
            
            for i in range(TIME_SLEEP):
                  print('......your request is being processed. Please wait......')
                  time.sleep(1) # sleep time
            
            status_new = seldon_api.get_query_status(token_parameter, task_id_new)[1]
            print(f'New tenders found: {status_new}\n')
            status_update = seldon_api.get_query_status(token_parameter, \
                                                        task_id_update)[1]
            print(f'Tenders updated: {status_update}\n')

            file_name = None
            if status_update:
                  pages_update = math.ceil(status_update/pages_threshold)
                  if pages_update > 1:
                        for page in range(1, pages_update+1):
                              if not file_name:
##                                    try:
                                    purchases_update = seldon_api.request_purchases(
                                          token_parameter, task_id_update, page)
                                    tenders_list = seldon_api.prettify_data(purchases_update)
                                    file_name = convert_to_xlsx(tenders_list)
                                    print(f'File {file_name} was created')
##                                    except TypeError as e:
##                                          print('Ошибка:\n', traceback.format_exc())
                              else:
##                                    try:
                                    purchases_update = seldon_api.request_purchases(
                                          token_parameter, task_id_update, page)
                                    tenders_list = seldon_api.prettify_data(purchases_update)
                                    write_to_xlsx(file_name, tenders_list)
                                    print(f'File {file_name} was updated')
##                                    except TypeError as e:
##                                          print('Ошибка:\n', traceback.format_exc())
                  else:
##                        try:
                        purchases_update = seldon_api.request_purchases(
                              token_parameter, task_id_update)
                        tenders_list = seldon_api.prettify_data(purchases_update)
                        file_name = convert_to_xlsx(tenders_list)
                        print(f'File {file_name} was created')
##                        except TypeError as e:
##                              print('Ошибка:\n', traceback.format_exc())
            print(f'Updated orders for filter {filter_} were reflected\n')
            
            if status_new:
                  pages_new = math.ceil(status_new/pages_threshold)
                  if pages_new > 1:
                        for page in range(1, pages_new+1):
                              if not file_name:
##                                    try:
                                    purchases_new = seldon_api.request_purchases(
                                          token_parameter, task_id_new, page)
                                    tenders_list = seldon_api.prettify_data(purchases_new)
                                    file_name = convert_to_xlsx(tenders_list)
                                    print(f'File {file_name} was created')
##                                    except TypeError as e:
##                                          print('Ошибка:\n', traceback.format_exc())
                              else:
##                                    try:
                                    purchases_new = seldon_api.request_purchases(
                                          token_parameter, task_id_new, page)
                                    tenders_list = seldon_api.prettify_data(purchases_new)
                                    write_to_xlsx(file_name, tenders_list)
                                    print(f'File {file_name} was updated')
##                                    except TypeError as e:
##                                          print('Ошибка:\n', traceback.format_exc())
                  else:
                        if not file_name:
##                              try:
                              purchases_new = seldon_api.request_purchases(
                                    token_parameter, task_id_new)
                              tenders_list = seldon_api.prettify_data(purchases_new)
                              file_name = convert_to_xlsx(tenders_list)
                              print(f'File {file_name} was created')
##                              except TypeError as e:
##                                    print('Ошибка:\n', traceback.format_exc())
                        else:
##                              try:
                              purchases_new = seldon_api.request_purchases(
                                    token_parameter, task_id_new)
                              tenders_list = seldon_api.prettify_data(purchases_new)
                              write_to_xlsx(file_name, tenders_list)
                              print(f'File {file_name} was updated')
##                              except TypeError as e:
##                                    print('Ошибка:\n', traceback.format_exc())
                              
            print(f'New orders for filter {filter_} were reflected\n')
            print('='*70)
            
      files_in_data = get_files_in_folder()
      mailer_package= mailer_split(files_in_data)
      for package in mailer_package:
            send_mail(package[0], package[1], package[2])
            print('File was successfully sent!\n')
      #SMTPSenderRefused
      print('='*70)
      
      for file in files_in_data:
            os.remove(os.path.join(path,file))
      print('All files were successfully removed!')
      
if __name__=='__main__':
      main()




































































      
      
