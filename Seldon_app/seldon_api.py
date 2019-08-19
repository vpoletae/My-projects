import requests
import json
from pprint import pprint

LOGIN = "your_login"
PASSWORD = "your_password"

class Seldon_api:
      def __init__(self, login=LOGIN, password=PASSWORD):
            self.login = login
            self.password = password

      def make_url(self, method):
            return f'https://api.seldon2010.ru/{method}'

      def user_login(self, params={}):
            '''
            requests for token
            returns str()
            '''
            url = self.make_url('user/login')
            params["name"] = self.login
            params["password"] = self.password
            response = requests.post(url, json=params)
            token = json.loads(response.content.decode("utf-8"))
            token = token['result']['token']
            token_parameter = f'?token={token}'
            return token_parameter
      
      def check_balance(self, token_parameter):
            '''
            requests for balance
            '''
            url = self.make_url(f'user/balance{token_parameter}')
            
            response = requests.post(url)
            balance = json.loads(response.content.decode("utf-8"))
            print(balance)
      
      def get_available_filters(self, token_parameter):
            '''
            requests for filter_id
            returns list of tuples (0-name, 1-filterId)
            '''
            url = self.make_url(f'user/filters{token_parameter}')
            
            response = requests.post(url)
            filters = json.loads(response.content.decode("utf-8"))
            filters = filters['result']['filters']

            filters_list = []
            for filter_ in filters:
                  filters_list.append((filter_['name'], filter_['filterId']))

            return filters_list

      def make_query(self, token_parameter, filter_id, date_from, date_to, params = {}):
            '''
            requests for task_id_new
            returns str()
            '''
            year_from = date_from[0]
            month_from = date_from[1]
            day_from = date_from[2]
            year_to = date_to[0]
            month_to = date_to[1]
            day_to = date_to[2]
            
            url = self.make_url(f'purchases/new{token_parameter}')
            params['filterId'] = int(filter_id)
            params['dateFrom'] = f"{year_from}-{month_from}-{day_from}T10:00:00.0000000+03:00"
            params['dateTo'] = f"{year_to}-{month_to}-{day_to}T21:00:00.0000000+03:00"
            response = requests.post(url, json=params)
            task_id_new = json.loads(response.content.decode("utf-8"))
            task_id_new = task_id_new['result']['taskId']
            return task_id_new
      
      def make_update(self, token_parameter, filter_id, date_from, date_to, params = {}):
            '''
            requests for task_id_update
            returns str()
            '''
            year_from = date_from[0]
            month_from = date_from[1]
            day_from = date_from[2]
            year_to = date_to[0]
            month_to = date_to[1]
            day_to = date_to[2]
            
            url = self.make_url(f'purchases/update{token_parameter}')
            params['filterId'] = int(filter_id)
            params['dateFrom'] = f"{year_from}-{month_from}-{day_from}T10:00:00.0000000+03:00"
            params['dateTo'] = f"{year_to}-{month_to}-{day_to}T21:00:00.0000000+03:00"
            response = requests.post(url, json=params)
            task_id_update = json.loads(response.content.decode("utf-8"))
            task_id_update = task_id_update['result']['taskId']
            return task_id_update
      
      def get_query_status(self, token_parameter, task_id, params = {}):
            '''
            requests for current task status
            returns number of results found: int
            '''
            url = self.make_url(f'purchases/status{token_parameter}')
            params['taskId'] = task_id
            response = requests.post(url, json=params)
            response = json.loads(response.content.decode("utf-8"))
            status = response['result']['searchStatus']['descr']
            quantity = response['result']['quantity']
            return (status, quantity)
      
      def request_purchases(self, token_parameter, task_id, page_index=1, params = {}):
            '''
            requests for data about actual purchases
            returns list of dictionaries
            '''
            url = self.make_url(f'purchases/result{token_parameter}')
            params['taskId'] = task_id
            params['pageIndex'] = page_index
            response = requests.post(url, json=params)
            purchases = json.loads(response.content.decode("utf-8"))
            purchases = purchases['result']['purchases']
            return purchases

      def prettify_data(self, purchases):
            tenders_list = []

            filter_name = str()
            fact_address = str()

            products_list = [] # of product dicts
            
            for purchase in purchases:
                  tender_dict = dict()
                  tender_dict['contract_type'] = purchase['contractType']['name']
                  tender_dict['subject'] = purchase['subject']
                  tender_dict['start_date'] = purchase['startDate']
                  tender_dict['end_date'] = purchase['endDate']
                  tender_dict['purchase_link'] = purchase['purchaseLink']
                  tender_dict['purchase_price'] = purchase['purchasePrice']
                  tender_dict['current_status'] = purchase['status']['statusSeldon']
                  filter_name = purchase['filterName']
                  tender_dict['filter_name'] = filter_name
                  
                  try:
                        tender_dict['delivery_place'] = purchase['lotsList'][0]['customersList']\
                                                        [0]['deliveryPlace']
                  except (TypeError, IndexError) as e:
                        tender_dict['delivery_place'] = ''

                  try:
                        tender_dict['delivery_term'] = purchase['lotsList'][0]['customersList']\
                                                       [0]['deliveryTerm']
                  except (TypeError, IndexError) as e:
                        tender_dict['delivery_term'] = ''

                  try:
                        tender_dict['contact_person'] = purchase['lotsList'][0]['customersList']\
                                                        [0]['organization']['contactPerson']
                  except (TypeError, IndexError) as e:
                        tender_dict['contact_person'] = ''
                        
                  try:
                        tender_dict['email'] = purchase['lotsList'][0]['customersList'][0]\
                                               ['organization']['email']
                  except (TypeError, IndexError) as e:
                        tender_dict['email'] = ''

                  try:
                         fact_address = purchase['lotsList'][0]['customersList']\
                                                      [0]['organization']['factAddress']
                         tender_dict['fact_address'] = fact_address
                  except (TypeError, IndexError) as e:
                        tender_dict['fact_address'] = ''

                  try:
                        split_address = fact_address.split(',')
                        if '223' in filter_name:
                              clean_address = split_address[1].strip()
                        else:
                              clean_address = split_address[2].strip()
                        tender_dict['clean_address'] = clean_address
                  except (TypeError, IndexError, AttributeError) as e:
                        tender_dict['clean_address'] = ''

                  try:
                        tender_dict['inn'] = purchase['lotsList'][0]['customersList']\
                                             [0]['organization']['inn']
                  except (TypeError, IndexError) as e:
                        tender_dict['inn'] = ''

                  try:
                        tender_dict['org_name'] = purchase['lotsList'][0]['customersList']\
                                                  [0]['organization']['name']
                  except (TypeError, IndexError) as e:
                        tender_dict['org_name'] = ''

                  try:
                        tender_dict['phone'] = purchase['lotsList'][0]['customersList']\
                                               [0]['organization']['phone']
                  except (TypeError, IndexError) as e:
                        tender_dict['phone'] = ''
                        
                  products = purchase['lotsList'][0]['productsList']
                  if products:
                        for product in products:
                              product_dict = dict() # name/price/quantity/value
                              
                              if product['classifier']['okdp']['name']:
                                    product_dict['group_name'] = product['classifier']['okdp']['name']
                              elif product['classifier']['okpd']['name']:
                                    product_dict['group_name'] = product['classifier']['okpd']['name']
                              elif product['classifier']['okpd2']['name']:
                                    product_dict['group_name'] = product['classifier']['okpd2']['name']
                              else:
                                    product_dict['group_name'] = ''

                              if product['name']:
                                    product_dict['name_unit'] = product['name']
                              else:
                                    product_dict['name_unit'] = ''
                              if product['price']:     
                                    product_dict['price_per_unit'] = product['price']
                              else:
                                    product_dict['price_per_unit'] = ''
                              if product['quantity']:
                                    product_dict['number_of_units'] = product['quantity']
                              else:
                                    product_dict['number_of_units'] = ''
                              if product['value']:
                                    product_dict['value'] = product['value']
                              else:
                                    product_dict['value'] = ''
                        
                              products_list.append(product_dict)

                  tender_dict['products_list'] = products_list
                  requirements = purchase['lotsList'][0]['requirementsList']
                  if requirements:
                        tender_dict['requirements'] = requirements[0]
                  else:
                        tender_dict['requirements'] = ''
                  tenders_list.append(tender_dict)

            return tenders_list
































































      
      
