import requests
import json

class Sanctions(object):

      def __init__(self, key):
            """init new object with a key"""
            self.api_key = key

      def make_url(self):
            """concatinating http + api_key"""
            return 'https://api.trade.gov/consolidated_screening_list/search?api_key={0}'.format(self.api_key)

      def make_request(self, params={}):
            """creating a request"""
            url = self.make_url() 
            response = requests.get(url, params=params)
            output = json.loads(response.content.decode("utf-8"))
            return output

      def search_by_fuzzy_name(self, name, fuzzy_name=True):
            """
            star => stir  (substitution)
            star => stars (insertion)
            star => tar   (deletion)
            star => tsar  (transposition)
            """
            response = self.make_request(params={'name':name, 'fuzzy_name':fuzzy_name})
            results_list = response['results']

            request_list = []
            precision = ''
            name = ''
            alt_names = ''
            license_policy = ''
            license_requirement = ''
            programs = ''
            remarks = ''
            source = ''
            source_url = ''

            for address in results_list:
                  try:
                        precision = address['adjusted_score']
                  except KeyError:
                        pass
                  try:
                        name = address['name']
                  except KeyError:
                        pass
                  try:
                        alt_names = address['alt_names']
                  except KeyError:
                        pass
                  try:
                        license_policy = address['license_policy']
                  except KeyError:
                        pass
                  try:
                        license_requirement = address['license_requirement']
                  except KeyError:
                        pass
                  try:
                        programs = address['programs']
                  except KeyError:
                        pass
                  try:
                        remarks = address['remarks']
                  except KeyError:
                        pass
                  try:
                        source = address['source']
                  except KeyError:
                        pass
                  try:
                        source_url = address['source_information_url']
                  except KeyError:
                        pass
                  match_tuple = (precision, name, alt_names, license_policy, license_requirement, programs, remarks, source, source_url)
                  request_list.append(match_tuple)
                  
            return request_list





            

