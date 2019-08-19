from zeep import Client
from zeep import helpers
import re

class Spark_api(object):
      def __init__(self, login, password, wsdl):
            self.login = login
            self.password = password
            self.wsdl = wsdl

      def initiate_soap(self):
            client = Client(wsdl=self.wsdl)
            return client.service

      def open_session(self, soap):
            soap.Authmethod(self.login, self.password)
            
      def close_session(self, soap):
            soap.End()

      def identify_company(self, soap, name, address):
            """
            defines company tax id
            returns TaxID, ShortName, Address, Probability
            returns list of tuples
            """
            response = soap.IdentifyCompany(companyName=name, address=address, fuzzySearch=1,
                           historySearch=1, activeCompanySearch=1, excludeBranch=1)
            response = response['xmlData']
            inn_list = []
            name_list = []
            address_list = []
            probability_list = []
            
            try:
                  inn_response = re.findall(r'<INN>\d+</INN>', response)
                  name_response = re.findall(r'<ShortName>.+?</ShortName>', response)
                  address_response = re.findall(r'Address=".+?REGION=', response)
                  probability_response = re.findall(r'<Probability>\d+</Probability>', response)
            except IndexError:
                  pass

            for inn in inn_response:
                  inn_list.append(int(inn.replace('<INN>','').replace('</INN>','')))
            for name in name_response:
                  name_list.append(name.replace('<ShortName>','').replace('</ShortName>',''))
            for address in address_response:
                  address_list.append(address.replace('Address="','').replace('" REGION=',''))
            for probability in probability_response:
                  probability_list.append(probability.replace('<Probability>','').replace('</Probability>',''))

            company_data = list(zip(inn_list, name_list, address_list, probability_list))
            
            return company_data
      
      def get_comp_structure(self, soap, tax_id):
            """
            captures actual company info
            returns  comp name, INN, share
            returns dict
            """
            response = soap.GetCompanyStructure(inn=tax_id)
            response = response['xmlData']
            #---------------------------------------------------------------------------------------------------------------------
            # collect overall company info
            #---------------------------------------------------------------------------------------------------------------------
            def get_coowners(response):
                  response_company = ''
                  response_rosstat = ''
                  response_egrul = ''
                  try:
                        response_company = re.findall(r'<CoownersFCSM[\w\W]+<\/CoownersFCSM>', \
                                                      response)
                        response_rosstat = re.findall(r'<CoownersRosstat[\w\W]+<\/CoownersRosstat>', \
                                                      response)
                        response_egrul = re.findall(r'<CoownersEGRUL[\w\W]+<\/CoownersEGRUL>', \
                                                    response)
                  except IndexError:
                        pass
                  return response_company, response_rosstat, response_egrul
            #---------------------------------------------------------------------------------------------------------------------
            # grab actual update data
            #---------------------------------------------------------------------------------------------------------------------
            def get_actual_dates(response_company, response_rosstat, response_egrul):
                  actual_date_company = int()
                  actual_date_rosstat = int()
                  actual_date_egrul = int()
                  if response_company and len(response_company) != 0:
                        actual_date_company = int(re.findall(r'ActualDate="\d+.?\d+.?\d+', \
                                                       response_company[0])[0].replace('ActualDate="','').replace('-',''))
                  if response_rosstat and len(response_rosstat) != 0:
                        actual_date_rosstat = int(re.findall(r'ActualDate="\d+.?\d+.?\d+', \
                                                       response_rosstat[0])[0].replace('ActualDate="','').replace('-',''))
                  if response_egrul and len(response_egrul) != 0:
                        actual_date_egrul = int(re.findall(r'ActualDate="\d+.?\d+.?\d+', \
                                                     response_egrul[0])[0].replace('ActualDate="','').replace('-',''))
                  return actual_date_company, actual_date_rosstat, actual_date_egrul
            #---------------------------------------------------------------------------------------------------------------------
            # discover the latest update dates
            #---------------------------------------------------------------------------------------------------------------------
            def get_latest_update(actual_date_company, actual_date_rosstat, actual_date_egrul):
                  the_latest_update = ''
                  if actual_date_company >= actual_date_rosstat:
                        if actual_date_company >= actual_date_egrul:
                              the_latest_update = actual_date_company
                        else:
                              the_latest_update = actual_date_egrul
                  else:
                        if actual_date_rosstat >= actual_date_egrul:
                              the_latest_update = actual_date_rosstat
                        else:
                              the_latest_update = actual_date_egrul
                  return the_latest_update
            #---------------------------------------------------------------------------------------------------------------------
            # get list of owners
            #---------------------------------------------------------------------------------------------------------------------
            def get_list_owners(the_latest_update, \
                                actual_date_company, actual_date_rosstat, actual_date_egrul):
                  if the_latest_update == actual_date_company:
                        owners_list = re.findall(r'<CoownerFCSM>.+?<\/CoownerFCSM>', \
                                                 response_company[0])
                  elif the_latest_update == actual_date_rosstat:
                        owners_list = re.findall(r'<CoownerRosstat>.+?<\/CoownerRosstat>', \
                                                 response_rosstat[0])
                  else:
                        owners_list = re.findall(r'CoownerEGRUL>.+?<\/CoownerEGRUL>', \
                                                 response_egrul[0])
                  return owners_list
            #---------------------------------------------------------------------------------------------------------------------
            # get detailed info about company
            #---------------------------------------------------------------------------------------------------------------------
            def get_detailed_info(owners_list):
                  name = ''
                  inn = ''
                  share = ''
                  owner_company = dict() 
                  for owner in owners_list:
                        try:
                              share = re.findall(r'<SharePart>.+<\/SharePart>', owner)[0]
                              share = share.replace('<SharePart>','').replace('</SharePart>','')
                              if ',' in share:
                                    share = share.replace(',', '.')
                              else:
                                    pass
                        except IndexError:
                              pass
                        try:
                              name = re.findall(r'<Name>.+<\/Name>', owner)[0]
                              name = name.replace('<Name>','').replace('</Name>','')
                        except IndexError:
                              pass
                        try:
                              inn = re.findall(r'<INN>\d+<\/INN>', owner)[0]
                              inn = inn.replace('<INN>','').replace('</INN>','')
                        except IndexError:
                              pass

                        owner_company[name] = (inn, share)
                  return owner_company
            #---------------------------------------------------------------------------------------------------------------------
            # base loop
            #---------------------------------------------------------------------------------------------------------------------
            response_company, response_rosstat, response_egrul = get_coowners(response)
            actual_date_company, actual_date_rosstat, actual_date_egrul = get_actual_dates(
                  response_company, response_rosstat, response_egrul)
            owner_company = dict()
            if actual_date_company or actual_date_rosstat or actual_date_egrul:
                  the_latest_update = get_latest_update(
                        actual_date_company, actual_date_rosstat, actual_date_egrul)
                  owners_list = get_list_owners(
                        the_latest_update, actual_date_company, actual_date_rosstat, actual_date_egrul)
                  owner_company = get_detailed_info(owners_list)
            return owner_company



















































































      
      
