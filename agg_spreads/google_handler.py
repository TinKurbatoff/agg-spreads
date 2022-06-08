# import libraries

import sys
import os
import logging
import argparse
import pandas as pd
import json
import csv
import itertools
# import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# #############################################
# ############ ENABLE LOGGING #################
# #############################################
# create logger
logger = logging.getLogger(__name__)
logging.basicConfig(filename=f'{__file__}.log', 
                    level=logging.INFO, 
                    format='%(asctime)s: /%(name)s/ %(levelname)s: %(message)s', 
                    datefmt='%m/%d/%Y %I:%M:%S %p')
logger.setLevel(logging.DEBUG)  # ALWAYS DEBUG in FILE
# create console handler and set level to debug
lg = logging.StreamHandler()
lg.setLevel(logging.INFO)
# create formatter
formatter = logging.Formatter(fmt='%(asctime)s: /%(name)s/ %(levelname)s: %(message)s', 
                              datefmt='%m/%d/%Y %I:%M:%S %p')
# add formatter to secondary logger
lg.setFormatter(formatter)
# add ch to logger
logger.addHandler(lg)  # when call logger both will destinations will be filled
#######################################


# ############### ———————————————————————————————————— #########################
# '''        CLASSES        '''
# #########################################################################
# ####### ———————- GOOGLE SHEETS HANDLING CLASS -——————— ##################
# #########################################################################
class GoogleSheetsObjects(object):
    """ This is a service object (used by GoogleSheetsHandler wrapper),
    opens Google Sheet object and operatest with it

    How to call:
    GS = GoogleSheets('GoogleSheetAPIKey.json')
    GS_file =GS.openWorksheet(file=ID='<a long hex Google Sheets ID>', page = 0) # a tab nposition in the file, 0 - the very first
    GS_tab = GS.workingSheet # a tab in the file (to access) 
    """

    def __init__(self, keyfile):
        logger.info("Initializing API...")
        from string import ascii_uppercase 
        # Create a scope of rights
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        # create some credential using that scope and content of startup_funding.json
        self.keyfile = os.path.join("./", keyfile)
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(self.keyfile, self.scope)
        # create gspread authorize using that credential
        self.client = gspread.authorize(self.creds)
        self.wks = None  # Objct with Google sheet file
        self.workingSheet = None  # a working sheet
        self.column_template = ascii_uppercase

    def batch_update(self, update):
        # print('!')
        result = self.workingSheet.batch_update(update)
        return result 

    def openWorksheet(self, fileID='', page=0, tab_name=None):
        ''' Open Google sheet by ID key '''
        logger.info("Opening google sheet...")
        # Now will can access our google sheets we call client.open on StartupName
        try:
            self.wks = self.client.open_by_key(fileID)  # Open by Sheet ID
            # self.wks = self.client.open(filename) # Open by Sheet filename
        except gspread.exceptions.APIError as e:
            error_message = json.loads(str(e))
            logger.critical("Error {}: {}".format(error_message['error']['code'], error_message['error']['message']))
            return None
        logger.debug("Opening sheet...")
        try:
            if tab_name:
                self.workingSheet = self.wks.worksheet(tab_name)
            else:
                self.workingSheet = self.wks.get_worksheet(page)  # Select a working sheet in the file
        except gspread.exceptions.APIError as e:
            error_message = json.loads(str(e))
            logger.critical("Error {}: {}".format(error_message['error']['code'], error_message['error']['message']))
            if error_message['error']['code'] == 403:
                with open(self.keyfile, 'r') as f:
                    key_data = json.load(f)
                    logger.critical('Add this email as editor to google doc: {}'.format(key_data['client_email']))
            return None

        logger.debug("--------")
        logger.info("Opened doc: {}".format(self.wks.title))
        # logger.info("Last modified: {}".format(self.wks.lastUpdated))
        logger.debug("Selected sheet: {}".format(self.workingSheet.title))
        logger.debug("Number of rows: {}".format(self.workingSheet.row_count))
        logger.debug("--------")            
        return self.workingSheet

    def clearRange(self, google_sheet_pointer, line1=0, line2=0, column1='A', column2='A'):
        result = None
        try:
            # range_pointer = google_sheet_pointer.range(range_to_delete)
            # range_2_delete = google_sheet_pointer.range(range_to_delete)
            sheet_title = google_sheet_pointer.title
            range_to_delete = "'{}'!{}{}:{}{}".format(sheet_title, column1, line1, column2, line2)  # read columns A to N from second row 
            # logger.info(range_to_delete)
            result = self.wks.values_clear(range_to_delete)
        except Exception as e:
            logger.error(e)
        return result

    def readRange(self, google_sheet_pointer, line1=0, line2=0, column1='A', column2='N'): 
        """ READ FROM LINE1 TO LINE2 column A-N """
        # result = sheet.row_values(5) #See individual row
        if line1 == 0:
            line1 = 1  # if not set read from the begging
        if line2 == 0:  
            lines = google_sheet_pointer.row_count  # read the maximum lines
        range_to_read = '{}{}:{}{}'.format(column1, line1, column2, line2)  # read columns A to N from second row 
        try:  # Reading all table data at once
            result = google_sheet_pointer.range(range_to_read)
            # result = self.wks.sheet1.range('A1:N1')
        except gspread.exceptions.APIError as e:
            error_message = json.loads(str(e))
            logger.critical("Error {}: {}".format(error_message['error']['code'], error_message['error']['message']))
            return "Error {}: {}".format(error_message['error']['code'], error_message['error']['message'])
        return result

    def updateRangeColor(self, google_sheet_pointer, 
                         column1='B', line1=2,  
                         column2='B', line2=2, 
                         red=0.2, green=0.8, blue=0.2):
        coloring_range = '{}{}:{}{}'.format(column1, line1, column2, line2)
        result = google_sheet_pointer.format(coloring_range, {"backgroundColor": 
                                                               {"red": red,  # noqa: E127
                                                                "green": green, 
                                                                "blue": blue}})

        return result

    def saveWorksheetToCSV(self, google_sheet_pointer, filename='googleSheet'):
        result = {}
        result = {'error': False, 'status': ''}
        try:
            with open(filename + '.csv', 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(google_sheet_pointer.get_all_values())
            result['status'] = filename+'.csv'  # if ok, return actual filename
        except Exception as e:
            result['error'] = True
            result['status'] = e
        return result

    def add_worksheet(self, title, rows, cols):
        result = self.wks.add_worksheet(title=title, rows=rows, cols=cols)
        return result

    def update_range_by_corner(self, google_sheet_pointer, corner='A1', data=[[]]):
        if len(data[0]) == 0: 
            return  # empty request
        # update_array = np.array(data)
        # result = google_sheet_pointer.update(corner, update_array.tolist())
        result = google_sheet_pointer.update(corner, data, raw=False)
        return result


# #############################################################        
class GoogleSheet(object):
    """This is a wrapper  for GoogleSheets Class for convenient operations with google sheets
        Call example:
        GoogleSheetTable = GoogleSheetObject(sheetID='<a long hex Google Sheet ID>', 
                                             keyfile='<API key file.json>'  # for example: `PayhamsterT-8d3756206202.json`
        print(GoogleSheetFile.select_tab('MyTab')) # print name of the active tab

        NOTE! 
        Do not forget to share with google key service account: example: `getgooglesheets@payhamstert.iam.gserviceaccount.com`
 
    """
    def __init__(self, sheetID, keyfile, tab_name=None):
        self.sheetID = sheetID
        # ——— Open Google sheet
        self.data_source = GoogleSheetsObjects(keyfile=keyfile)  # Open Google Sheet with using key file 
        logger.info(f'Open sheet ID: `{self.sheetID}`')
        if tab_name:
            self.active_sheet = self.data_source.openWorksheet(fileID=self.sheetID, tab_name=tab_name)  # open sheet in file
        else:
            self.active_sheet = self.data_source.openWorksheet(fileID=self.sheetID, page=0)  # open sheet in file
        self.sheetTitle = self.active_sheet.title
        self.file_handler = self.data_source.wks
        self.sheetsList = [x.title for x in self.file_handler.worksheets()]
        logger.info(f'Open sheet: {self.sheetTitle}')
        # DEBUG ** update_result = self.data_source.update_range_by_corner(self.active_sheet, corner='A1', data=[['-'.join(self.sheetTitle.split(' '))]])

    def get_sheet_by_name(self, name):
        self.active_sheet = self.data_source.wks.worksheet(name)
        self.sheetTitle = self.active_sheet.title
        return self.sheetTitle    

    def select_tab(self, tab_name):
        self.get_sheet_by_name(name=tab_name)
        logger.info(f'Selected other sheet {self.active_sheet.title}')
        return

    def update_sheets_list(self):
        self.sheetsList = [x.title for x in self.file_handler.worksheets()]
        return self.sheetsList    

    def add_worksheet(self, title, rows=100, cols=255):
        result = self.data_source.add_worksheet(title=title, rows=f'{rows}', cols=f'{cols}')
        self.update_sheets_list()
        return result

    def rename_sheet(self, newtitle):
        self.sheetTitle = self.active_sheet.update_title(newtitle)
        self.sheetTitle = self.active_sheet.title
        # update_result = self.data_source.update_range_by_corner(self.active_sheet, corner='A1', data=[['-'.join(self.sheetTitle.split(' '))]])
        return self.sheetTitle  

    def duplicate_sheet(self, title, new_title, insert_index=None):
        """ duplicate sheet that titled by title """
        if new_title in self.sheetsList:
            self.get_sheet_by_name(new_title)
            return 'Duplicate'
        if not insert_index:  # if not index specified add to the end
            insert_index = len(self.sheetsList) 
        try:
            self.get_sheet_by_name(title)
            self.file_handler.duplicate_sheet(self.active_sheet.id, insert_sheet_index=insert_index, new_sheet_id=None, new_sheet_name=new_title)
            self.update_sheets_list()
            self.get_sheet_by_name(new_title)
        except Exception as e:
            return f'{e}'
        return 'Success'

    def read_sheet_to_dataframe(self, corner=None, width=None, heigh=None, range=None):
        # list_of_lists = self.active_sheet.get_all_values()
        dataframe = pd.DataFrame(self.active_sheet.get_all_values())
        # print(dataframe.head(n=10)) # DEBUG * DEBUG * DEBUG
        return dataframe

    def read_sheet_to_list(self, corner=None, width=None, heigh=None, range=None):
        list_of_lists = self.active_sheet.get_all_values()
        # print(dataframe.head(n=10)) # DEBUG * DEBUG * DEBUG
        return list_of_lists

    def read_sheet_to_dict(self, corner=None, width=None, heigh=None, range=None):
            dictionary = {}
            list_of_lists = self.active_sheet.get_all_values()
            # Parse lists to dict
            try:
                keys = list_of_lists[0]
                values = list(map(list, itertools.zip_longest(*list_of_lists[1:], fillvalue=None)))  # Transpose rows to columns
                for key, column in zip(keys, values):
                    dictionary[key] = column
            except Exception as e:
                logger.error(f"{e}")
            return dictionary

    def update_range_by_corner(self, corner='A1', data=[['OPENED']]):
        """ Updates data by corner A1 notation, list of lists, inner list is a row """
        # print('data to send: ',data) 
        update_result = 'Fail'
        update_result = self.data_source.update_range_by_corner(self.active_sheet, corner=corner, data=data)
        return update_result

    def updateRangeColor(self, column1='B', line1=2, column2='B', line2=2, red=0.2, green=0.8, blue=0.2):
        """ Function to change background color """
        update_result = self.data_source.updateRangeColor(self.active_sheet, 
                                                          column1=column1, line1=line1,  
                                                          column2=column2, line2=line2, 
                                                          red=red, green=green, blue=blue)
        # update_result — ??
        return update_result


# ################## END OF GOOGLE SHEET HANDLING CLASS #######################
# #############################################################################

# #############################################################################
# ############### ————————————- MAIN -——————————————— #########################
# #############################################################################
def main():
    print("Running in CLI mode is disabled")
    return 

# #############################################################################
# ###################### COMMAND LINE INTERFACE ###############################
# ————————————————————————————————————————————————————————————————————————————#
# #############################################################################
if __name__ == '__main__':
    logger.info('——————————— -  BEGIN  - ————————————')  # will not be logged in API mode
    
    # ——— parse command-line arguments
    parser = argparse.ArgumentParser(description='Google table handler')
    parser.add_argument('-g', '--googel-id', dest='google_id', help='Google ID table')
    parser.add_argument('-d', '--disable-something', dest='disable_', action="store_true", default=False, help='Boolean key')
    args = parser.parse_args()
    
    # —————————— Call function to get info —————————— 
    main()
##############################################################################
