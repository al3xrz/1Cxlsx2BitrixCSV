import csv
from colorama import Fore

'''
тестер
'''


file_from_site = 'export_file_959921.csv'
file_from_xlsx = 'price.csv'

class my_dialect(csv.Dialect):
    delimiter = ';'
    quoting = csv.QUOTE_NONE
    quotechar = ','
    lineterminator = '\n'
    escapechar = '\\'

fieldnames = ['IE_XML_ID',
                  'IE_NAME',
                  'IE_PREVIEW_TEXT',
                  'IE_DETAIL_TEXT',
                  'IP_PROP44',
                  'IP_PROP45',
                  'IP_PROP46',
                  'IC_GROUP0',
                  'IC_GROUP1',
                  'IC_GROUP2',
                  'CP_QUANTITY',
                  'CP_WEIGHT',
                  'CP_WIDTH',
                  'CP_HEIGHT',
                  'CP_LENGTH',
                  'CV_QUANTITY_FROM',
                  'CV_QUANTITY_TO',
                  'CV_PRICE_1',
                  'CV_CURRENCY_1']


list_from_site = []
list_from_xlsx = []
total_in_site = 0
tottal_in_xlsx = 0


dialect = my_dialect()

with open(file_from_site, encoding='utf-8', newline='') as csvfile:
    r = csv.DictReader(csvfile, dialect=dialect)
    list_from_site = list(r)
    total_in_site = len(list_from_site)
    '''
    for row in list_from_site:
        print(row['IE_NAME'])
    '''


with open(file_from_xlsx,  newline='') as csvfile:
    r = csv.DictReader(csvfile, dialect=dialect)
    list_from_xlsx = list(r)
    tottal_in_xlsx = len(list_from_xlsx)
    '''
    for row in list_from_xlsx:
        print(row['IE_NAME'])
    '''

counter = 0
for row in list_from_site:
    for internal_row in list_from_xlsx:
        if row['IP_PROP44'] == internal_row['IP_PROP44'] :
            modificator = Fore.RESET
            if row['IE_NAME']!=internal_row['IE_NAME']:
                modificator = Fore.RED
            print(modificator+'Name in site: {} Name in XLSX: {} Code in site: {} Code in file: {} Counter {}'.format(
                row['IE_NAME'],
                internal_row['IE_NAME'],
                row['IP_PROP44'],
                internal_row['IP_PROP44'],
                counter))
            counter += 1
    #print(row)
    #print('name: {} code_in_site: {} code_in_xlsx:{}'.format(row['IE_NAME'], row['IP_PROP44'], 5))

print(Fore.GREEN+'Total records in file from site : {}'.format(total_in_site))
print('Total records in file from XLSX : {}'.format(tottal_in_xlsx))
print(Fore.RESET)
