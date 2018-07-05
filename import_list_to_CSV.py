import get_list_from_excel as xl
import csv

class my_dialect(csv.Dialect):
    delimiter = ';'
    quoting = csv.QUOTE_NONE
    quotechar = ','
    lineterminator = '\n'
    escapechar = '\\'

class ToCSV:
    '''
    генерация CSV  на основе XLSX по заданному в битриксе шаблону
    '''
    def __init__(self, xlsx_filename, csv_filename, start_position=14):
        self.__xlsx_filename = xlsx_filename
        self.__csv_filename = csv_filename
        self.__start_position = start_position
        self.counter_records = 0
        self.counter_headers = 0

    def run(self):
        x = xl.XlsxItems(self.__xlsx_filename, self.__start_position)
        ll = x.get_items()
        with open(self.__csv_filename, 'w', newline='') as csvfile:
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

            dialect = my_dialect
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, dialect=dialect)
            writer.writeheader()
            for row in ll:
                if isinstance(row, xl.Item):
                    try:
                        writer.writerow({'IE_XML_ID': abs(hash(row.name)),
                                     'IE_NAME': row.name,
                                     'IE_PREVIEW_TEXT' : row.name,
                                     'IP_PROP44': row.code,
                                     'CP_QUANTITY': row.left})
                        print('name: {}  code: {} art: {} left: {}'.format(row.name, row.code, row.art, row.left))
                        self.counter_records+=1
                    except UnicodeEncodeError:
                        print('Unicode error')
                elif isinstance(row, xl.Header):
                    self.counter_headers+=1


if __name__ == '__main__':
    mycsv = ToCSV('prise list 07,06.2018.xlsx', 'price.csv')
    mycsv.run()
    print('Total records converted: {}'.format(mycsv.counter_records))
    print('Total headers found: {}'.format(mycsv.counter_headers))


