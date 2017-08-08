import dateutil.parser
import folder_crawler as fc
import pandas as pd
import datetime
import math

############
# the purpose of this file is to detect the datetime format automatically and convert them into the standard python datetime object
############


def datetime_converter(input_date_string):
    print('\ncurrent date time is: ', input_date_string)
    if input_date_string != '':
        if input_date_string == 'nan':
            return 'no date'
        elif type(input_date_string) == type(6.3):
            if math.isnan(input_date_string):
                # print(math.isnan(input_date_string))
                return input_date_string
            else:
                try:
                    date_converted = dateutil.parser.parse(input_date_string)
                    print('converted old date', input_date_string, ' to new date ', date_converted)
                    return date_converted
                except:
                    print('date cannot be converted')
                    return input_date_string
        elif type(input_date_string) != type(datetime.datetime.now()):
            try:
                date_converted = dateutil.parser.parse(input_date_string)
                print('converted old date', input_date_string, ' to new date ', date_converted)
                return date_converted
            except:
                print('date cannot be converted')
                return input_date_string
        else:
            return input_date_string
    else:
        return ""
# print(type(6.3))