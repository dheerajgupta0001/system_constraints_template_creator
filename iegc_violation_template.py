from docx.shared import Pt
from docxtpl import DocxTemplate, InlineImage
import cx_Oracle
import argparse
import pandas as pd
from datetime import datetime
from src.config.appConfig import getConfig

appConfig = getConfig()
print(appConfig)
tmplPath = "assets/docxtpl/template_example.docx"
appDbConnStr = appConfig['appDbConStr']

# get an instance of argument parser from argparse module
parser = argparse.ArgumentParser()

# setup firstname, lastname arguements
parser.add_argument('--startdate', help="Enter start date here")
parser.add_argument('--enddate', help="Enter end date here")

# get the dictionary of command line inputs entered by the user
args = parser.parse_args()

# access each command line input from the dictionary
sDate = args.startdate
eDate = args.enddate

doc = DocxTemplate(tmplPath)
try:
    connection= cx_Oracle.connect(appDbConnStr)
    cursor=connection.cursor()
    #print(connection.version)

    sql_fetch =""" SELECT * FROM IEGC_VIOLATION_MESSAGE_DATA where date_time BETWEEN TO_DATE(:col1, 'YYYY-MM-DD')\
                and TO_DATE(:col2, 'YYYY-MM-DD')"""
    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD' ")
    df= pd.read_sql(sql_fetch, params={'col1': sDate, 'col2': eDate}, con=connection)
    '''cursor.execute(sql_fetch, {'col1':sDate, 'col2':eDate})
    row = cursor.fetchall()'''
    #print(df)
    #print(type(df))

except:
    print('Error while fetching data from db')
finally:
    # closing database cursor and connection
    if cursor is not None:
        cursor.close()
    connection.close()

iegcData = []
for i in df.index:
    tempDict={
        'message': df['MESSAGE'][i],
        'date': df['DATE_TIME'][i],
        'entity': df['ENTITY'][i],
        'schedule': int(round(df['SCHEDULE'][i])),
        'drawal': int(round(df['DRAWAL'][i])),
        'deviation': int(round(df['DEVIATION'][i]))
    }
    iegcData.append(tempDict)
    
context = {
    'yr_str': "2020-21",
    'wk_num': 18,
    'st_dt': sDate,
    'end_dt': eDate,
    'iegc_data': iegcData
}
#print(iegcData)
print(type(iegcData))

doc.render(context)
#doc.render(iegcData)
doc.save("assets/docxtpl/report_created.docx")
