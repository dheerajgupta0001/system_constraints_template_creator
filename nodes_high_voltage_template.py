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

doc = DocxTemplate(tmplPath)
try:
    connection= cx_Oracle.connect(appDbConnStr)
    cursor=connection.cursor()
    print(connection.version)
    
    sql_fetch ="""select * from nodes_high_voltage_data \
        where start_date IN (select start_date as s from nodes_high_voltage_data where rownum<2)"""
    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD' ")
    df= pd.read_sql(sql_fetch, con=connection)
    
    #print(df)
    #print(type(df))

except:
    print('Error while fetching data from db')
finally:
    # closing database cursor and connection
    if cursor is not None:
        cursor.close()
    connection.close()
highVoltageNodeData = []
for i in df.index:
    tempDict={
        'nodes': df['NODES'][i],
        'season': df['SEASON_ANTECEDENT'][i],
        'description': df['DESCRIPTION_CONSTRAINTS'][i]
    }
    highVoltageNodeData.append(tempDict)

print(len(highVoltageNodeData))