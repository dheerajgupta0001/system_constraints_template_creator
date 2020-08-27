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
    
    sql_fetch ="""SELECT * from ( select start_date, ict, season_antecedent,\
        description_constraints, MAX(START_DATE) over(PARTITION BY id) max_date \
            from ict_constraint_data)WHERE start_DATE = MAX_DATE"""
    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD' ")
    df= pd.read_sql(sql_fetch, con=connection)
    
    #print(df)

except:
    print('Error while fetching data from db')
finally:
    # closing database cursor and connection
    if cursor is not None:
        cursor.close()
    connection.close()
ictConstraintsData = []
for i in df.index:
    tempDict={
        'ict': df['ICT'][i],
        'season': df['SEASON_ANTECEDENT'][i],
        'description': df['DESCRIPTION_CONSTRAINTS'][i]
    }
    ictConstraintsData.append(tempDict)
print(ictConstraintsData)