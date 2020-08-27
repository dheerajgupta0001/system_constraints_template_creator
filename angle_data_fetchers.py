#%%
from docx.shared import Pt
from docxtpl import DocxTemplate, InlineImage
import cx_Oracle
import argparse
import pandas as pd
from datetime import datetime

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
#%%
#print('Range of Data {0} {1}'.format(sDate, eDate))
print("sDate is: {0}".format(sDate))
print(type(sDate))

#print(type(start_date))
#print(end_date)

#%%
# connection creation
try:
    con= cx_Oracle.connect("system/torreto@localhost")
    cursor=con.cursor()
    print(con.version)
    
    print("wide angle report")
    sql_fetch =""" SELECT * FROM (select distinct wide_angle_pair, angular_limit, violation \
        from angle_data where type = 'wide' and date_time BETWEEN TO_DATE(:col1, 'YYYY-MM-DD')\
                and TO_DATE(:col2, 'YYYY-MM-DD')) A\
        left JOIN (select MAX(max_degree), min(min_degree), wide_angle_pair from angle_data\
            where type= 'wide' and date_time BETWEEN TO_DATE(:col1, 'YYYY-MM-DD')\
                and TO_DATE(:col2, 'YYYY-MM-DD') GROUP BY wide_angle_pair) B\
                    on A.wide_angle_pair = B.wide_angle_pair"""

    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD' ")
    cursor.execute(sql_fetch, {'col1':sDate, 'col2':eDate})
    row = cursor.fetchall()
    #print(row)  #records are fetched as a list of touples

    for index, record in enumerate(row):
        print(index,record)
    
    print("Adjacent angle report")
    sql_fetch =""" SELECT * FROM (select distinct wide_angle_pair, angular_limit, violation \
        from angle_data where type = 'adj' and date_time BETWEEN TO_DATE(:col1, 'YYYY-MM-DD')\
                and TO_DATE(:col2, 'YYYY-MM-DD')) A\
        left JOIN (select MAX(max_degree), min(min_degree), wide_angle_pair from angle_data\
            where type= 'adj' and date_time BETWEEN TO_DATE(:col1, 'YYYY-MM-DD')\
                and TO_DATE(:col2, 'YYYY-MM-DD') GROUP BY wide_angle_pair) B\
                    on A.wide_angle_pair = B.wide_angle_pair"""

    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY-MM-DD' ")
    cursor.execute(sql_fetch, {'col1':sDate, 'col2':eDate})
    row = cursor.fetchall()
    #print(row)  #records are fetched as a list of touples

    for index, record in enumerate(row):
        print(index,record)
    print("QUERY FETCHED")

except:
    print('Error while fetching data from db')
finally:
    # closing database cursor and connection
    if cursor is not None:
        cursor.close()
    con.close()