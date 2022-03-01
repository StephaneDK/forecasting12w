# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 15:47:43 2021

@author: Becky
"""

from __future__ import print_function
import os
import sys


os.chdir('C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv')
#os.chdir('C:\\dev\\env\\dkwebsols-forecasting-12w-22ef3797943e')

country_var = sys.argv[1]


class LatestDataCheck(Exception):
    pass




#Connecting to DB --------------------------------------------------------------------------------------------------------------------
#!/usr/bin/python
import snowflake.connector as sfc
from snowflake.connector.pandas_tools import write_pandas
import pandas as pd
import psycopg2
from configparser import ConfigParser

def config(filename='database.ini', section='snowflake_db'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)

    # get section, default to snowflake_db
    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))

    return db


def connect():
    """ Connect to the GDH database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the GDH server
        print('\nConnecting to the GDH database...')
        conn = sfc.connect(account='PRH',
                                region='us-east-1',
                                user= params['user'],
                                password= params['password'],
                                database="PRH_GLOBAL_DK_SANDBOX",
                                schema="PUBLIC")
		
        # create a cursor
        cur = conn.cursor()
        
        #execute a statement
        print('Snowflake database version:')
        cur.execute('SELECT CURRENT_VERSION()')

        # display the GDH database server version
        db_version = cur.fetchone()
        print(db_version, ' \n\n ')
        
        #Copying csv file to forecasting_temp
        file_fc = pd.read_csv("Write_data_sources/Q1 Forecast {0} - Formatted.csv".format(country_var))
        write_pandas(conn, file_fc, "FORECASTING_TEMP")
        
        
        #Copying csv file to forecasting_statistics_temp
        file_fcs = pd.read_csv("Write_data_sources/Q1 Forecast {0}.csv".format(country_var))
        file_fcs.columns = file_fcs.columns.str.upper()
        write_pandas(conn, file_fcs, "FORECASTING_STATISTICS_TEMP")
       
       
        
        print("Forecasts uploaded to database\n")
        


       
        #close the communication with the GDH
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed \n')

if __name__ == '__main__':
    connect()









