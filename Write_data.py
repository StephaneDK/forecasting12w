# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 15:47:43 2021

@author: Stephane
"""

from __future__ import print_function
import os
import sys


os.chdir('C:\\Users\\Stephane\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv')


country_var = sys.argv[1]

class LatestDataCheck(Exception):
    pass




#Connecting to DB --------------------------------------------------------------------------------------------------------------------
#!/usr/bin/python
import psycopg2
from configparser import ConfigParser
import csv

def config(filename='database.ini', section='postgresql'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)

    # get section, default to postgresql
    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))

    return db


def connect():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        print('\nConnecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)
		
        # create a cursor
        cur = conn.cursor()
        
        #execute a statement
        print('PostgreSQL database version:')
        cur.execute('SELECT version()')

        # display the PostgreSQL database server version
        db_version = cur.fetchone()
        print(db_version, ' \n\n ')
        
        #Copying csv file to forecasting_temp
        file = open("C:\\Users\\Stephane\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv\\Q1 Forecast {0} - Formatted.csv".format(country_var))

        cur.copy_expert("""
        COPY experiment.forecasting_temp(region, model, asin, pred_at, pred_for, pred_units)
        FROM STDIN WITH CSV HEADER DELIMITER as ','
        """, file
        )

        conn.commit()


        #Copying csv file to forecasting_statistics_temp
        file = open("C:\\Users\\Stephane\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv\\Q1 Forecast {0}.csv".format(country_var))

        cur.copy_expert("""
        COPY experiment.forecasting_statistics_temp(region, model, asin, pred_at, "Amz_inv", "DK_inv", "Total_inventory", "Forecast_12w", 
        "Adjusted_forecast_12w", "Weeks_on_hand", "Weeks_on_hand_AMZ", "Inv_issue","reprint_date","reprint_quantity")
        FROM STDIN WITH CSV HEADER DELIMITER as ','
        """, file
        )

        conn.commit()

        
        
        print("Forecasts uploaded to database\n")
        


       
        #close the communication with the PostgreSQL
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed \n')

if __name__ == '__main__':
    connect()









