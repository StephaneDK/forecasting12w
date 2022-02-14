# -*- coding: utf-8 -*-
"""
Created on Tue Apr  6 10:33:22 2021

@author: Stephane
"""

from __future__ import print_function
import os
import sys

country_var = sys.argv[1]

os.chdir('C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv')

class LatestDataCheck(Exception):
    pass

#Getting last saturday date
from datetime import date
from datetime import timedelta
today = date.today()
last_saturday = today - timedelta(days= (today.weekday() - 5) % 7)
last_saturday = last_saturday.strftime('%Y-%m-%d')




#Connecting to DB --------------------------------------------------------------------------------------------------------------------
#!/usr/bin/python
import psycopg2
from config import Config
from configparser import ConfigParser

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

        print('\nRetrieving data for forecasting')

        # display the PostgreSQL database server version
        db_version = cur.fetchone()
        print(db_version, ' \n ')
        
        try:

            sql_context = """
            SELECT * FROM public.vw_amazon_sales_traffic_display
            WHERE country = '%s' AND date = '%s'
            """ % (country_var,last_saturday)
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
            
            #Checking if latest data is available
            if not record :
                raise LatestDataCheck()
           
            sql_context = """
            SELECT * FROM public.vw_amazon_sales_traffic_display
            WHERE country = '%s'
            """% country_var
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
                

            #Writing csv file
            outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)



            with open('Sales %s.csv'% (country_var), 'w') as f:
                cur.copy_expert(outputquery, f)    
                print('Latest Sales data saved \n')

            

            


            #Getting print status
            sql_context = """
            select a.edition_answer_answer, b.isbn from 
            b3.b3_edition_answer a, b3.onix_b3 b
            where a.b3_id = b.id and country = 'US'
            """
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
                

            #Writing csv file
            outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)

            with open('Print Status US.csv', 'w') as f:
                cur.copy_expert(outputquery, f)    
                print('Latest print status data saved \n')



            #Getting AA status
            sql_context = """
            select asin, curr_status  from mvw_amazon_priority_list
            where country = 'us'
            """
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
                

            #Writing csv file
            outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)

            with open('AA Status US.csv', 'w') as f:
                cur.copy_expert(outputquery, f)    
                print('Latest AA status data saved \n')





            #Getting Amazon scrape info
            sql_context = """
            SELECT title_asin, availability_text FROM scrape.amazon_product_page_info
            where date_of_extraction = '%s' AND country = '%s'
            """% (last_saturday, country_var)
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
                

            #Writing csv file
            outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)

            with open('Scrape info %s.csv'% (country_var), 'w') as f:
                cur.copy_expert(outputquery, f)    
                print('Latest scrape data saved \n')




            
            #Getting AMZ Stock status
            sql_context = """
            select * from vw_amazon_stock_alert_14d
            """
            
            cur.execute(sql_context)
        
            # Fetch all rows from database
            record = cur.fetchall()
                

            #Writing csv file
            outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)

            with open('AMZ Stock status.csv', 'w') as f:
                cur.copy_expert(outputquery, f)    
                print('Latest stock status data saved \n')






                        
        except (Exception, LatestDataCheck):
            print('No values yet \n')
            print('ValueError1', file=sys.stderr)

            
       
        #close the communication with the PostgreSQL
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Forecasting data retrieved succesfully/n')
            print('Database connection closed/n')

if __name__ == '__main__':
    connect()








