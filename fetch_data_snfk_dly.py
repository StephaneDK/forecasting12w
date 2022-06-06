# -*- coding: utf-8 -*-
"""
Created on Tue Apr  6 10:33:22 2021

@author: Becky
"""

from __future__ import print_function
import os
import sys

country_var = sys.argv[1]
#country_var = 'uk'

os.chdir('C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv')
#os.chdir('C:\\dev\\env\\dkwebsols-forecasting-12w-22ef3797943e')

class LatestDataCheck(Exception):
    pass

#Getting last saturday date
from datetime import date
from datetime import timedelta
today = date.today()
last_saturday = today - timedelta(days= (today.weekday() - 5) % 7)
last_saturday = last_saturday.strftime('%Y-%m-%d')
#last_saturday = '2022-02-19' 


#Connecting to DB --------------------------------------------------------------------------------------------------------------------
#!/usr/bin/python
import snowflake.connector as sfc
import psycopg2
from config import Config
from configparser import ConfigParser
import csv



def config(filename='database.ini', section='snowflake_db'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)

    # get section, default to snowflake
    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))

    return db


def connect():
    """ Connect to the GDH server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the GDH server
        print('\nConnecting to the GDH database...')
        conn = sfc.connect(account='PRH',
                                region='us-east-1',
                                user= params['user'],
                                password= params['password'])
		
        # create a cursor
        cur = conn.cursor()
        
        #execute a statement
        print('Snowflake database version:')
        cur.execute('SELECT CURRENT_VERSION()')

        print('\nRetrieving data for forecasting')

        # display the Snowflake database server version
        db_version = cur.fetchone()
        print(db_version, ' \n ')
        

        
        sql_context = """
            with sub as ( select  b.region_code, b.prod_key,  sum(qty_ord) from PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES a
                    left join PRH_GLOBAL.PUBLIC.D_REGION_PROD b
                    ON a.prod_key = b.prod_key and a.region_code = b.region_code
                    WHERE lower(a.country) = '%s' and  b.company = 'DK' and a.POS_ACCT in ('A4', 'AMZ') and
                        format NOT IN ('EL','DN', 'EBK') and
                        sale_date >= CURRENT_DATE - 28
                    group by 1,2
                    HAVING sum(qty_ord) > 5
                            )
                select
                    c.week_end_date as date,
                    lower(a.country) as country,
                    b.asin,
                    b.material as isbn,
                    b.title_full as title,
                    CASE
                        WHEN imprint_desc ilike '%%Garrick Street Press%%' THEN 'Garrick Street Press'
                        WHEN imprint_desc ilike '%%0-9%%' or imprint_desc ilike '%%0 - 9%%' THEN 'Children 0-9'
                        WHEN imprint_desc ilike '%%Licensi%%' THEN 'Licensing'
                        WHEN imprint_desc ilike '%%Adult%%' AND imprint_desc ilike '%%Knowledg%%' THEN 'Knowledge Adult'
                        WHEN imprint_desc ilike '%%Children%%' AND imprint_desc ilike '%%Knowledg%%' THEN 'Knowledge Children'
                        WHEN imprint_desc ilike '%%Alpha%%' THEN 'Alpha'
                        WHEN imprint_desc ilike '%%Life%%' THEN 'Life'
                        WHEN imprint_desc ilike '%%Travel%%' THEN 'Travel'
                        ELSE 'Others'
                        END AS division,
                    b.imprint_desc as imprint,
                    b.onsale_date as pub_date,
                    sum(a.qty_ord) as units
                from sub s
                left join PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES a
                    ON s.prod_key = a.prod_key and s.region_code = a.region_code
                left join PRH_GLOBAL.PUBLIC.D_REGION_PROD b
                    ON a.prod_key = b.prod_key and a.region_code = b.region_code
                left join PRH_GLOBAL..D_region_date c 
                    on a.sale_date = c.date
                where
                    extract(year from a.sale_date) >= '2017' and
                    a.sale_date <= (SELECT MAX(WEEK_END_DATE) FROM PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES_WK WHERE lower(REGION_CODE) = '%s' and POS_ACCT in ('A4', 'AMZ') ) and
                    b.company = 'DK' and
                    lower(a.country) = '%s' and
                    a.POS_ACCT in ('A4', 'AMZ') and
                    format NOT IN ('EL','DN', 'EBK')
            group by 1,2,3,4,5,6,7,8
            order by 1 desc, 9 desc

        """% (country_var, country_var,country_var)
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()
            
        #Writing csv file
        with open('Sales %s.csv'% (country_var), 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)
            print('Latest Sales data saved \n')




        #Getting print status
        sql_context = """
        select material as isbn, sales_status_desc
        from PRH_GLOBAL.PUBLIC.D_REGION_PROD
        where region_code = 'US' and company = 'DK' and format not in ('EL', 'DN')
        """
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No Print Status')
            raise LatestDataCheck()

        #Writing csv file
        with open('Print Status US.csv', 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record) 
            print('Latest print status data saved \n')
            
            # BH: only US need as the UK file has got the print status already


        #Getting AA status
        sql_context = """
        select asin, curr_status 
        from PRH_GLOBAL_DK_SANDBOX..VW_AMAZON_PRIORITY_LIST
        where country = 'us'
        """
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No AA Status')
            raise LatestDataCheck()

        #Writing csv file
        with open('AA Status US.csv', 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)   
            print('Latest AA status data saved \n')



        

        #Getting Amazon scrape info
        sql_context = """
        select asin, availability_text
        from PRH_GLOBAL_DK_SANDBOX..AMZ_PRODUCT_PAGE
        where date_of_extraction = '%s' AND lower(country) = '%s'
        """% (last_saturday, country_var)
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No Amazon scrape')
            raise LatestDataCheck()  

        #Writing csv file
        with open('Scrape info %s.csv'% (country_var), 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)  
            print('Latest scrape data saved \n')



        #BH: Ask Stephane if we still need it
        
        #Getting AMZ Stock status
        #sql_context = """
        #select * from vw_amazon_stock_alert_14d
        #"""
        
        #cur.execute(sql_context)
    
        # Fetch all rows from database
        #record = cur.fetchall()
            

        #Writing csv file
        #outputquery = "COPY ({0}) TO STDOUT WITH CSV HEADER".format(sql_context)

        #with open('AMZ Stock status.csv', 'w') as f:
        #    cur.copy_expert(outputquery, f)    
        #    print('Latest stock status data saved \n')



                        


            
       
        #close the communication with the PostgreSQL
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Forecasting data retrieved succesfully\n')
            print('Database connection closed\n')

if __name__ == '__main__':
    connect()








