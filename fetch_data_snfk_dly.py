# -*- coding: utf-8 -*-
"""
Created on Tue Apr  6 10:33:22 2021

@author: Becky
"""

from __future__ import print_function
import os
import sys

country_var = sys.argv[1]
#country_var = 'us'

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
                    a.sale_date <=  DATE_TRUNC('week', CURRENT_DATE)  - INTERVAL '2 DAY' and
                    b.company = 'DK' and
                    lower(a.country) = '%s' and
                    a.POS_ACCT in ('A4', 'AMZ') and
                    format NOT IN ('EL','DN', 'EBK')
            group by 1,2,3,4,5,6,7,8
            order by 1 desc, 9 desc

        """% (country_var, country_var)
        
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




        #Bookscan fetch for forecasting
        sql_context = """
            
            with sub as ( 
                select 
                    b.region_code, 
                    a.prod_key,
                    sum(a.total_week_sold)
                
                from prh_global_dk_sandbox..bookscan_industry_us a
                join PRH_GLOBAL.PUBLIC.D_REGION_PROD b on a.prod_key = b.prod_key and b.region_code = 'US' and b.company = 'DK' 
                
                WHERE   
                    a.week_end_date  >= CURRENT_DATE - 28
                    
                
                group by 1,2
                
                HAVING sum(a.total_week_sold) > 5
                        )
            select
                a.week_end_date as date,
                'us' as country,
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
                sum(a.total_week_sold) as units
                
            from sub s
            left join prh_global_dk_sandbox..bookscan_industry_us  a
                ON s.prod_key = a.prod_key 
            left join PRH_GLOBAL.PUBLIC.D_REGION_PROD b
                ON s.prod_key = b.prod_key and  b.region_code = 'US'  and b.company = 'DK'

            where b.asin != ''
                
            group by 1,2,3,4,5,6,7,8
            order by 1 desc, 9 desc 

        """
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()
            
        #Writing csv file
        with open('Sales US Bookscan.csv', 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)
            print('Latest Bookscan Sales data saved \n')







        #Getting print status
        sql_context = """
        select material as isbn, sales_status_desc
        from PRH_GLOBAL.PUBLIC.D_REGION_PROD
        where lower(region_code) = '%s' and company = 'DK' and format not in ('EL', 'DN')
        """% (country_var)
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No Print Status')
            raise LatestDataCheck()

        #Writing csv file
        with open('Print Status %s.csv'% (country_var), 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record) 
            print('Latest print status data saved \n')
            
            # BH: only US need as the UK file has got the print status already


        #Getting AA status
        sql_context = """
        select isbn, curr_status  
        from PRH_GLOBAL_DK_SANDBOX..VW_AMAZON_PRIORITY_LIST
        where country = '%s'
        """% (country_var)
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No AA Status')
            raise LatestDataCheck()

        #Writing csv file
        with open('AA Status %s.csv'% (country_var), 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)   
            print('Latest AA status data saved \n')

 



        #Getting Inventory data
        sql_context = """
        with sub as ( 

            select  b.region_code, b.prod_key, b.title_full, b.material, b.asin, sum(qty_ord) 
            
            from PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES a
            left join PRH_GLOBAL.PUBLIC.D_REGION_PROD b
            
            ON a.prod_key = b.prod_key and a.region_code = b.region_code
            
            WHERE lower(a.country) = '%s' and  b.company = 'DK' and a.POS_ACCT in ('A4', 'AMZ') and
                format NOT IN ('EL','DN', 'EBK') and
                sale_date >= CURRENT_DATE - 28
            group by 1,2,3,4,5
            HAVING sum(qty_ord) > 5
        )

        , AMZ_INV_US as (

            select 
            a.country, a.week_end_date, b.asin ,b.material, b.title_full, 
            sum(QTY_INV) as AMZ_INV,
            'NA' as AMZ_Open_Orders
            
            from PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES_WK a
            join PRH_GLOBAL.PUBLIC.D_REGION_PROD b ON a.prod_key = b.prod_key and a.country = b.region_code
            
            where b.company = 'DK' and a.country = 'US'
            and a.POS_ACCT in ('AMZ')
            and week_end_date =   (select max(week_end_date) from PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES_WK where company = 'DK' and country = 'US' )
            
            group by 1,2,3,4,5 
            order by 1,2 desc,6 desc nulls last    
        )

        , AMZ_INV_UK as (

            select 
            a.country, a.forecast_date, b.asin ,b.material, b.title_full, 
            sum(available_inventory) as AMZ_INV,
            sum(open_purchase_order_quantity) as AMZ_Open_Orders
            
            from PRH_GLOBAL.PUBLIC.F_REGION_AMZ_FORECAST a
            join PRH_GLOBAL.PUBLIC.D_REGION_PROD b ON a.prod_key = b.prod_key and a.region_code = b.region_code
            
            where b.company = 'DK' and a.country = 'UK'
            and forecast_date =   (select max(forecast_date) from  PRH_GLOBAL.PUBLIC.F_REGION_AMZ_FORECAST where region_code = 'UK' )
            group by 1,2,3, 4, 5 
            order by 1,2 desc,6 desc nulls last

        )

        , DK_INV as (

            select 
                a.region_code, b.asin , b.material, b.title_full, 
                sum(a.net_inv) as TBS_INV
            
            from PRH_GLOBAL..D_REGION_INV_DAILY a
            left join PRH_GLOBAL.PUBLIC.D_REGION_PROD b	
            ON a.prod_key = b.prod_key and a.region_code = b.region_code       
            
            where 
                b.company = 'DK' 
                and lower(a.region_code) = '%s'
                and a.date = (select max(date) from  PRH_GLOBAL..D_REGION_INV_DAILY where lower(region_code) = '%s' )
            
            group by 1,2,3,4
            order by 1, 5 desc nulls last

        )
        , TOTAL_INV as (

            select a.country, a.prod_key ,b.material, b.title_full, a.week_end_date, sum(a.qty_inv) as retail_inv 
            
            from PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES_WK a 
            join PRH_GLOBAL.PUBLIC.D_REGION_PROD b ON a.prod_key = b.prod_key and a.region_code = b.region_code
            
            where lower(a.country) = '%s' and a.company = 'DK' and b.company = 'DK'  
            and a.week_end_date = (select max(week_end_date) from  PRH_GLOBAL.PUBLIC.F_REGION_POS_SALES_WK where lower(country)= '%s' and company = 'DK' )
            and pos_acct != 'BSN'
            
            group by 1,2,3,4,5
            order by retail_inv desc
            
        )

        select a.region_code, a.asin, a.material, a.title_full, 
        case when lower(a.region_code) = 'uk' then b.AMZ_INV
             when lower(a.region_code) = 'us' then e.AMZ_INV end as AMZ_INV,  
        c.TBS_INV, d.retail_inv , b.AMZ_Open_Orders  from sub a
        left join AMZ_INV_UK b on a.material = b.material and a.region_code = b.country
        left join AMZ_INV_US e on a.material = e.material and a.region_code = e.country
        left join DK_INV c on a.material = c.material and a.region_code = c.region_code
        left join TOTAL_INV d on a.material = d.material and a.region_code = d.country

        order by 5 desc nulls last
        """% (country_var, country_var, country_var, country_var, country_var)
        
        cur.execute(sql_context)
    
        # Fetch all rows from database
        record = cur.fetchall()

        #Checking if latest data is available
        if not record :
            print('No Inventory data')
            raise LatestDataCheck()

        #Writing csv file
        with open('Inventory %s.csv'% (country_var), 'w') as f:
            column_names = [i[0] for i in cur.description]
            file = csv.writer(f, lineterminator = '\n')
            file.writerow(column_names)
            file.writerows(record)   
            print('Latest Inventory data saved \n')
       
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








