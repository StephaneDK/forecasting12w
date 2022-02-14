# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 15:47:43 2021

@author: Stephane
"""

from __future__ import print_function
import os
import sys




os.chdir('C:\\Users\\steph\\Documents\\DK\\Work\\Forecasting book sales and inventory\\Pipeline\\csv')


#country_var = sys.argv[1]

class LatestDataCheck(Exception):
    pass




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


        conn = psycopg2.connect(**params)
		
        # create a cursor
        cur = conn.cursor()
        

        


            
            

        
        sql_context = """
        WITH cte AS
        (
        SELECT ctid,
               row_number() OVER (PARTITION BY region, model, asin,pred_at,pred_for,pred_units
                                  ORDER BY asin) rn
               FROM experiment.forecasting_temp
        )
        DELETE FROM experiment.forecasting_temp
               USING cte
               WHERE cte.rn > 1
                     AND cte.ctid = forecasting_temp.ctid;
        """ 
            
        cur.execute(sql_context)
        conn.commit()





        sql_context = """
        WITH cte AS
        (
        SELECT ctid,
               row_number() OVER (PARTITION BY region, model, asin,pred_at, 'Amz_inv', 'DK_inv', 'Total_inventory', 'Forecast_12w', 'Adjusted forecast_12w', 'Weeks_on_hand', 'Weeks_on_hand_AMZ', 'Inv_issue', 'reprint_date', 'reprint_quantity'
                                  ORDER BY asin) rn
               FROM experiment.forecasting_statistics_temp
        )
        DELETE FROM experiment.forecasting_statistics_temp
               USING cte
               WHERE cte.rn > 1
                     AND cte.ctid = forecasting_statistics_temp.ctid;
        """ 
            
        cur.execute(sql_context)
        conn.commit()

        
        print("Omitted duplicate rows")
        

        
        

                        
#        except (Exception, LatestDataCheck):
#            print('Error writing \n')


       
        #close the communication with the PostgreSQL
        cur.close()
        

        
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()


if __name__ == '__main__':
    connect()








