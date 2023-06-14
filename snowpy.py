import pandas as pd
from snowflake.snowpark import Session

connection_parameters = {
  "account": "prh.us-east-1",
  "user": "snichanian",
  "password": "Kaepernick21",
  "role": "SNICHANIAN_USER_ROLE",
  "warehouse": "PRH_GLOBAL",
  "database": "PRH_GLOBAL",
  "schema": "PUBLIC"
}

session = Session.builder.configs(connection_parameters).create()

df_sql = session.sql("select * from PRH_GLOBAL.PUBLIC.D_REGION_PROD where company = 'DK' and region_code in ('US', 'UK') limit 100").show()
df_sql.show()
