# -*- coding: utf-8 -*-
"""
script to access MS Access DAtabase

@author: volom
"""

import pyodbc
import pandas as pd
from datetime import datetime

now = datetime.now()
current_time = now.strftime("date %Y_%M_%d time %H_%M_%S")

selected_dir = input("Choose dir to download tables ")
path_accdb_file = input(r'Choose dir to MS Access DB path\name.accdb ')


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' + f'DBQ={path_accdb_file};')
cursor = conn.cursor()


# указание таблиц
print("Available tables:")
for i in cursor.tables():
    if i[3] == 'TABLE':
        print(i[2])
 
while True:  
    while True:
        try:
            query = input("Insert SQL query ")
            data = pd.read_sql(query, conn)
            print(data.info)
        except Exception as e:
            print("The query is not allowed!")
            print(e)
            print("Try again")
            try:
                del data
            except:
                pass
            continue
        else:
            break
    
    ask2down = input("Download? [y/n] ")
    
    if ask2down.lower() == 'y':
        data.to_csv(f'{selected_dir}\db_result_{current_time}.csv')
        print("Download is successful")
        cont = input("Continue? [y/n] ")
        if cont.lower() == 'y':
            continue
        else:
            break
    else:
        cont = input("Continue? [y/n] ")
        if cont.lower() == 'y':
            continue
        else:
            break
        



