# -*- coding: utf-8 -*-

import MySQLdb


def connectToDB():
    dbconnect = MySQLdb.connect("localhost", "root", "12345", "mail")
    
    cursor = dbconnect.cursor()
    
    query = 'insert into mails values (22,"a","a","a")'
    try:
       cursor.execute(query)
       dbconnect.commit()
    except:
       dbconnect.rollback()
    finally:
       dbconnect.close()
        
