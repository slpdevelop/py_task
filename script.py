import requests
import re
import MySQLdb
from datetime import datetime
from openpyxl import Workbook



db = MySQLdb.connect(host="your_host", user="your_user", passwd="your_pass", db="your_db")
cursor = db.cursor()
table_name = "city_time"
city_names = ['London', 'Moscow', 'New_York', 'Tokyo', 'Beijing']

create_table = 'create table if not exists {0} (city varchar(256) null, time_city time(0) null, our_time time(0) null, difference int null)'.format(table_name)

cursor.execute(create_table)
db.commit()


cursor.execute("SELECT count(city) FROM city_time")
r=cursor.fetchall()[0][0]
if(r == 0):
    for city in city_names:
        cursor.execute('insert into {0} (city) values ("{1}")'.format(table_name, city))
    db.commit()




class CityTime:

    def __init__(self, city):
        self.city = city

    def getTime(self):
        page = requests.get('https://time.is/Ru', headers = self.setHeaders())
        city_html = re.findall('<a href="\/{0}" id=.*?<\/a>'.format(self.city), page.text)
        if (len(city_html)==0):
            return 'Город -- {0} не найден'.format(self.city)
        time_str = re.findall('<span.*?>(.*?)<\/span>', city_html[0])[0]
        time = datetime.strptime(time_str, '%H:%M').time()
        return time

    def setHeaders(self):
        return {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36'
        }    

our_time = datetime.now()
cursor.execute("SELECT city FROM {0}".format(table_name))
wb = Workbook()

for row in cursor.fetchall():
    city = row[0]
    city_time = CityTime(city).getTime()
    if (type(city_time) is str):
        print(city_time)
        continue
    difference = our_time.hour - city_time.hour
    sql = 'UPDATE {4} SET time_city="{0}", our_time="{1}", difference={2} WHERE city="{3}"'.format(city_time.strftime('%H:%M:%S'), our_time.strftime('%H:%M:%S'), difference, city, table_name)
    cursor.execute(sql)
    db.commit()
    


cursor.execute("SELECT city, DATE_FORMAT(time_city, '%H:%i'), DATE_FORMAT(our_time, '%H:%i'), difference FROM {0}".format(table_name))
ws = wb.create_sheet(0)
ws.title = "pikta_task"

for row in cursor.fetchall():
    ws.append(row)

workbook_name = "pikta_task"
wb.save(workbook_name + ".xlsx")

