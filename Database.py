import pyodbc
import pandas
import array as arr
###################################
#Database connection
date='2020-03-20'
connection=pyodbc.connect(Driver='{SQL Server}',Server='192.168.10.30',Database='Cheetah',UID ='plcadmin',PWD = 'plc1310')
a=pyodbc.Cursor
result = pandas.read_sql("select Partstatus from Cheetah.dbo.S1141 where Timestamp > '2020-03-20 08:21:34'",connection,params=None)
###################################
statuscode=[0]*len(result)
for row in range(0, len(result)):
    statuscode[row]=result._get_value(row, 'Partstatus')
statuscode = list(dict.fromkeys(statuscode))
print(statuscode)

connection.close()

