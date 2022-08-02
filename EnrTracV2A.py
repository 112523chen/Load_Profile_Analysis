from openpyxl import load_workbook, Workbook
import pandas as pd
import operator
from collections import OrderedDict
import os
from matplotlib import pyplot as plt

#!inputs
#* file used for script
workbook = 'RTM_Export_CUNY College of Technology_2017-08-01_2022-08-01_20220801234211.xlsx'
#* period for the percent change for energy consumption (3 ~ percentage change from row 1 to row 4)
p = 5
# #* level of significances for percentage(marker for greater percent change)
# lvl = 1
# # * level of significances for frequency (percentage of certain value in all periods ~ when p=5 and lvl2=0.6 row must be in 3 of the periods )
# lvl2 = 0.6


#pulls data from excel file
wb2 = load_workbook(workbook)
ws2 = wb2['Data']

#removes first row of worksheet
points = list()
ws2.delete_rows(0, 1)

#creates dataframe from the rows of the worksheet
for i in ws2.values:
    points.append(i)
df = pd.DataFrame(points)
df.columns = df.iloc[0]

#updates data types for the weather columns from string to float64
c = list(df.columns)[1:4]
for x in c:
    df[x] = pd.to_numeric(df[x], errors='coerce')

#deletes machine energy consumptions columns and creates a new column for the collective energy consumption
c = list(df.columns)[0:4]
for i in range(4, len(df.columns)):
    c.append(str(i))
df.columns = c
df["total"] = df[c[4:]].sum(axis=1)
c.append("total")
targets = c[:4]
targets.append("total")
df = df[(targets)]
df = df.iloc[1:]

#adds columns for references within code about date and time
df["year"] = pd.DatetimeIndex(df["Interval End"]).year
df["month"] = pd.DatetimeIndex(df["Interval End"]).month
df["day"] = pd.DatetimeIndex(df["Interval End"]).day
df["hour"] = pd.DatetimeIndex(df["Interval End"]).hour

#adds columns for references within code about percent change of energy consumption, dry-bulb weather, wet-bulb weather, and dew-point weather
items = ["total", "Dry-Bulb F", "Wet-Bulb F", "Dew-Point F"]
for item in items:
    key = ""
    if(item == "total"):
        key = "ENE"
    elif(item == "Dry-Bulb F"):
        key = "DRY"
    elif(item == "Wet-Bulb F"):
        key = "WET"
    elif(item == "Dew-Point F"):
        key = "DEW"
    for x in range(1,(p+1)):
        #energy consumption
        df[f"{x}f~{key}"] = df["total"].pct_change(periods=x)

#creates points for reference to find row for report: tuple(year,month,energy_consumption)
values = (df.groupby(["year", "month"])["total"].max().tolist())
points = list(df.groupby(["year","month"]).groups.keys())
for count, point in enumerate(points):
    points_ = list(point)
    points_.append(values[count])
    points[count] = tuple(points_)

#creates sub directory with report as txt file (soon with graphs)
for point in points:
    data = df[(df.year == point[0]) & (df.month ==  point[1]) & (df.total ==  point[2])]

    date = f"Date:{data.month.values[0]}-{data.day.values[0]}-{data.year.values[0]}\n"
    energy = f"Energy Consumption: {data.total.values[0]}Kwh\n"
    temps = f"Weather Data:\nDry-Bulb (F): {data['Dry-Bulb F'].values[0]}\nWet-Bulb (F): {data['Wet-Bulb F'].values[0]}\nDew-Point (F): {data['Dew-Point F'].values[0]}\n"

    if((os.path.exists("./peak_reports")) == False):
        os.mkdir("peak_reports")
    os.chdir("./peak_reports/")
    os.mkdir(
        f"{data.year.values[0]}-{data.month.values[0]}-{data.day.values[0]}_{data.total.values[0]}Kwh")
    os.chdir(
        f"./{data.year.values[0]}-{data.month.values[0]}-{data.day.values[0]}_{data.total.values[0]}Kwh")
    info = (f"{date}\n{energy}\n{temps}")
    f = open("report.txt", "a")
    f.write(info)
    os.chdir("../../")