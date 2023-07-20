#Developed by Adriano Sanna at the Institute for Anatomy and Cell Biology, JLU Giessen. Finished on 14.07.2023

import matplotlib.pyplot as plt
import xlsxwriter
from datetime import datetime

file_pathway = "C:\\Users\\Adriano\\Desktop\\sunny v3\\example.cla"

name_excel_file =input(" please writhe the name of the excel file: ")

open_file = open( file_pathway, "r")

Lines = open_file.readlines()

Zeit = []
dP = []
dP0= []
I= []
Rt= []
Marker= []
Textnr= []
IndexB= []
IndexG= [] 

#"Zeit";"dP";"dP0";"I(µA/cm²)";"Rt";"-";"Marker/Fehler";"Textnr.";"Index Betr.-Art.";"Index Gt/Rt-Par."

#start it finds where the data collected from the machine are
count=0
starting_line_n_list=[]
starting_line_n = 0

for line in Lines:
    if line.strip().split(";")[0] =='00:00:00':
        starting_line_n == (Lines.index(line))
        starting_line_n_list.append((Lines.index(line)))

starting_line_n = starting_line_n_list[0]
#end

#start it appends to the several lists the corresponding str values registered with the Ussing Chamber
for line in Lines:
    count += 1
    if Lines.index(line) >=starting_line_n:
        a = line.strip()
        Zeit.append(a.split(";")[0])
        dP.append(a.split(";")[1])
        dP0.append(a.split(";")[2])
        I.append(a.split(";")[3])
        Rt.append(a.split(";")[4])
        Marker.append(a.split(";")[6])
#end

Zeit_num =[]
I_num =[]
Marker_positions=[]
dP_num= []
dp0_num=[]

for x in range(0,len(Zeit)):
    Zeit_num.append(6*x)
    
#convert str values into float
for c in range(len(I)):
    d= I[c].replace(" ", "")
    d= I[c].replace(",", ".")
    e = float(d)
    I_num.append(e)

for c in range(len(dP)):
    d= dP[c].replace(" ", "")
    d= dP[c].replace(",", ".")
    e = float(d)
    dP_num.append(e)

for c in range(len(dP0)):
    d= dP0[c].replace(" ", "")
    d= dP0[c].replace(",", ".")
    e = float(d)
    dp0_num.append(e)

Zeit_num_minutes =[]

#Determining vertical lines position (X) for then the pharmaceutics were administered
for c in range(len(Marker)):
    d= Marker[c].replace(" ", "")
    d= Marker[c].replace(",", ".")
    e = int(d)
    if e !=0.0:
        Marker_positions.append(c)

#plot the vertical lines 
def plotting_vertical_lines():
    for x in Marker_positions:
        plt.axvline(x = x, color = 'b', label = 'axvline - full height',linewidth=0.60)

max_list=[]

Positions = []
for x in range(len(Zeit_num)):
    Positions.append(x)

interval_max_dP =[]
interval_markers_dP=[]
interval_max_dP0=[]
interval_markers_dP0=[]
interval_max_I=[]
interval_markers_I=[]


Marker_positions_with_0_and_last_value =[0]
for x in Marker_positions:
    Marker_positions_with_0_and_last_value.append(x)
Marker_positions_with_0_and_last_value.append(Positions[-1])
#print(Marker_positions_with_0_and_last_value)

def getting_max(value1,value2):
    temporary_dP_interval_max=[]
    temporary_dP0_interval_max=[]
    temporary_I_interval_max=[]
    for number in dP_num[value1:value2]:
            temporary_dP_interval_max.append(number)
    for number in dp0_num[value1:value2]:
            temporary_dP0_interval_max.append(number)
    for number in I_num[value1:value2]:
            temporary_I_interval_max.append(number)
    max_dP=max(temporary_dP_interval_max)
    max_dP0=max(temporary_dP0_interval_max)
    max_I=max(temporary_I_interval_max)
    interval_max_dP.append(max_dP)
    interval_max_dP0.append(max_dP0)
    interval_max_I.append(max_I)
    print(value1)
    print(value2)
    print(max_dP)

interval1=0
interval2=1

interval_markers_dP.append(0)
interval_markers_dP0.append(0)
interval_markers_I.append(0)


for x in Marker_positions:
    interval_markers_dP.append(dP_num[x])
    interval_markers_dP0.append(dp0_num[x])
    interval_markers_I.append(I_num[x])

#Creating an excel file
workbook = xlsxwriter.Workbook(name_excel_file)
worksheet = workbook.add_worksheet()

#adding the first line of labels
Excels_Labels =["Time", "Time (s)", "dP",
                "dP0","I(µA/cm²)", "Markers Interval",
                "Max I", "Min I","delta I"]
row= 0
col= 0
for x in (Excels_Labels):
    worksheet.write(row, col, x)
    col += 1

#calculating delta Dp, Dp0, I

delta_I=[]
max_I_list=[]
min_I_list=[]

def max_min_I(x,x1):
    maximum = max(I_num[x:x1])
    minimum = min(I_num[x:x1])
    max_I_list.append(maximum)
    min_I_list.append(minimum)

for x in Marker_positions_with_0_and_last_value:
    if x == Marker_positions_with_0_and_last_value[-1]:
        break
    else: 
        x1=Marker_positions_with_0_and_last_value[Marker_positions_with_0_and_last_value.index(x)+1]
        print(x)
        print(x1)
        max_min_I(x,x1)

for x in max_I_list:
    if min_I_list[max_I_list.index(x)] >= 0:
        delta_I.append(x - min_I_list[max_I_list.index(x)])
    
    elif min_I_list[max_I_list.index(x)] <= 0 and x < 0:
        delta_I.append(((x*-1) + min_I_list[max_I_list.index(x)])*-1)
    
    elif min_I_list[max_I_list.index(x)] <= 0 and x >=0 :
        delta_I.append(x + (min_I_list[max_I_list.index(x)])*-1)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for x in (Zeit):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=1
for x in (Zeit_num):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=2
for x in (dP_num):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=3
for x in (dp0_num):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=4
for x in (I_num):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=5
for x in range(len(Marker_positions_with_0_and_last_value)-1):
    left = Marker_positions_with_0_and_last_value[x]
    x_2 = x +1
    right=Marker_positions_with_0_and_last_value[x_2]
    range_ = str(left)+"-"+str(right)
    worksheet.write(row, col, range_)
    row += 1

row = 1
col=6
for x in (max_I_list):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=7
for x in (min_I_list):
    worksheet.write(row, col, x)
    row += 1

row = 1
col=8
for x in (delta_I):
    worksheet.write(row, col, x)
    row += 1

workbook.close()

open_file.close

message_2 = input("Graph? [y/n] ")
if message_2 == "y":
    plt.plot(Positions, I_num,"k",linewidth=0.60)
    plotting_vertical_lines()
    plt.title("I")
    plt.xlabel("6s")
    plt.ylabel("µA/cm²")
    plt.show()
    plt.plot(Positions, dP,"k",linewidth=0.60)
    plotting_vertical_lines()
    plt.title("dP")
    plt.xlabel("6s")
    plt.ylabel("mV")
    plt.show()
    plt.plot(Positions, dP0,"k",linewidth=0.60)
    plotting_vertical_lines()
    #plt.xticks(fontsize=14, rotation=90)
    plt.yticks(fontsize=6)
    plt.title("dP0")
    plt.xlabel("6s")
    plt.ylabel("mV")
    plt.show()

