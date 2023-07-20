# sunny
A program to convert Ussing Chamber files (.cla) into .xlsx file ready to be plotted in R. Additionally the user can also choose to plot the graphs by just pressing "y"

Sunny is a program (written in python) that in only 0.23 seconds is able to convert the data produced by Ussing Chamber software,
from .cla in ordered .xlsx tables.

It has many advantages:
- each column starts with the corresponding title. Files can be therefore immediately imported into R
- intervals are specified in the Excel file
- Max and min for each interval are automatically recognized by the program, which plot also their difference
- it gives you the possibility to save and edit the graphs of your experiment just by pressing "y"

Before to start you need to have installed the following python libraries:
(to instal them you can easily execute pip install in cmd)

datetime
xlsxwriter
matplotlib.pyplot

At line 7 you need to specifiy which file you want to convert.
I.e. :
file_pathway = "C:\\Users\\Adriano\\Desktop\\sunny v3\\example.cla"
(Remember always to use double backslash because it is a special character)

later the programm will ask you how to save your file. Please, beside the name, specify also the extension 
(i.e. example.xlsx). Extension must be .xlsx

The programm will finally ask you if you want to print the graphs relative to your data.
Only two answers are possible "y" or "n".
By printing "n" the program will convert your file by genereting an Excel file without creating/showing any graph.
By printing "y" the programm will open a new window and show you the I, Dp and Dp0 graphs relative to your 
data. You can decide to save them by pressing on the floppy disk icon or to close the window and pass directly
to your next graph. After that the 3rd window the program will convert the .cla file into a .xlsx file

by question please write me at: Adriano.Sanna@anatomie.med.uni-giessen.de
