# This is a sample Python script.
# encoding: utf-8


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os

from openpyxl import load_workbook
import sys
reload(sys)
sys.setdefaultencoding('utf8')

if __name__ == '__main__':
    source_in = (sys.argv[1])
    #source_in = 'F:\PP-02_03_04_05_06 V A F R - v4.0\A DCAC.xlsm'
    source_out = (sys.argv[2])
    #source_out = 'C:/temp/temp_labwiev_I.txt'
    wzorzec = (sys.argv[3])
    dane = load_workbook(source_in,data_only=True)
    try:
        os.remove(source_out)
    except:
        print("file not found")
    temp = open(source_out,"w")
    for komorka in range(9, 465, 24):
        if dane['karta']["B"+str(komorka+1)].value != None:
            zakres = float(dane['karta']["B"+str(komorka+1)].value)
            if ("µV" in dane['karta']["B"+str(komorka)].value) or ("µA" in dane['karta']["B"+str(komorka)].value):
                zakres = float(dane['karta']["B"+str(komorka+1)].value)/1000000
            if ("mV" in dane['karta']["B"+str(komorka)].value) or ("mA" in dane['karta']["B"+str(komorka)].value):
                zakres = float(dane['karta']["B"+str(komorka+1)].value)/1000
            if ("kHz" in dane['karta']["B" + str(komorka)].value) and ("V" not in dane['karta']["B"+str(komorka)].value) and ("A" not in dane['karta']["B"+str(komorka)].value):
                zakres = float(dane['karta']["B" + str(komorka + 1)].value) * 1000
            if "MHz" in dane['karta']["B" + str(komorka)].value and ("V" not in dane['karta']["B"+str(komorka)].value) and ("A" not in dane['karta']["B"+str(komorka)].value):
                zakres = float(dane['karta']["B" + str(komorka + 1)].value) * 1000000
            czestotliwosc=0
            if "Hz" in dane['karta']["B"+str(komorka)].value and ("V" in dane['karta']["B"+str(komorka)].value or "A" in dane['karta']["B"+str(komorka)].value):
                czestotliwosc=int(dane['karta']["B"+str(komorka)].value.split(" ")[2])
            if "kHz" in dane['karta']["B"+str(komorka)].value and ("V" in dane['karta']["B"+str(komorka)].value or "A" in dane['karta']["B"+str(komorka)].value):
                czestotliwosc=int(dane['karta']["B"+str(komorka)].value.split(" ")[2])*1000
            if "MHz" in dane['karta']["B"+str(komorka)].value and ("V" in dane['karta']["B"+str(komorka)].value or "A" in dane['karta']["B"+str(komorka)].value):
                czestotliwosc=int(dane['karta']["B"+str(komorka)].value.split(" ")[2])*1000000
            for i in ["D", "F", "H", "J", "L"]:
                if ((dane['karta'][i+str(komorka+4)].value != None) or (dane['karta'][i+str(komorka+3)].value != None)) and (dane['karta'][i+str(komorka-1)].value == wzorzec) and (dane['karta'][i+str(komorka+9)].value == None):
                    punkt = float(dane['karta'][i+str(komorka+6)].value)
                    if "µV" in dane['karta']["B"+str(komorka)].value or "µA" in dane['karta']["B"+str(komorka)].value:
                        punkt = float(dane['karta'][i+str(komorka+6)].value)/1000000
                    if "mV" in dane['karta']["B"+str(komorka)].value or "mA" in dane['karta']["B"+str(komorka)].value:
                        punkt = float(dane['karta'][i+str(komorka+6)].value)/1000
                    if "Hz" in dane['karta']["B"+str(komorka)].value and "V" not in dane['karta']["B"+str(komorka)].value and "A" not in dane['karta']["B"+str(komorka)].value:
                        punkt = 1
                        czestotliwosc = float(dane['karta'][i+str(komorka+6)].value)
                    if "kHz" in dane['karta']["B"+str(komorka)].value and "V" not in dane['karta']["B"+str(komorka)].value and "A" not in dane['karta']["B"+str(komorka)].value:
                        punkt = 1
                        czestotliwosc = float(dane['karta'][i+str(komorka+6)].value)*1000
                    if "MHz" in dane['karta']["B"+str(komorka)].value and "V" not in dane['karta']["B"+str(komorka)].value and "A" not in dane['karta']["B"+str(komorka)].value:
                        punkt = 1
                        czestotliwosc = float(dane['karta'][i+str(komorka+6)].value)*1000000
                    temp.write("{0:f}".format(zakres)+";"+"{0:f}".format(punkt)+";"+"{0:f}".format(czestotliwosc)+"\n")
    temp.close





# See PyCharm help at https://www.jetbrains.com/help/pycharm/
