import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
#from os import listdir


# returns a month it's number
def num_month(month):
    if month == "Styczeń":
        n_m = 1
    elif month == "Luty":
        n_m = 2
    elif month == "Marzec":
        n_m = 3
    elif month == "Kwiecień":
        n_m = 4
    elif month == "Maj":
        n_m = 5
    elif month == "Czerwiec":
        n_m = 6
    elif month == "Lipiec":
        n_m = 7
    elif month == "Sierpień":
        n_m = 8
    elif month == "Wrzesień":
        n_m = 9
    elif month == "Październik":
        n_m = 10
    elif month == "Listopad":
        n_m = 11
    elif month == "Grudzień":
        n_m = 12
    else:
        return False
    return n_m


# looks for shop's name and returns it's cell as a variable
def lf_shop():
    shop = {}
    for i in range(int(sh.max_row + 1)):
        for j in range(int(sh.max_column + 1)):
            if sh.cell(row=i+1, column=j+1).value in shops:
                shop = sh.cell(row=i+1, column=j+1)
    return shop


# returns dictionary: [worker's_number: worker's_name]
def list_of_workers():
    workers = {}
    for i in range(4, int(sh.max_column)):
        if type(sh.cell(row=13, column=i).value) == str:
            name = (sh.cell(row=13, column=i)).value
            surname = (sh.cell(row=14, column=i)).value
            workers[i] = (name + " " + surname) # key in dictionary is a number of column for each worker
    return workers


# creating variables for specifying the sheet you are looking for and giving month it's number
month = (input("Wprowadz słownie miesiąc: ")).title()
n_m = num_month(month)

# list of our shops
shops = ["DBC", "DKD", "DGS", "DGK", "DO", "DPP", "DPA", "DSP", "DWW", "DWB", "DLM", "DWT"]

# opening work schedule file
wb = openpyxl.load_workbook("grafik.xlsx", data_only=True)
sh = wb.get_sheet_by_name("Grafik_"+month)

# looking for shop's name and setting it's code to variable
shop = lf_shop() # variable set to cell object

# making list of workers
workers = list_of_workers()


shop_n_row = shop.row
shop_n_col = column_index_from_string(shop.column)
# looking for year
year = sh.cell(row=shop_n_row+2, column=shop_n_col)

for worker in workers:
    worker_hours = 0
    a_cell = sh.cell(row=shop_n_row+3, column=int(worker)+3)
    for i in range(0,sh.max_row):
        c_day = sh.cell(row=(shop_n_row+3)+i, column=shop_n_col+1)

        # counts employee's sum of worked hours
        if type(c_day.value) == int:
            worker_hours += sh.cell(row=c_day.row, column=column_index_from_string(a_cell.column) - 1).value

        # looks for event
        if type(c_day.value) == int:
            a_cell = sh.cell(row=c_day.row, column=int(worker)+3)

            # checking value for vacation leave
            if a_cell.value == "UW":
                begin_day = sh.cell(row=a_cell.row, column=column_index_from_string(c_day.column))
                end_day = begin_day
                vl = True
                while vl == True:
                    if a_cell.value == sh.cell(row=a_cell.row+1, column=column_index_from_string(a_cell.column)).value:
                        end_day = sh.cell(row=a_cell.row+1, column=column_index_from_string(c_day.column))
                        a_cell = sh.cell(row=end_day.row+1, column=column_index_from_string(a_cell.column))
                        print("kolejny")
                    else:
                        #end_day = sh.cell(row=a_cell.row, column=column_index_from_string(c_day.column))
                        #print(begin_day)
                        #print(end_day)
                        #vl = False
                        if end_day.value == begin_day.value:
                            a_cell = sh.cell(row=end_day.row + 1, column=column_index_from_string(a_cell.column))
                            print("%s Urlop wypoczynkowy w dniu: %02d.%02d.%04d" % (workers[worker], begin_day.value, n_m, year.value))
                            vl = False
                        else:
                            a_cell = sh.cell(row=end_day.row + 1, column=column_index_from_string(a_cell.column))
                            print("%s Urlop wypoczynkowy w dniach: %02d-%02d.%02d.%04d" % (workers[worker], begin_day.value, end_day.value, n_m, year.value))
                            vl = False








        else:
            break
    print("\n%s w miesiącu %s przepracował %s godzin" % (workers[worker], month.lower(), worker_hours))

#print("Urlop wypoczynkowy w dniach: %02d-%02d.%04d" % (begin_day.value, end_day.value, year.value))