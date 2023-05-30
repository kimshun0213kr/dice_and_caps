from random import randint as rd
import openpyxl

# book = openpyxl.load_workbook("C:\\Users\\sicp3\\Desktop\\生命科学演習\\根本先生\\data.xlsx")

def mean_dice100(mainsheet,file):
    book = openpyxl.load_workbook(file)
    sheet = book[mainsheet]
    cell_B = "B"
    cell_C = "C"
    cell_D = "D"
    for i in range(100):
        cell1 = cell_B + str(i+2)
        cell2 = cell_C + str(i+2)
        cell3 = cell_D + str(i+2)
        num1 = rd(1,6)
        num2 = rd(1,6)
        num3 = rd(1,6)
        print(num1,num2,num3)
        sheet[cell1] = num1
        sheet[cell2] = num2
        sheet[cell3] = num3
    book.save(file)
    print("SAVED.")
