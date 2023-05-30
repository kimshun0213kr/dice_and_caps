from random import randint as rd
import openpyxl


def dice100(mainsheet,file):
    book = openpyxl.load_workbook(file)
    sheet = book[mainsheet]
    cell = "B"
    for i in range(100):
        cell1 = cell + str(i+2)
        num = rd(1,6)
        print(num)
        sheet[cell1] = num
    book.save(file)
    print("SAVED.")