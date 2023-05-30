from random import choice as ch
import openpyxl


def cap300(mainsheet,file,caps):
    book = openpyxl.load_workbook(file)
    sheet = book[mainsheet]
    cell = "B"

    for i in range(300):
        cell1 = cell + str(i+2)
        cap_res = ch(caps)
        print(cap_res)
        sheet[cell1] = cap_res
    book.save(file)
    print("SAVED.")