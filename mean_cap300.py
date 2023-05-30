from random import choice as ch
import openpyxl


def mean_cap300(mainsheet,file,caps):
    book = openpyxl.load_workbook(file)
    sheet = book[mainsheet]
    cell_B = "B"
    cell_C = "C"
    cell_D = "D"

    for i in range(300):
        cell1 = cell_B + str(i+2)
        cell2 = cell_C + str(i+2)
        cell3 = cell_D + str(i+2)
        tmp = []
        for j in range(10):
            tmp.append(ch(caps))
        print(tmp)

        sheet[cell1] = tmp.count("1")
        sheet[cell2] = tmp.count("2")
        sheet[cell3] = tmp.count("3")
    book.save(file)
    print("SAVED.")