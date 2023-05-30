from random import randint as rd
import datetime
import dice100
import mean_dice100
import cap300
import mean_cap300

caps = []
for a in range(24):
    caps.append("1")
for b in range(57):
    caps.append("3")
for c in range(19):
    caps.append("2")

tmp = input("excelファイルのパスをそのままコピペしてください。>>")
file_space = tmp.replace("\"","")

# file_space = "C:\\Users\\\sicp3\\Desktop\\生命科学演習\\根本先生\\22RB035ShunsukeKimura.xlsx"
dice100.dice100("Dice100",file_space)
mean_dice100.mean_dice100("Mean_Dice100",file_space)
cap300.cap300("Cap300",file_space,caps)
mean_cap300.mean_cap300("Mean_Cap300",file_space,caps)

print(file_space+"に正しくデータが保存されました。")
print(datetime.datetime.now())