import xlrd
import random
import tkinter as tk


def btn():
    row = random.randint(1, 6083)
    cell = sheet.row_values(row)
    word.set(cell[1])
    cn.set(cell[2])
    eg.set(cell[3])


def key(event):
    row = random.randint(1, 6083)
    cell = sheet.row_values(row)
    word.set(cell[1])
    cn.set(cell[2])
    eg.set(cell[3])


if __name__ == '__main__':
    book = xlrd.open_workbook('Words.xlsx')
    sheet = book.sheets()[0]

    # # 表格行数列数
    # nrows = sheet1.nrows
    # ncols = sheet1.ncols
    # # 第3行值
    # row3_values = sheet1.row_values(2)
    # print(row3_values)
    # # 第3列值
    # col3_values = sheet1.col_values(2)
    # # 第3行第3列的单元格的值
    # cell_3_3 = sheet1.cell(2, 2).value

    window = tk.Tk()
    window.title("HELLO WORD")
    window.geometry('500x300')

    word = tk.StringVar()
    cn = tk.StringVar()
    eg = tk.StringVar()
    btn()

    l1 = tk.Label(window, textvariable=word, font=('Arial', 24), width=30, height=3)
    l1.pack()
    l2 = tk.Label(window, textvariable=cn, font=('Arial', 20), width=100, height=2)
    l2.pack()
    l3 = tk.Label(window, textvariable=eg, font=('Arial', 20), width=100, height=3)
    l3.pack()
    b = tk.Button(window, text='下一个', font=('Arial', 20), width=10, height=2, command=btn)
    b.pack()
    window.bind("<Key>", key)
    window.mainloop()
