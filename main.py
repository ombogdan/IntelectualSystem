import csv
import tkinter as tk
from decimal import Decimal
from tkinter import messagebox

import xlrd
from sympy.codegen.tests.test_applications import np


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        root.geometry("700x600+390+130")
        self.create_widgets()


    def create_widgets(self):
        global e
        e = tk.Entry(width=20)
        # e.pack()
        btn1 = tk.Button(text="Розрахувати", background="#555", foreground="#ccc", padx="15", pady="6", font="15")
        btn1['command'] = self.say_hi
        btn1.pack(side='bottom')
        quit = tk.Button(text="QUIT", fg="red",
                         command=self.master.destroy)
        quit.pack(side="top")


    def say_hi(self):
        data = self.csv_from_excel()
        coefficients = self.counting(data)
        F = coefficients[0] + coefficients[1]*0.593+coefficients[2]*0.683+coefficients[3]*0.683*0.593
        message = e.get()
        messagebox.showinfo("GUI Python", F)
        # messagebox.showinfo("GUI Python", count)

    def counting(self, data):
        a1 = []
        b1 = []
        c1 = []
        d1 = []
        t1 = []

        a2 = []
        b2 = []
        c2 = []
        d2 = []
        t2 = []

        a3 = []
        b3 = []
        c3 = []
        d3 = []
        t3 = []

        a4 = []
        b4 = []
        c4 = []
        d4 = []
        t4 = []

        for i in range(len(data)):
            if i > 1:
                a1.append(data[i])
                b1.append(1)
                c1.append(data[i - 2] * 1)
                d1.append(data[i - 1])
                t1.append(data[i - 1] * data[i - 2])

        for l in range(len(a1)):
            a2.append(a1[l] * c1[l])
            b2.append(b1[l] * c1[l])
            c2.append(c1[l] * c1[l])
            d2.append(d1[l] * c1[l])
            t2.append(t1[l] * c1[l])

            a3.append(a1[l] * d1[l])
            b3.append(b1[l] * d1[l])
            c3.append(c1[l] * d1[l])
            d3.append(d1[l] * d1[l])
            t3.append(t1[l] * d1[l])

            a4.append(a1[l] * t1[l])
            b4.append(b1[l] * t1[l])
            c4.append(c1[l] * t1[l])
            d4.append(d1[l] * t1[l])
            t4.append(t1[l] * t1[l])

        sum_a1 = sum(a1)
        sum_b1 = sum(b1)
        sum_c1 = sum(c1)
        sum_d1 = sum(d1)
        sum_t1 = sum(t1)

        sum_a2 = sum(a2)
        sum_b2 = sum(b2)
        sum_c2 = sum(c2)
        sum_d2 = sum(d2)
        sum_t2 = sum(t2)

        sum_a3 = sum(a3)
        sum_b3 = sum(b3)
        sum_c3 = sum(c3)
        sum_d3 = sum(d3)
        sum_t3 = sum(t3)

        sum_a4 = sum(a4)
        sum_b4 = sum(b4)
        sum_c4 = sum(c4)
        sum_d4 = sum(d4)
        sum_t4 = sum(t4)

        A = np.array([[sum_b1, sum_c1, sum_d1, sum_t1], [sum_b2, sum_c2, sum_d2, sum_t2], [sum_b3, sum_c3, sum_d3, sum_t3], [sum_b4, sum_c4, sum_d4, sum_t4]])
        b = np.array([sum_a1, sum_a2, sum_a3, sum_a4])
        coefficients = np.linalg.solve(A, b)
        return coefficients

    def csv_from_excel(self):
        wb = xlrd.open_workbook('Data.xlsx', encoding_override='utf-8')
        sh = wb.sheet_by_name('Лист1')
        your_csv_file = open('Data.csv', 'w', encoding='utf-8')
        wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))
        your_csv_file.close()

        data = []
        with open('Data.csv', 'r') as inf:
            for line in inf:
                if line.strip():
                    numbers = Decimal(line.replace('"', '').replace('\n', '').replace("'", ''))

                    data.append(float(numbers))
        return data


root = tk.Tk()
app = Application(master=root)
app.mainloop()
