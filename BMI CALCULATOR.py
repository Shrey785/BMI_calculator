#creating a bmi calculator using tkinter

import tkinter as tk
from tkinter import*
from tkinter import messagebox
import matplotlib.pyplot as plt
import numpy as np
import openpyxl #using openpyxl to connect exel sheets with the program as a database

def validateLogin(bmi):
    excel_file = "bmi.xlsx"
    height, age, weight, name = collect_user_data()
    update_excel_sheet(excel_file,height,age,weight,bmi,name)
    return True

def collect_user_data():#using this to collect all the data user enterd to store it in database
    height =  height_entry.get()
    age = Age_entry.get()
    weight = weight_entry.get()
    name = Name_entry.get()
    
    return height,age,weight,name

def update_excel_sheet(file_path, height,age,weight,bmi_index, name):#used to update every time another user enters
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        next_row = sheet.max_row + 1
        if sheet.max_row == 1:
            sheet.cell(row=1, column=1, value="Name")
            sheet.cell(row=1, column=2, value="Age")
            sheet.cell(row=1, column=3, value="Height")
            sheet.cell(row=1, column=4, value="Weight")
            sheet.cell(row=1, column=5, value="Bmi_index")

        sheet.cell(row=next_row, column=1, value=name)
        sheet.cell(row=next_row, column=2, value=age)
        sheet.cell(row=next_row, column=3, value=height)
        sheet.cell(row=next_row, column=4, value=weight)
        sheet.cell(row=next_row, column=5, value=bmi_index)

        workbook.save(file_path)
        print("Data saved successfully!")

    except Exception as e:
        print(f"Error: {e}")

ws = Tk()
ws.title('BMI_CALCULATOR')
ws.geometry('750x650')
ws.config(bg='green')

#the class calc_bmi is used to calculate the bmi of the user
def calc_bmi():
    try:
        height = float((height_entry.get()))/100
        weight = float((weight_entry.get()))
        bmi = float(weight / (height**2))
        bmi = round(bmi, 1)
        bmi_index(bmi)
        validateLogin(bmi)
        plot_graph(weight, height, bmi)

    except ValueError:
        result_label.config("Pls enter valid weight anf height")

def bmi_index(bmi):
    if bmi < 18.5:
        result_label.config(text=f"Your BMI is: {bmi:.2f}, is Underweight")
        # messagebox.show(f'BMI = {bmi} is Underweight')
    elif (bmi > 18.5) and (bmi < 24.9):
        result_label.config(text=f"Your BMI is: {bmi:.2f}, is Normal")
        # messagebox.showinfo(f'BMI = {bmi} is Normal')
    elif (bmi > 24.9) and (bmi < 29.9):
        result_label.config(text=f"Your BMI is: {bmi:.2f}, is overweight")
        # messagebox.showinfo(f'BMI = {bmi} is Overweight')
    else:
        result_label.config(text=f"Your BMI is: {bmi:.2f}, is Obesity")
        # messagebox.showinfo('Obesity')  
var = IntVar()
# creating the name section to get name
Name_label = Label(text='Name', font=('Arial', 10, 'bold')).pack(pady=0)
Name_entry = tk.Entry(ws)
Name_entry.pack(pady=2)

# creating the age section to get age
Age_label = Label(text='Age', font=('Arial', 10, 'bold')).pack(pady=0)
Age_entry = tk.Entry(ws)
Age_entry.pack(pady=2)

# creating the weight section to get weight
tk.Label(ws, text="Weight (kg):", font=('Arial', 10, 'bold')).pack(pady=0)
weight_entry = tk.Entry(ws)
weight_entry.pack(pady=2)

# creating the hight section to get hight
tk.Label(ws, text="Height (cm):",font=('Arial', 10, 'bold')).pack(pady=0)
height_entry = tk.Entry(ws)
height_entry.pack(pady=2)

calculate_button = tk.Button(ws, text="Calculate BMI", command=calc_bmi)
calculate_button.pack(pady=10)

result_label = tk.Label(ws, text="")
result_label.pack(pady=10)

#now using the data , ploting the graph

def plot_graph(weight, height, bmi):
    weight = np.array([20, 40, 60, 80, 100, 120, 140, 160, 180, 200])

    height = np.array([150, 155, 160, 165, 170, 175, 180, 185, 190.195, 200])
    print(weight, height, bmi)
    plt.plot(weight, height, bmi)

    plt.xlabel('weight')
    plt.ylabel('height')

    plt.title("BMI GRAPH")

    plt.show()
ws.mainloop()