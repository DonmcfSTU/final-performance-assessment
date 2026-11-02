# Name: Donavan McFadden
# Student ID: Donmcf8357
# Course: SDC205 Python for Data Analytics
# Assignment: 5.5 Final Performance Assessment
# Description: This program adds up five numbers from the user, appends the
# names and annual incomes of five people to final.csv, builds an Excel pie
# chart of all the incomes in final.xlsx with openpyxl, and shows a vertical
# bar graph of the same data with matplotlib.
#
# Save this file in C:\FinalExam with final.csv. The program finds that folder
# from its own location, so it always opens the files by their absolute path.

import csv
import os
from datetime import datetime

import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference

STUDENT_ID = "Donmcf8357"
FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(FOLDER_PATH, "final.csv")
EXCEL_PATH = os.path.join(FOLDER_PATH, "final.xlsx")


# Builds the chart title: my Student ID and today's date, like
# "Donmcf8357 November 2, 2026"
def chartTitle():
    today = datetime.now()
    return STUDENT_ID + " " + today.strftime("%B") + " " + str(today.day) + ", " + str(today.year)


# Reads final.csv and returns two lists: the names and the incomes. The
# incomes are cast to int so they can be charted and used in math.
def readIncomeData():
    names = []
    incomes = []
    with open(CSV_PATH, "r", newline="") as csvFile:
        for row in csv.reader(csvFile):
            if len(row) == 2 and row[0].strip() != "":
                names.append(row[0].strip())
                incomes.append(int(row[1]))
    return names, incomes


# Asks the user for five numbers and displays their total.
def askUser():
    print("Part 1: Total of Five Numbers")
    total = 0

    # This for loop runs five times, once for each number. Each time through,
    # it asks for a number, and the inner while loop keeps asking until the
    # entry is a valid number. Then the number is added to the running total,
    # so after the fifth pass, total holds the sum of all five numbers.
    for count in range(1, 6):
        while True:
            try:
                number = float(input("Enter number " + str(count) + " of 5: "))
                break
            except ValueError:
                print("That is not a number. Please try again.")
        total = total + number

    if total == int(total):
        total = int(total)
    print("The total of your five numbers is " + str(total) + ".")
    print()


# Asks for the names and annual incomes of five people and appends each one
# to the end of final.csv.
def askIncome():
    print("Part 2: Names and Annual Incomes")

    with open(CSV_PATH, "a") as csvFile:
        # This for loop runs five times, once for each person. Each time
        # through, it asks for a name and an annual income, and the inner while
        # loop keeps asking until the income is a whole number. Then it appends
        # "name,income" plus a new line character to final.csv, so each person
        # goes on their own line after the four entries already in the file.
        for person in range(1, 6):
            personName = input("Enter the name of person " + str(person) + " of 5: ")
            while True:
                try:
                    annualIncome = int(input("Enter the annual income for " + personName + ": "))
                    break
                except ValueError:
                    print("Please enter the income as a whole number, like 55000.")
            csvFile.write(personName + "," + str(annualIncome) + "\n")
            print("Saved " + personName + "," + str(annualIncome) + " to final.csv.")

    names, incomes = readIncomeData()
    print("final.csv now has " + str(len(names)) + " entries.")
    print()


# Copies the data from final.csv into a new Excel file called final.xlsx and
# adds a pie chart of the incomes.
def excelPie():
    names, incomes = readIncomeData()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Income Data"
    sheet.append(["Name", "Annual Income"])
    for index in range(len(names)):
        sheet.append([names[index], incomes[index]])
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 16

    # Create a new, empty pie chart object
    pieChart = PieChart()
    # Give the pie chart its title: my Student ID and today's date
    pieChart.title = chartTitle()
    # Point to the income values in column B (rows 2 to the last row) to use as the slices
    incomeValues = Reference(sheet, min_col=2, min_row=2, max_row=len(names) + 1)
    # Point to the names in column A (rows 2 to the last row) to use as the slice labels
    nameLabels = Reference(sheet, min_col=1, min_row=2, max_row=len(names) + 1)
    # Add the income values to the pie chart as its data series
    pieChart.add_data(incomeValues)
    # Set the names as the categories so each slice is labeled with a person's name
    pieChart.set_categories(nameLabels)
    # Make the chart big enough to read all nine slices
    pieChart.width = 16
    pieChart.height = 10
    # Place the pie chart on the worksheet with its top-left corner at cell D2
    sheet.add_chart(pieChart, "D2")

    workbook.save(EXCEL_PATH)
    print("Part 3: The Excel pie chart was saved to " + EXCEL_PATH)
    print()


# Shows a vertical bar graph of the same income data with matplotlib.
def verticalBar():
    names, incomes = readIncomeData()

    plt.figure(figsize=(9, 5))
    plt.bar(names, incomes, color="steelblue")
    plt.title(chartTitle())
    plt.xlabel("Name")
    plt.ylabel("Annual Income ($)")
    plt.tight_layout()
    print("Part 4: Showing the vertical bar graph.")
    plt.show()


def main():
    print("Student ID: " + STUDENT_ID)  # Print my Student ID first
    print()
    askUser()  # Ask for five numbers and show their total
    askIncome()  # Ask for five names and incomes and append them to final.csv
    excelPie()  # Build final.xlsx with the data and a pie chart
    verticalBar()  # Show the vertical bar graph of the same data


main()
