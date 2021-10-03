# PyBank

# import dependencies for reading and importing excel file
import csv

# import and read csv file
file = '/Users/kristenhanold/Desktop/GaTechBootCamp/budget_data.csv'

# create empty lists to iterate through and calculate total months, total profits and total avg changes
totalNumMonths = []
totalProfits = []
totalChanges = []

# open csv file
with open(file) as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")

    # skip headers
    header = next(csvreader)

    # Looping through the total number of months included in the dataset
    # Also looping through the net total amount of "Profit/Losses" over the entire period
    # added to corresponding lists created above
    for row in csvreader:
        totalNumMonths.append(row[0])
        totalProfits.append(int(row[1]))

    # Calculate the changes in "Profit/Losses" over the entire period, then find the average of those changes
    for x in range(len(totalProfits) - 1):
        totalChanges.append(totalProfits[x + 1] - totalProfits[x])

    # The greatest increase in profits (date and amount) over the entire period
    # The greatest decrease in profits (date and amount) over the entire period
    max_increase = max(totalChanges)
    max_decrease = min(totalChanges)
    max_increase_month = totalChanges.index(max(totalChanges)) + 1
    max_decrease_month = totalChanges.index(min(totalChanges)) + 1

# print statements
print('Financial Analysis')
print('------------------------------')
print('Total Months: ', len(totalNumMonths))
print('Total: ', '$', sum(totalProfits))
print('Total Average Change: ', round(sum(totalChanges) / len(totalChanges), 2))
print('Greatest Increase in Profits: ', totalNumMonths[max_increase_month], '(', '$', max_increase, ')')
print('Greatest Decrease in Profits: ', totalNumMonths[max_decrease_month], '(', '$', max_decrease, ')')

# import dependencies for output
import xlsxwriter

# create and file (workbook) and worksheet
workbook = xlsxwriter.Workbook('PythonHW_PyBank.xlsx')
sheet1 = workbook.add_worksheet('Sheet 1')

# declare data and assign to a cell
sheet1.write('A1', "Financial Analysis")
sheet1.write('A2', "----------------------------")
sheet1.write('A3', f"Total Months: {len(totalNumMonths)}")
sheet1.write('A4', f"Total: ${sum(totalProfits)}")
sheet1.write('A5', f"Average Change: ${round(sum(totalChanges) / len(totalChanges), 2)}")
sheet1.write('A6', f"Greatest Increase in Profits: {totalNumMonths[max_increase_month]}")
sheet1.write('B6', f"${(str(max_increase))})")
sheet1.write('A7', f"Greatest Decrease in Profits: {totalNumMonths[max_decrease_month]}")
sheet1.write('B7', f"${(str(max_decrease))})")

workbook.close()
