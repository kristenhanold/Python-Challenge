# import dependencies for reading and importing excel file
import csv

file = '/Users/kristenhanold/Desktop/GaTechBootCamp/election_data.csv'

# setting variables to 0 as a starting point so that we are able
# to increment to the amount for total votes and votes per candidate
totalVotes = 0
khanVotes = 0
correyVotes = 0
liVotes = 0
otooleyVotes = 0

# open csv file
with open(file) as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")

    # skip headers
    header = next(csvreader)

    # for every row in our csv file, we want to increment by 1 row
    # for each loop and count each row with a vote, and add to variable "totalVotes"
    for row in csvreader:
        totalVotes += 1

        # for each candidate, we want to count how many times their name appears
        if row[2] == 'Khan':
            khanVotes += 1
        if row[2] == 'Correy':
            correyVotes += 1
        if row[2] == 'Li':
            liVotes += 1
        if row[2] == "O'Tooley":
            otooleyVotes += 1

    # based on total votes and each candidate's num of votes, we want to calculate the percentage of each candidate
    # represented in total amount of votes
    khan_percentage = round(((khanVotes / totalVotes) * 100), 2)
    correy_percentage = round(((correyVotes / totalVotes) * 100), 2)
    li_percentage = round(((liVotes / totalVotes) * 100), 2)
    otooley_percentage = round(((otooleyVotes / totalVotes) * 100), 2)

    # creating 2 lists to be used later in our print statement, to determine which candidate had the max votes
    candidates = ['Khan', 'Correy', 'Li', "O'Tooley"]
    votes = [khanVotes, correyVotes, liVotes, otooleyVotes]
    indexOfWinner = votes.index(max(votes))

    print('Election Results')
    print('---------------------------')
    print('Total Votes: ', totalVotes)
    print('---------------------------')
    print('Khan: {:.3f}%'.format(khan_percentage), '(', khanVotes, ')')
    print('Correy: {:.3f}%'.format(correy_percentage), '(', correyVotes, ')')
    print('Li: {:.3f}%'.format(li_percentage), '(', liVotes, ')')
    print("O'Tooley: {:.3f}%".format(otooley_percentage), '(', otooleyVotes, ')')
    print('Winner: ', candidates[indexOfWinner])

# import dependencies for output
from xlsxwriter import Workbook

# create and file (workbook) and worksheet
pypoll_wb = Workbook('PythonHW_PyPoll.xlsx')
sheet1 = pypoll_wb.add_worksheet('Sheet 1')

# declare data and assign to a cell
sheet1.write('A1', "Election Results")
sheet1.write('A2', "----------------------------")
sheet1.write('A3', f"Total Votes: {totalVotes}")
sheet1.write('A4', "----------------------------")
sheet1.write('A5', f"Khan: {khan_percentage:.3f}%")
sheet1.write('B5', f'{khanVotes}')
sheet1.write('A6', f"Correy: {correy_percentage:.3f}%")
sheet1.write('B6', f'{correyVotes}')
sheet1.write('A7', f"Li: {li_percentage:.3f}%")
sheet1.write('B7', f'{liVotes}')
sheet1.write('A8', f"O'Tooley: {otooley_percentage:.3f}%")
sheet1.write('B8', f'{otooleyVotes}')
sheet1.write('A9', "----------------------------")
sheet1.write('A10', f'Winner: {candidates[indexOfWinner]}')
sheet1.write('A11', "----------------------------")

pypoll_wb.close()
