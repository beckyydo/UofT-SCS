import os
import csv

#************************************************PYBANK************************************************

#Define analysis summary function
def analysis(myList):
    average = 0
    greatest_increase = 0
    greatest_decrease = 0
    total = int(myList[0][1])

    for i in range(1,len(myList)):
        #Calculating the total net profit/loss
        total += int(myList[i][1])

        #Calculating change in profit/loss
        compare = int(myList[i][1])-int(myList[i-1][1])

        #Summing the change in profit/loss to calculate average of changes
        average += compare

        #Calculating Greatest Increase and Decrease in Profits/Losses
        if compare > greatest_increase:
            greatest_increase = compare
            date_increase = myList[i][0]

        if compare  < greatest_decrease:
            greatest_decrease = compare 
            date_decrease = myList[i][0]

    #Print Fiancial Analysis Summary
    line1 = "Financial Analysis"
    line2 = "----------------------------"
    line3 = f"Total Months: {len(myList)}"
    line4 = f"Total: ${total}"
    line5 = f"Average Change: ${round(average/(len(myList)-1),2)}"
    line6 = f"Greatest Increase in Profits: {date_increase} (${round(greatest_increase,2)})"
    line7 =f"Greatest Decrease in Profits: {date_decrease} (${round(greatest_decrease,2)})"
    summary = (line1, line2, line3, line4, line5, line6, line7)

    #Join element in list with a new line
    summary = "\n".join(map(str,summary))
    return summary

#Declare path to open file and create output file
filepath = os.path.join("Resources", "budget_data.csv")
output_path = os.path.join("analysis", "financial_analysis.txt")

budget = []
with open(filepath, 'r') as input_file:
    csv_reader = csv.reader(input_file)
    csv_header = next(input_file)
    
    #Open file and apply analysis function to file
    for rows in csv_reader:
        budget.append(rows)
    summary = analysis(budget)

    #Create text file of analysis summary and print in console
    with open(output_path,'w') as output_file:
        output_file.write(summary)
        print(summary)

#************************************************PYPOLL************************************************
#Define function for election summary
def election_summary(myList):
    myList.sort()
    tally = [[myList[0],1]]
    k=0
    
    #Count the number of voters in the list
    for i in range(1,len(myList)):
        if myList[i] == myList[i-1]:
            tally[k][1] += 1
        
        #Append candidate if not already in the list
        else:
            tally.append([myList[i],1])
            k += 1

    #The beginning of election results summary
    line1 = "Election Results"
    dash_line = "-------------------------"
    line2 = f"Total: {len(myList)}"
    summary = list((line1, dash_line, line2))
    
    #Loop and return a list of the candidate with percentage and # of votes
    winner_total=tally[0][1]
    winner= tally[0][0]
    for i in range(len(tally)):
        summary.append(f"{tally[i][0]}: {round(tally[i][1]/len(myList)*100,3)}% ({tally[i][1]})")

        #Find the candidate with the highest votes
        if winner_total <= tally[i][1]:
            winner_total = tally[i][1]
            winner = tally[i][0]

    #Append the remaining election summary with the winning candidate
    summary.append(dash_line)
    summary.append(f"Winner: {winner}") 
    summary.append(dash_line)

    #Join element in list with a new line
    summary = "\n".join(map(str,summary))   
    return summary

#Declare path to open file and create output file
filepath= os.path.join("Resources","election_data.csv")
output_path = os.path.join("analysis","election_results.txt")
#Open File
election = []
with open(filepath, 'r') as input_file:
    csv_reader = csv.reader(input_file)
    csv_header = next(input_file)
    
    #Create a list of the candidate
    for rows in csv_reader:
        election.append(rows[2])
    vote = election_summary(election)

    #Create text file of election summary and print in console
    with open(output_path, 'w') as output_file:
        output_file.write(vote)
        print(vote)
    

