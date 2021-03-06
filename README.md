# Dinah---Stocks-Analysis

2nd Project on Boot Camp Exercise for Data Analystics


## Purpose
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

### Overview of Project: 
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.


Deliverable 1: Refactor VBA Code and Measure Performance (80 points)
#### The VBA Code:

Sub AllStocksAnalysis_prev()
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   
   
    Dim startTime As Single
    Dim endTime  As Single
    Dim startingPrice As Single
    Dim endingPrice As Single
  
  
   'Formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = x1Continuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    yearValue = InputBox("What year would you like to run the analysis on?")
   
   'initiate timer
    startTime = Timer
   
   'Range("A1").Value = "All Stocks (2018)"
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


   '2) Initialize array of all tickers
   Dim tickers(12) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   
   '3a) Initialize variables for starting price and ending price
   startingPrice = 0
   endingPrice = 0
   
   
   '3b) Activate data worksheet
    Worksheets(yearValue).Activate
   
   
   '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row


   '4) Loop through tickers
    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
           
       Next j
       
       
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

    dataRowStart = 4
    dataRowEnd = 15
    
    
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

---------  End of The VBA Code


#### Deliverable 1 Instructions
Use your knowledge of VBA and the starter code provided in this Challenge to refactor the Module2_VBA_Script so you loop through the data one time and collect all of the information. Your refactored code should run faster than it did in this module.

If you open the file: https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge.xlsm
There are buttons that will give you the speed of the code in:
  Results for running 2017 data is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2017.png
  Results for running 2018 data is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2018.png

Download the challenge_starter_code.vbs file and rename it VBA_Challenge.vbs.
Create a folder called “Resources” to hold the run-time pop-up messages that you’ll screenshot after running refactored analyses for 2017 and 2018.
Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm.
Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
Use the steps below to add code where indicated by the numbered comments in the starter code file.
Step 1a:

Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
Step 1b:

Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
Step 2a:

Create a for loop to initialize the tickerVolumes to zero.
Step 2b:

Create a for loop that will loop over all the rows in the spreadsheet.
Step 3a:

Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.
If you’d like a hint on how to increase the current tickerVolumes by using the tickerIndex variable as the index, that’s totally okay. If not, that’s great too. You can always revisit this later if you change your mind.


Step 3b:

Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices variable.
Step 3c:

Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
Step 3d:

Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
Step 4:

Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
Finally, run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module (as shown in the images below). In your Resources folder, save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook.

  Results for running 2017 data is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2017.png
  Results for running 2018 data is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2018.png



#### Deliverable 1 Requirements
You will earn a perfect score for Deliverable 1 by completing all requirements below:

The tickerIndex is set equal to zero before looping over the rows. (5 pt).
Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (15 pt).
The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays (15 pt).
The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (25 pt).
Code for formatting the cells in the spreadsheet is working (5 pt).
There are comments to explain the purpose of the code (5 pt).
The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module (5 pt).
The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png (5 pt).


### Deliverable 2: Written Analysis of Results

#### Deliverable 2 Instructions
Initialize your repository with a README, and write your analysis of Deliverable 1.

NOTE
This documentation (Links to an external site.) provides more information about basic writing and formatting syntax for GitHub.
The document is at https://github.com/Dybondzy/stocks-analysis/edit/main/README.md

For your written analysis, be sure to use complete and coherent sentences. Your written analysis should contain three sections, which cover the following:

Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


#### Deliverable 2 Requirements
Structure, Organization, and Formatting Requirements (8 points)
The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).
Analysis Requirements (12 points)
The written analysis has the following:

Overview of Project
The purpose and background are well defined (2 pt).
Results
The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
Submission
Once you’re ready to submit, make sure to check your work against the rubric to ensure you are meeting the requirements for this Challenge one final time. It’s easy to overlook items when you’re in the zone!

As a reminder, upload the following deliverables to your stock-analysis GitHub repository:

Deliverable 1: Refactor VBA code and measure performance
The VBA_Challenge.xlsm workbook
The Resources folder with VBA_Challenge_2017.png and VBA_Challenge_2018.png
Deliverable 2: A written analysis of your results (README.md)
To submit your challenge assignment, click Submit, then provide the URL of your stock-analysis GitHub repository for grading.


This new assignment consists of one technical deliverable and a written report to deliver your results. You will submit the following:

#### Deliverable 1: Refactor VBA code and measure performance
  1.  This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time
  2.  The repository stock_analysis holds the updated workbook: StockAnalysis, and the screen prints in 
  3.  The location for the repository is https://github.com/Dybondzy/stocks-analysis
  4.  The location of the worksheet is https://github.com/Dybondzy/stocks-analysis/blob/main/AllStocksAnalysis.xlsm
  5.  The location of the worksheet is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge.xlsm
  6.  The location of the screen prints (png) is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2017.png
  7.  The location of the screen prints (png) is https://github.com/Dybondzy/stocks-analysis/blob/main/VBA_Challenge_2018.png


#### Deliverable 2: A written analysis of your results (README.md)
  1.  This README file is the written analysis
  2.  The location is https://github.com/Dybondzy/stocks-analysis/edit/main/README.md
  

## Summary

Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
" Maintainability: After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. 
Disadvantages of Code Refactoring: Time Consuming: You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go"

