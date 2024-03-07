# VBA-challenge

Description:
  This is the module 2 Homework Challange for VBA Scripting

Getting Started:
  This challenge uses the Starter Code file 'Multiple_year_stock_data'.xlsx excel spreadsheet

Installing  
  This output is produced from a Windows 11 environment running Excel in Windows 365
  Developer mode is activated in Excel to allow for VBA modules to be created and saved 
  The final file is therefore the 'Multiple_year_stock_data'.xlsm excel spreadsheet with macro included

Executing program
  Instructions 1 'Get list of the Ticker Symbols' and 4 'Total Stock Volume' have been executed using the following VBA code:

Sub VBA_Challenge_Stock_data()

    'Set column headings for data to be input
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Value"

    'Set column widths to fit new headings
    Columns("I:L").AutoFit
    'Reference: https://analysistabs.com/excel-vba/change-row-height-column-width/
    
        'Set variable to store the ticker code
        Dim Ticker As String
    
       'Set variable to store the stock volume
        Dim stock_volume As Double
        stock_volume = 0

        'Set location for ticker code in column I summary
        Dim Ticker_Summary_Row As Integer
        Ticker_Summary_Row = 2
        
        'Set row counter for loop to continue until the last row with data in column A
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
                       
        'Create loop to loop through Column A <ticker> to get unique list of tickers and total the stock volume per ticker from Column G <vol>
        For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        stock_volume = stock_volume + Cells(i, 7).Value
                
        Else
        
        'Set the ticker code
         Ticker = Cells(i, 1).Value
        
        'Sum the ticker stock volume
         stock_volume = stock_volume + Cells(i, 7).Value
        
        'Print Ticker code in Column I Ticker
        Range("I" & Ticker_Summary_Row).Value = Ticker
        
        'Print Stock Volume to Ticker in Column L Total Stock Volume
        Range("L" & Ticker_Summary_Row).Value = stock_volume
        
        'Add 1 to move to the next row for the Ticker_Summary_Row
         Ticker_Summary_Row = Ticker_Summary_Row + 1
        
        'Reset the stock_volume
         stock_volume = 0
        
    End If
    
  Next i

End Sub


This produces the following output on the 2018 tab of the excel spreadsheet ''Multiple_year_stock_data'.xlsm'
which appears as follows:
![image](https://github.com/ElizabethDashwood/VBA-challenge/assets/160380658/16fea123-9213-49c3-9339-06874e9d22ba)

Help:
Only Steps 1 and 4 of this challenge have been completed to date. 
A later re-submission for this module will include all completed steps
