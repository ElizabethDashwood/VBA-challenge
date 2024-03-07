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

<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>		Ticker	Yearly Change	Percentage Change	Total Stock Value
AAB	20180102	24.44	24.56	24.44	24.47	261879		AAB			765628638
AAB	20180103	24.45	24.45	24.22	24.28	15721045		AAF			2348251513
AAB	20180104	24.27	24.36	24.27	24.28	5954		AAR			44163252
AAB	20180105	24.24	24.33	24.22	24.33	58161100		AAT			1121804
AAB	20180108	24.28	24.72	24.28	24.72	267347		ABJ			866724156
AAB	20180109	24.73	24.73	24.64	24.64	1129348		ABK			10391386
AAB	20180110	24.62	24.78	24.57	24.68	5453423		ABKV			5324614313
AAB	20180111	24.68	24.85	24.68	24.85	14391		ABM			1029560425
AAB	20180112	24.86	24.86	24.71	24.86	13701		ACJ			6169423
AAB	20180116	24.85	24.93	24.79	24.79	101867		ACYQ			234097825
AAB	20180117	24.81	24.85	24.68	24.69	515155		ADB			41980152
AAB	20180118	24.67	24.86	24.67	24.86	657377		ADF			13907867898
AAB	20180119	24.89	24.93	24.71	24.71	1106987		AEL			4339867319
AAB	20180122	24.66	24.9	24.66	24.85	1418811		AEV			843970
AAB	20180123	24.87	24.94	24.8	24.84	52744		AEY			8274367
AAB	20180124	24.83	24.83	24.56	24.56	3545845		AFZ			175801284
AAB	20180125	24.58	24.66	24.54	24.66	3119820		AGF			1903886691
![image](https://github.com/ElizabethDashwood/VBA-challenge/assets/160380658/16fea123-9213-49c3-9339-06874e9d22ba)




Help:
Only Steps 1 and 4 of this challenge have been completed to date. 
A later submission will include all completed steps
