Attribute VB_Name = "Module1"
Sub StockTickerData()

Dim CurrentWs As Worksheet

'Loop through all worksheets in the workbook
For Each CurrentWs In Worksheets

'    Set variable for holding the ticker name
    Dim Ticker_name As String
    Ticker_name = " "
    
'    Set variable for holding the total volume per ticker name
    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0

'    Set variables for stock open price, close price, yearly price change, and yearly percent change
    Dim Open_Price As Double
    Open_Price = 0
    
    Dim Close_Price As Double
    Close_Price = 0

    Dim Yearly_Change As Double
    Yearly_Change = 0

    Dim Percent_Change As Double
    Percent_Change = 0

'    Hard solution variables for greatest % increase, greatest % decrease, and greatest total volume
'    Greatest % Increase
    Dim Max_Ticker_Name As String
    Max_Ticker_Name = " "
    Dim Max_Percent_Increase As Double
    Max_Percent_Increase = 0
    
'    Greatest % Decrease
    Dim Min_Ticker_Name As String
    Min_Ticker_Name = " "
    Dim Max_Percent_Decrease As Double
    Max_Percent_Decrease = 0
    
'    Greatest Total Volume
    Dim Max_Vol_Ticker As String
    Max_Vol_Ticker = " "
    Dim Max_Total_Volume As Double
    Max_Total_Volume = 0

    
    'Define summary table and location for each current WS
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    'Set initial row count for the current worksheet and define last row
    Dim Lastrow As Long
    Dim i As Long

        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
    
        'For all worksheets set summary table headers

            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"

    'Set initial value of open price for the first ticker of current WS (for the rest of the tickers, the open price will be initialize in the for loop

    Open_Price = CurrentWs.Cells(2, 3).Value

For i = 2 To Lastrow

'    Check if in the same ticker name, if not write results to summary table
        If CurrentWs.Cells(i, 1).Value <> CurrentWs.Cells(i + 1, 1).Value Then
'            Insert ticker name
            Ticker_name = CurrentWs.Cells(i, 1).Value
            
'            Calculate Yearly Change
            Close_Price = CurrentWs.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
'            Calculate percent change
            If Open_Price <> 0 Then
                Percent_Change = (Yearly_Change / Open_Price) * 100
'            Else
'                MsgBox ("For " & Ticker_name & "fix issue")
            End If
            
'            Add to ticker name total volume
            Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            
'            Print the ticker Name, Yearly Price, Total, and % change in Summary Table
            CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_name
            CurrentWs.Range("J" & Summary_Table_Row).Value = Yearly_Change
                If (Yearly_Change > 0) Then
'                    If yearly change value is positive, fill cell green
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
'                    If yearly change value is negative, fill cell red
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
        
            CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
            CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
            
'            Add 1 to summary table row count
            Summary_Table_Row = Summary_Table_Row + 1
'            Reset Yearly_Change for working with new ticker
            Yearly_Change = 0
            Close_Price = 0
'            Get next Ticker's open price
            Open_Price = CurrentWs.Cells(i + 1, 3).Value
            
'            Hard solution code to populate new summary table
            If (Percent_Change > Max_Percent_Increase) Then
                Max_Percent_Increase = Percent_Change
                Max_Ticker_Name = Ticker_name
            ElseIf (Percent_Change < Max_Percent_Decrease) Then
                Max_Percent_Decrease = Percent_Change
                Min_Ticker_Name = Ticker_name
            End If
            
            If (Total_Ticker_Volume > Max_Total_Volume) Then
                Max_Total_Volume = Total_Ticker_Volume
                Max_Vol_Ticker = Ticker_name
            End If
            
'            Reset percent change and ticker volume for hard solution
            Percent_Change = 0
            Total_Ticker_Volume = 0
            
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            
    Next i
    
'        Hard solution: enter greatest % increase, greatest % decrease, and greatest total volume into new summary table ranges
'
                CurrentWs.Range("O2").Value = "Greatest % Increase"
                CurrentWs.Range("P2").Value = Max_Ticker_Name
                CurrentWs.Range("Q2").Value = (CStr(Max_Percent_Increase) & "%")
                
                CurrentWs.Range("O3").Value = "Greatest % Decrease"
                CurrentWs.Range("P3").Value = Min_Ticker_Name
                CurrentWs.Range("Q3").Value = (CStr(Max_Percent_Decrease) & "%")
                
                CurrentWs.Range("O4").Value = "Greatest Total Volume"
                CurrentWs.Range("P4").Value = Max_Vol_Ticker
                CurrentWs.Range("Q4").Value = Max_Total_Volume
                
                CurrentWs.Range("P1").Value = "Ticker"
                CurrentWs.Range("Q1").Value = "Value"
'
                
        
Next CurrentWs

End Sub
