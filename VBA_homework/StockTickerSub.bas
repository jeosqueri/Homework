Attribute VB_Name = "Module1"
Sub StockTicker()

Dim CurrentWs As Worksheet
'Set header variable
Dim Need_Summary_Table_Header As Boolean
'Challenge part: run script one time and have script run on each worksheet
Dim Command_Spreadsheet As Boolean

Need_Summary_Table_Header = False
Command_Spreadsheet = True

'Loop through all worksheets in active workbook
For Each CurrentWs In Worksheets

'    Set initial variable for holding the ticker name
    Dim Ticker_name As String
    Ticker_name = " "
    
'    Set initial variable for holding the total per ticker name
    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0

    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
'    Yearly Price Change
    Dim Yearly_Change As Double
    Yearly_Change = 0
'   Yearly percent change
    Dim Percent_Change As Double
    Percent_Change = 0
'    Set moderate solution variables
'    Set hard solution variables
    
    'Keep track of the  location for each ticker name in the summary table  for the current WS
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

    'Set initial row count for the current worksheet
    Dim Lastrow As Long
    Dim i As Long

    Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        'For all worksheets except the first one, the results
        If Need_Summary_Table_Header Then
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
        Else
'           This  is the first, resulting worksheet
            Need_Summary_Table_Header = True
        End If

'Set initial  value of Open Price for the first Ticker of CurrentWs
'The rest ticker's open price  will be initialized within the for loop below
Open_Price = CurrentWs.Cells(2, 3).Value

For i = 2 To Lastrow

'    Check if we are stilll within the same ticker name, if not write results to summary table
        If CurrentWs.Cells(i, 1).Value <> CurrentWs.Cells(i + 1, 1).Value Then
'            Set the ticker name
            Ticker_name = CurrentWs.Cells(i, 1).Value
            
'            Calculate Yearly Change and Percent Change
            Close_Price = CurrentWs.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
'            Check Division by 0 condition
            If Open_Price <> 0 Then
                Percent_Change = (Yearly_Change / Open_Price) * 100
            Else
                MsgBox ("For " & Ticker_name & "fix issue")
            End If
            
'            Add to Ticker name total volume
            Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            
'            Print the Ticker Name, Yearly Price, Total, and % change in Summary Table
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
'            Reset Yearly_Change and Percent_Change holders for working with new ticker
            Yearly_Change = 0
            Close_Price = 0
'            Capture next Ticker's Open_Price
            Open_Price = CurrentWs.Cells(i + 1, 3).Value
            
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            
    Next i

Next CurrentWs

End Sub
