Attribute VB_Name = "Module1"
Sub Stock()

' Loop through and Activate all Sheets
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
    ' Create Heading for summary table (before defining variables is common practice according to internet).
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"

    ' Define Variables
    Dim LastRow As Long
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim j As Long
    Dim k As Long
        
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume As Double
    Dim Row As Double
    Dim Column As Integer
    Dim YearlyChangeLastRow As Long
    
    
    ' Assigning initial value to variables
    Row = 2
    Column = 1
    Volume = 0

    ' Acknowledge Initial Open Price (which is given to us in data)
    Open_Price = Cells(2, Column + 2).Value

    ' Loop through all ticker symbol
    For i = 2 To LastRow
        
    ' Check if moved onto new ticker symbol
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
            'Set Ticker Cell  (which is given to us in data) & add ticker data to summary table
            Ticker = Cells(i, Column).Value
            Cells(Row, Column + 8).Value = Ticker
                
            'Set Close Price  (which is given to us in data)
            Close_Price = Cells(i, Column + 5).Value
                
            'Set Yearly Change & add Yearly Change to summary table
            Yearly_Change = Close_Price - Open_Price
            Cells(Row, Column + 9).Value = Yearly_Change
        
            'Set Percent Change condition & set Cell format to % & add Percent Change to summary table
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
            
            'Add Total Volume & add Volume to sumarry table
            Volume = Volume + Cells(i, Column + 6).Value
            Cells(Row, Column + 11).Value = Volume
            
                'Increment the row for our new data in summary table
                Row = Row + 1
            
            'Continue to next stock with reset of Open Price
            Open_Price = Cells(i + 1, Column + 2)
        
            'Reset Total Volume for next stock
            Volume = 0
            ' If cells are the same ticker symbol, then...
        Else
            Volume = Volume + Cells(i, Column + 6).Value
        End If
    Next i
    
    'Determine Last Row of Yearly Change
    YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
           
    'Set Cell Colors condition on summary table
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
    ' Create new summary table (next to first one)
    ' Set Greatest % Increase, Greatest % Decrease, & Greatest Total Volume
    Cells(1, Column + 16).Value = "Value"
    Cells(1, Column + 15).Value = "Ticker"
    Cells(2, Column + 14).Value = "Greatest % Increase"
    Cells(3, Column + 14).Value = "Greatest % Decrease"
    Cells(4, Column + 14).Value = "Greatest Total Volume"
    
    ' Create Loop to search through for greatest % increase, greatest % decrease, and greatest total volume
    For k = 2 To YearlyChangeLastRow
        If Cells(k, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
            Cells(2, Column + 15).Value = Cells(k, Column + 8).Value
            Cells(2, Column + 16).Value = Cells(k, Column + 10).Value
            Cells(2, Column + 16).NumberFormat = "0.00%"
        ElseIf Cells(k, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
            Cells(3, Column + 15).Value = Cells(k, Column + 8).Value
            Cells(3, Column + 16).Value = Cells(k, Column + 10).Value
            Cells(3, Column + 16).NumberFormat = "0.00%"
        ElseIf Cells(k, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
            Cells(4, Column + 15).Value = Cells(k, Column + 8).Value
            Cells(4, Column + 16).Value = Cells(k, Column + 11).Value
        End If
    Next k
    
  
Next WS

End Sub
