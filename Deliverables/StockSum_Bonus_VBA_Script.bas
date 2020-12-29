Attribute VB_Name = "Module2"
Sub StockSummaryBONUS()

' Delcaring varaibles with appropriate data types
Dim ws As Worksheet
Dim i As Integer
Dim x As Integer
Dim y As Integer
Dim largestIncrease As Double
Dim largestDecrease As Double
Dim largestTV As Double
Dim lastrow As Long
Dim dataRange As Range

' Iterate through each seperate worksheet
For Each ws In Worksheets

' Extracting last row of the data set and inserting into VBA variable
lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row

' Set range of primary data table and store in VBA variable
Set dataRange = ws.Range("I1:L" & lastrow)

' Calculate largest % increase
largestIncrease = Application.WorksheetFunction.Max(ws.Range("J:J")) / 100

' Calculate largest % decrease
largestDecrease = Application.WorksheetFunction.Min(ws.Range("J:J")) / 100

' Calculate the greatest total stock value
largestTV = Application.WorksheetFunction.Max(ws.Range("L:L"))

' Create headers of summary table
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Value"

' Format percentage
ws.Cells(2, 17).Value = FormatPercent(largestIncrease, 2)

' Iterate through rows and make necessary calculations
For i = 2 To lastrow

    If (ws.Range("J" & i).Value) / 100 = largestIncrease Then
    
    ws.Cells(2, 16).Value = ws.Range("I" & i)
    
    End If
    
Next i

ws.Cells(3, 17).Value = FormatPercent(largestDecrease, 2)

For x = 2 To lastrow

    If (ws.Range("J" & x).Value) / 100 = largestDecrease Then
    
    ws.Cells(3, 16).Value = ws.Range("I" & x)
    
    End If
    
Next x

ws.Cells(4, 17).Value = largestTV

For y = 2 To lastrow

    If ws.Range("L" & y).Value = largestTV Then
    
    ws.Cells(4, 16).Value = ws.Range("L" & y).Offset(0, -3)
    
    End If
    
Next y

' Move to next worksheet
Next ws

End Sub

