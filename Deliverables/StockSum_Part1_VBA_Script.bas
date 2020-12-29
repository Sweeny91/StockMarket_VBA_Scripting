Attribute VB_Name = "Module1"
Sub StockSummary()

' Declaring variables w/ appropriate data types
Dim i As Long
Dim a As Long
Dim ticker As String
Dim last As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim volume As Double
Dim rangeCount As Long
Dim lastrow As Long
Dim dataTable As Long
Dim lastDataTableRow As Long
Dim ws As Worksheet

' Iterate through each worksheet in the workbook
For Each ws In Worksheets

' Extracting last row of the data set and inserting into VBA variable
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Developing placement and headers for summary data
dataTable = 2
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

 ' Iterate through all stock data on ONE sheet
 For i = 2 To lastrow
 
    ' Create counter to be used to calculate length of ticker range
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
    
        rangeCount = rangeCount + 1

        ' Calculate increaing stock volume for current stock
        volume = volume + ws.Cells(i, 7).Value
        
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Print ticker value in appropriate cell
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & dataTable).Value = ticker

        ' Variables to hold value for beginning and end of current ticker range
        Start = ws.Cells(i - rangeCount, 3).Value
        last = ws.Cells(i, 6).Value

        If Start <> 0 Then
        
        ' List yearly change from open to close
        yearlyChange = last - Start
        
        ' List percent change from open to close
        percentChange = yearlyChange / last
        
        Else
        
        percentChange = 0
        yearlyChange = 0
        
        End If
        
        ' Place correctly formatter annual change in appropriate cell
        ws.Range("J" & dataTable).Value = yearlyChange
        ws.Range("K" & dataTable).Value = FormatPercent(percentChange, 2)
    
        ' Sum total stock volume per ticker
        volume = volume + ws.Cells(i, 7).Value
        ws.Range("L" & dataTable).Value = volume

        ' Reset values
        rangeCount = 0
        volume = 0

        dataTable = dataTable + 1

    End If

Next i

lastDataTableRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

For a = 2 To lastDataTableRow

    ' Conditional to change cell color for positive or negative change
    If ws.Cells(a, 10).Value > 0 Then
    
        ' Color Code
        ws.Cells(a, 10).Interior.ColorIndex = 4

    ElseIf ws.Cells(a, 10).Value < 0 Then
    
        ' Alternate color Code
        ws.Cells(a, 10).Interior.ColorIndex = 3

    End If

Next a

' Move to next worksheet
Next ws

End Sub

