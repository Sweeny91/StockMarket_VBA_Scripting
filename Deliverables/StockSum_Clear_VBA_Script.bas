Attribute VB_Name = "Module3"
' Function used to clear contents of summary data table.
' Necessary to make code run faster in order to check to see if my code is working properly.


Sub Clear()
Dim ws As Worksheet

For Each ws In Worksheets
    
    ws.Range("I:L").ClearContents
    ws.Range("I:L").ClearFormats
    ws.Range("O:Q").ClearContents
    ws.Range("O:Q").ClearFormats

Next ws

End Sub

