Dim currentRow As Integer

Sub ShowNext20Rows()
    currentRow = currentRow + 20
    ShowData
    PrintData
End Sub

Sub ShowData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")  'write your sheet name    - {point 1}
    
    For i = 1 To 4 ' Assuming you have 4 columns in your excel sheet
        ws.Range(Cells(currentRow, i), Cells(currentRow + 19, i)).Copy Destination:=ws.Cells(1, i + 5) ' Adjust the destination column (i + 5) as needed
    Next i
End Sub

Sub PrintData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")  ' go back to {point - 1}
    
    ' Identify the used range (visible cells) in columns F to I ( as you needed)
    Dim printRange As Range
    On Error Resume Next
    Set printRange = Intersect(ws.UsedRange, ws.Range("F:I")).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not printRange Is Nothing Then
        ' code for printing data
        ws.PageSetup.PrintArea = printRange.Address
        
        ' Print the specified area
        ws.PrintOut
    Else
        MsgBox "No visible data to print."
    End If
End Sub
