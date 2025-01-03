
'This macro inserts blank rows in an Excel worksheet based on the time difference between consecutive values in column C. 
'It compares the time difference between each row and the previous one, and if the difference exceeds a specified threshold, a blank row is inserted. 
'The function operates in reverse order, starting from the last data row to avoid disrupting the data during insertion.

Sub InsertBlankRowsBasedOnTimeDifference()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim timeDiff As Double
    Dim threshold As Double
    
    ' Settings
    Set ws = ThisWorkbook.Sheets(1) ' The sheet containing your data
    threshold = 0.005    ' The specified time difference 
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' Last filled row in column C

    ' Adding blank rows by checking in reverse order
    For i = lastRow To 3 Step -1 ' Starting from row 3 to skip header rows
        If IsNumeric(ws.Cells(i, "C").Value) And IsNumeric(ws.Cells(i - 1, "C").Value) Then
            timeDiff = ws.Cells(i, "C").Value - ws.Cells(i - 1, "C").Value
            If timeDiff > threshold Then
                ws.Rows(i).Insert Shift:=xlDown
            End If
        End If
    Next i

    MsgBox "Process completed!", vbInformation
End Sub
