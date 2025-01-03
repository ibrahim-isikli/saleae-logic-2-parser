Sub CombineDataIntoRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim outputRow As Long
    Dim threshold As Double
    Dim startTimeDiff As Double
    Dim currentData As String
    Dim outputCol As Long
    
    ' Ayarlar
    Set ws = ThisWorkbook.Sheets(1) ' Çalışma sayfası
    threshold = 0.08417675 ' Eşik değer
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' C sütunundaki son dolu satır
    outputRow = 2 ' Sonuçların başlayacağı satır (düzenleyebilirsiniz)
    outputCol = 8 ' Sonuçların yazılacağı sütun (H sütunu)

    ' Verileri birleştir
    currentData = ""
    For currentRow = 2 To lastRow ' 2. satırdan başla
        If currentRow > 2 Then
            startTimeDiff = ws.Cells(currentRow, "C").Value - ws.Cells(currentRow - 1, "C").Value
            If startTimeDiff > threshold Then
                ' Mevcut veri satırını yazdır
                ws.Cells(outputRow, outputCol).Value = Trim(currentData)
                outputRow = outputRow + 1 ' Bir alt satıra geç
                currentData = "" ' Yeni veri paketi için sıfırla
            End If
        End If
        currentData = currentData & " " & ws.Cells(currentRow, "E").Value
    Next currentRow
    
    ' Son kalan veriyi yazdır
    If currentData <> "" Then
        ws.Cells(outputRow, outputCol).Value = Trim(currentData)
    End If

    MsgBox "Veriler birleştirildi!", vbInformation
End Sub
