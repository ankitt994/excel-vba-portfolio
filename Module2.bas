Attribute VB_Name = "Module2"
Sub CleanMessyData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim priceVal As String, qtyVal As String
    Dim nameVal As String
    
    Set ws = ThisWorkbook.Sheets("Data_Cleaning")
    
    ' Find last row in messy table (Column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row of messy data
    For i = 7 To lastRow
        ' Product Name cleanup
        nameVal = CStr(ws.Cells(i, 1).Text)
        nameVal = Application.WorksheetFunction.Trim(nameVal) ' Remove extra spaces
        nameVal = ProperCaseText(nameVal)                      ' Proper case
        ws.Cells(i, 1).Value = nameVal
        
        ' Price cleanup
        ws.Cells(i, 2).NumberFormat = "General" ' Remove formatting (? etc.)
        priceVal = CStr(ws.Cells(i, 2).Text)    ' Read what is visible
        priceVal = ReplaceSymbols(priceVal)     ' Remove currency & units
        If IsNumeric(priceVal) Then
            ws.Cells(i, 2).Value = CDbl(priceVal)
        Else
            ws.Cells(i, 2).Value = 0
        End If
        
        ' Quantity cleanup
        ws.Cells(i, 3).NumberFormat = "General"
        qtyVal = CStr(ws.Cells(i, 3).Text)
        qtyVal = ReplaceSymbols(qtyVal)
        If qtyVal = "" Or UCase(qtyVal) = "N/A" Then qtyVal = "0"
        If IsNumeric(qtyVal) Then
            ws.Cells(i, 3).Value = CLng(qtyVal)
        Else
            ws.Cells(i, 3).Value = 0
        End If
    Next i
    
    MsgBox "Data cleaning completed!", vbInformation
End Sub

'===========================
' Function: Remove Symbols
'===========================
Function ReplaceSymbols(txt As String) As String
    Dim temp As String
    temp = txt
    temp = Replace(temp, "$", "")
    temp = Replace(temp, "?", "")
    temp = Replace(temp, "?", "")
    temp = Replace(temp, "USD", "", , , vbTextCompare)
    temp = Replace(temp, "R", "")
    temp = Replace(temp, "rs", "", , , vbTextCompare)
    temp = Replace(temp, "rupees", "", , , vbTextCompare)
    temp = Replace(temp, "pcs", "", , , vbTextCompare)
    temp = Replace(temp, "pc", "", , , vbTextCompare)
    temp = Replace(temp, "kg", "", , , vbTextCompare)
    temp = Replace(temp, "units", "", , , vbTextCompare)
    ReplaceSymbols = Application.WorksheetFunction.Trim(temp)
End Function

'===========================
' Function: Proper Case
'===========================
Function ProperCaseText(txt As String) As String
    ProperCaseText = WorksheetFunction.Proper(txt)
End Function

