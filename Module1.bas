Attribute VB_Name = "Module1"
Sub ConsolidateFiles_SingleHeader()
    Dim fldr As FileDialog
    Dim FolderPath As String
    Dim Filename As String
    Dim wsMaster As Worksheet
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    Dim lastRow As Long, LastCol As Long
    Dim PasteRow As Long
    Dim FirstFile As Boolean
    Dim filesFound As Long, rowsCopied As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1) ask you to pick the folder with your monthly files
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select folder that contains your Excel files (e.g., Jan/Feb/Mar)"
        If .Show <> -1 Then
            MsgBox "No folder selected. Cancelled.", vbExclamation
            GoTo CleanExit
        End If
        FolderPath = .SelectedItems(1)
    End With
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"

    ' 2) create/clear MasterData sheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("MasterData")
    If wsMaster Is Nothing Then
        Set wsMaster = ThisWorkbook.Sheets.Add
        wsMaster.Name = "MasterData"
    Else
        wsMaster.Cells.Clear
    End If
    On Error GoTo 0

    PasteRow = 1
    FirstFile = True

    ' 3) loop all .xlsx files in the selected folder
    Filename = Dir(FolderPath & "*.xlsx")
    Do While Filename <> ""
        filesFound = filesFound + 1

        Set wbTemp = Workbooks.Open(FolderPath & Filename, ReadOnly:=True)
        Set wsTemp = wbTemp.Sheets(1)                    ' use first sheet of each file

        ' find used rows/cols
        lastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
        LastCol = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column

        If lastRow >= 1 Then
            If FirstFile Then
                ' copy header + data from first file
                wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(lastRow, LastCol)) _
                      .Copy wsMaster.Cells(PasteRow, 1)
                PasteRow = PasteRow + lastRow
                FirstFile = False
                rowsCopied = rowsCopied + lastRow
            Else
                ' copy only data (skip header row) from subsequent files
                If lastRow > 1 Then
                    wsTemp.Range(wsTemp.Cells(2, 1), wsTemp.Cells(lastRow, LastCol)) _
                          .Copy wsMaster.Cells(PasteRow, 1)
                    PasteRow = PasteRow + (lastRow - 1)
                    rowsCopied = rowsCopied + (lastRow - 1)
                End If
            End If
        End If

        wbTemp.Close SaveChanges:=False
        Filename = Dir
    Loop

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox filesFound & " files merged. " & rowsCopied & _
           " rows copied into 'MasterData'.", vbInformation
    Exit Sub

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


