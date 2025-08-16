Attribute VB_Name = "Module3"
Sub Build_Clean_Sales_Dashboard()
    Dim wsData As Worksheet, wsRpt As Worksheet
    Dim pc As PivotCache
    Dim ptSummary As PivotTable, ptTop As PivotTable
    Dim rngData As Range
    Dim startRow As Long, nextRow As Long
    Dim okMonths As Boolean
    
    '---- Set sheets ----
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("MasterData")
    Set wsRpt = ThisWorkbook.Sheets("Report_Automation")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "Sheet 'MasterData' not found.", vbExclamation: Exit Sub
    End If
    If wsRpt Is Nothing Then
        MsgBox "Sheet 'Report_Automation' not found.", vbExclamation: Exit Sub
    End If
    
    '---- Validate headers ----
    If LCase(wsData.Range("A1").Value) <> "date" _
       Or LCase(wsData.Range("B1").Value) <> "region" _
       Or LCase(wsData.Range("C1").Value) <> "product" _
       Or LCase(wsData.Range("D1").Value) <> "sales" _
       Or LCase(wsData.Range("E1").Value) <> "quantity" Then
        MsgBox "MasterData must have headers in A1:E1: Date | Region | Product | Sales | Quantity", vbExclamation
        Exit Sub
    End If
    
    '---- Data range ----
    Set rngData = wsData.Range("A1").CurrentRegion
    If rngData.Rows.Count < 2 Then
        MsgBox "No rows found in MasterData.", vbExclamation: Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '---- Clear a safe area on Report_Automation (keep your button intact) ----
    wsRpt.Range("B2:Z1000").Clear
    ' Try to delete old slicers (by our names), ignore errors if not present
    On Error Resume Next
    wsRpt.Shapes("slRegion").Delete
    wsRpt.Shapes("slProduct").Delete
    wsRpt.Shapes("slMonths").Delete
    wsRpt.Shapes("slDate").Delete
    On Error GoTo 0
    
    ' Also clear old pivot tables in the area
    Dim p As PivotTable
    For Each p In wsRpt.PivotTables
        If Not Intersect(p.TableRange2, wsRpt.Range("B2:Z1000")) Is Nothing Then
            p.TableRange2.Clear
        End If
    Next p
    
    '---- Create cache ----
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngData)
    
    '========================
    ' 1) SUMMARY PIVOT (compact)
    '========================
    Set ptSummary = pc.CreatePivotTable(TableDestination:=wsRpt.Range("B8"), TableName:="ptSummary")
    With ptSummary
        ' Rows: Region only (short report)
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1
        
        ' Columns: Date (try to group by Months)
        .PivotFields("Date").Orientation = xlColumnField
        .PivotFields("Date").Position = 1
        
        On Error Resume Next
        .PivotFields("Date").NumberFormat = "dd-mmm-yyyy"
        .PivotFields("Date").Group Start:=True, End:=True, By:=xlMonths
        okMonths = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        
        ' Values: Sales + Quantity
        .AddDataField .PivotFields("Sales"), "Total Sales", xlSum
        .PivotFields("Total Sales").NumberFormat = "#,##0"
        
        .AddDataField .PivotFields("Quantity"), "Total Qty", xlSum
        .PivotFields("Total Qty").NumberFormat = "#,##0"
        
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleLight16"
    End With
    
    ' Heading
    wsRpt.Range("B4").Value = "Sales Dashboard (Compact)"
    With wsRpt.Range("B4")
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    '========================
    ' 2) TOP 5 PRODUCTS (tiny pivot)
    '========================
    nextRow = ptSummary.TableRange2.Row + ptSummary.TableRange2.Rows.Count + 2
    Set ptTop = pc.CreatePivotTable(TableDestination:=wsRpt.Cells(nextRow, 2), TableName:="ptTop5")
    With ptTop
        .PivotFields("Product").Orientation = xlRowField
        .AddDataField .PivotFields("Sales"), "Total Sales", xlSum
        .PivotFields("Total Sales").NumberFormat = "#,##0"
        
        ' Top 5 filter (version-safe)
        On Error Resume Next
        .PivotFields("Product").PivotFilters.Add2 Type:=xlTopCount, DataField:=.PivotFields("Total Sales"), Value1:=5
        If Err.Number <> 0 Then
            Err.Clear
            ' fallback: no filter if Add2 not available
        End If
        On Error GoTo 0
        
        .TableStyle2 = "PivotStyleLight18"
    End With
    wsRpt.Cells(nextRow - 1, 2).Value = "Top 5 Products by Sales"
    wsRpt.Cells(nextRow - 1, 2).Font.Bold = True
    
    '========================
    ' 3) SLICERS (Region, Product, Month/Date) + connect to both pivots
    '========================
    Dim ok As Boolean
    ok = SafeAddSlicerCacheConnect(ptSummary, ptTop, "Region", wsRpt, "slRegion", "Region", 60, 420, 150, 140)
    ok = SafeAddSlicerCacheConnect(ptSummary, ptTop, "Product", wsRpt, "slProduct", "Product", 60, 580, 180, 140)
    
    ' Month slicer: if grouping worked, field name is "Months"; else use "Date"
    If okMonths Then
        ok = SafeAddSlicerCacheConnect(ptSummary, ptTop, "Months", wsRpt, "slMonths", "Month", 60, 770, 150, 140)
    Else
        ok = SafeAddSlicerCacheConnect(ptSummary, ptTop, "Date", wsRpt, "slDate", "Date", 60, 770, 150, 140)
    End If
    
    ' Auto-fit
    ptSummary.TableRange2.Columns.AutoFit
    ptTop.TableRange2.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "? Clean, compact dashboard created on 'Report_Automation'.", vbInformation
End Sub


'========================
' Helper: Add slicer (Add2 fallback to Add) and connect to 2 pivots
'========================
Function SafeAddSlicerCacheConnect(ByVal ptA As PivotTable, ByVal ptB As PivotTable, _
                                   ByVal fieldName As String, ByVal ws As Worksheet, _
                                   ByVal slName As String, ByVal caption As String, _
                                   ByVal topPos As Double, ByVal leftPos As Double, _
                                   ByVal w As Double, ByVal h As Double) As Boolean
    Dim sc As SlicerCache
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches.Add2(ptA, fieldName)
    If sc Is Nothing Then
        Err.Clear
        Set sc = ThisWorkbook.SlicerCaches.Add(ptA, fieldName)
    End If
    On Error GoTo 0
    If sc Is Nothing Then SafeAddSlicerCacheConnect = False: Exit Function
    
    ' connect to second pivot
    On Error Resume Next
    sc.PivotTables.AddPivotTable ptB
    On Error GoTo 0
    
    ' add slicer shape
    On Error Resume Next
    sc.Slicers.Add SlicerDestination:=ws, Name:=slName, caption:=caption, _
                   Top:=topPos, Left:=leftPos, Width:=w, Height:=h
    On Error GoTo 0
    
    SafeAddSlicerCacheConnect = True
End Function


