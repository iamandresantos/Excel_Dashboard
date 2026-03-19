Attribute VB_Name = "Search_by_name"
' Search by Name VBA Code
' Author: Andre dos Santos
' ============================================================
Sub SearchByName()
    ' --- PERFORMANCE: Turn off screen updates and calculations ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsData As Worksheet
    Dim wsResults As Worksheet
    Dim searchTerm As String
    Dim lastRow As Long
    Dim resultRow As Long
    Dim i As Long
    Dim importerVal As String
    Dim exporterVal As String
    Dim found As Long
    Dim j As Integer
    Dim altFill As Long
    Dim sheetName As String

    Set wsData = ThisWorkbook.Sheets("shipment_database")

    searchTerm = InputBox("Enter name to search (importer or exporter):" & Chr(10) & "Partial matches are supported.", "Name Search")
    If Trim(searchTerm) = "" Then GoTo Cleanup

    sheetName = "search_" & searchTerm

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws

    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = sheetName

    Dim headers(1 To 15) As String
    headers(1) = "package_id"
    headers(2) = "arrival_date"
    headers(3) = "package_description"
    headers(4) = "items_qt"
    headers(5) = "package_weight"
    headers(6) = "package_value"
    headers(7) = "flight_id"
    headers(8) = "broker_id"
    headers(9) = "import_query"
    headers(10) = "exporter"
    headers(11) = "exporter_country"
    headers(12) = "importer"
    headers(13) = "importer_country"
    headers(14) = "package_priority"
    headers(15) = "hold_import_responsability"

    For j = 1 To 15
        With wsResults.Cells(1, j)
            .Value = headers(j)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Arial"
            .Font.Size = 10
            .Interior.Color = RGB(68, 114, 196)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
    Next j

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    resultRow = 2
    found = 0

    ' --- PERFORMANCE: Load all data into an array first ---
    Dim dataArr As Variant
    dataArr = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, 15)).Value

    ' --- PERFORMANCE: Build results in an array, write once at the end ---
    Dim resultsArr() As Variant
    ReDim resultsArr(1 To lastRow, 1 To 15)

    Dim lowerTerm As String
    lowerTerm = LCase(searchTerm)

    For i = 2 To lastRow
        importerVal = LCase(CStr(dataArr(i, 12)))
        exporterVal = LCase(CStr(dataArr(i, 10)))

        If InStr(importerVal, lowerTerm) > 0 Or InStr(exporterVal, lowerTerm) > 0 Then
            For j = 1 To 15
                resultsArr(found + 1, j) = dataArr(i, j)
            Next j
            found = found + 1
        End If
    Next i

    ' --- Write all results in a single operation ---
    If found > 0 Then
        wsResults.Range(wsResults.Cells(2, 1), wsResults.Cells(found + 1, 15)).Value = resultsArr
    End If

    ' --- Apply formatting in bulk ---
    With wsResults.Range(wsResults.Cells(2, 1), wsResults.Cells(found + 1, 15))
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .Interior.Color = RGB(210, 227, 252)
    End With

    ' Highlight matched cells in yellow
    For i = 2 To found + 1
        If InStr(LCase(CStr(wsResults.Cells(i, 10).Value)), lowerTerm) > 0 Then
            wsResults.Cells(i, 10).Interior.Color = RGB(255, 235, 156)
        End If
        If InStr(LCase(CStr(wsResults.Cells(i, 12).Value)), lowerTerm) > 0 Then
            wsResults.Cells(i, 12).Interior.Color = RGB(255, 235, 156)
        End If
    Next i

    wsResults.Columns.AutoFit
    wsResults.Activate

    If found = 0 Then
        MsgBox "No records found for [" & searchTerm & "].", vbInformation, "Search Complete"
    Else
        MsgBox found & " record(s) found for [" & searchTerm & "]." & Chr(10) & "Sheet '" & sheetName & "' has been created.", vbInformation, "Search Complete"
    End If

Cleanup:
    ' --- PERFORMANCE: Always restore Excel settings ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
