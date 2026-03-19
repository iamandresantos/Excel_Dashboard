Attribute VB_Name = "delayed_shipments"
' Delayed Shipments VBA Code
' Author: André dos Santos
' ============================================================
Sub ExportDelayedToNewSheet()
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long

    Set ws = Sheets("shipment_database")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Delete existing sheet if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Delayed Shipments").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' -- Speed optimizations ------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' ----------------------------------------------------

    ' Add new sheet at the end
    Set newWs = Sheets.Add(After:=Sheets(Sheets.Count))
    newWs.Name = "Delayed Shipments"

    ' -- Copy header row only -----------------------------
    ws.Rows(1).Copy newWs.Rows(1)

    ' -- Use AutoFilter on SOURCE sheet to isolate Delayed -
    ws.AutoFilterMode = False                            ' Clear any existing filter
    ws.Range("A1:M" & lastRow).AutoFilter Field:=9, Criteria1:="Delayed"

    ' -- Copy only visible (filtered) rows ----------------
    Dim visibleRows As Range
    On Error Resume Next
    Set visibleRows = ws.Range("A2:M" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleRows Is Nothing Then
        visibleRows.Copy newWs.Range("A2")
    End If

    ' -- Remove filter from source sheet ------------------
    ws.AutoFilterMode = False

    newWs.Columns.AutoFit

    ' -- Restore Excel settings ---------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' ----------------------------------------------------

    Sheets("Delayed Dashboard").Activate

    Dim delayedCount As Long
    delayedCount = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row - 1
    MsgBox delayedCount & " delayed shipment(s) exported!", vbInformation, "Delayed Shipments"

End Sub
