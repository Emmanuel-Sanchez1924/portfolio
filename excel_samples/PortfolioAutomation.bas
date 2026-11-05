Attribute VB_Name = "PortfolioAutomation"
Option Explicit

' Portfolio demo workbook: expected tables/sheets are listed below.
' Tables: tblTasks on "Task Tracker"; tblInvoices on "Invoice Control".
' All records in the sample workbook are fictional.

Public Sub RefreshPortfolioDashboard()
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing operations dashboard..."
    ThisWorkbook.Calculate
    Application.CalculateFull
    ThisWorkbook.Worksheets("Ops Dashboard").Activate
    ThisWorkbook.Worksheets("Ops Dashboard").Range("A1").Select
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed.", vbInformation, "Operations Portfolio"
    Exit Sub
CleanFail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "The dashboard could not be refreshed: " & Err.Description, vbExclamation, "Operations Portfolio"
End Sub

Public Sub AddPortfolioTask()
    Dim ws As Worksheet, tbl As ListObject, newRow As ListRow
    Dim taskName As String, ownerName As String, priority As String
    Dim dueInput As String, dueDate As Date

    Set ws = ThisWorkbook.Worksheets("Task Tracker")
    Set tbl = ws.ListObjects("tblTasks")

    taskName = Trim$(InputBox("Enter the task or deliverable:", "Add Task"))
    If Len(taskName) = 0 Then Exit Sub
    ownerName = Trim$(InputBox("Enter the owner:", "Add Task"))
    If Len(ownerName) = 0 Then Exit Sub
    priority = Trim$(InputBox("Enter priority: High, Medium, or Low", "Add Task", "Medium"))
    If priority <> "High" And priority <> "Medium" And priority <> "Low" Then
        MsgBox "Priority must be High, Medium, or Low.", vbExclamation, "Add Task"
        Exit Sub
    End If
    dueInput = Trim$(InputBox("Enter the due date (for example, 11/20/2026):", "Add Task"))
    If Len(dueInput) = 0 Then Exit Sub
    If Not IsDate(dueInput) Then
        MsgBox "Please enter a valid date.", vbExclamation, "Add Task"
        Exit Sub
    End If
    dueDate = CDate(dueInput)

    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, tbl.ListColumns("TASK ID").Index).Value = "T-" & Format$(Now, "yymmddhhmmss")
        .Cells(1, tbl.ListColumns("TASK / DELIVERABLE").Index).Value = taskName
        .Cells(1, tbl.ListColumns("OWNER").Index).Value = ownerName
        .Cells(1, tbl.ListColumns("PRIORITY").Index).Value = priority
        .Cells(1, tbl.ListColumns("DUE DATE").Index).Value = dueDate
        .Cells(1, tbl.ListColumns("STATUS").Index).Value = "Not Started"
        .Cells(1, tbl.ListColumns("EST. HOURS").Index).Value = 0
        .Cells(1, tbl.ListColumns("ACTUAL HOURS").Index).Value = 0
    End With
    Application.CalculateFull
    MsgBox "Task added to the tracker.", vbInformation, "Add Task"
End Sub

Public Sub ExportDashboardPDF()
    Dim target As Variant, defaultName As String
    If Len(ThisWorkbook.Path) > 0 Then
        defaultName = ThisWorkbook.Path & Application.PathSeparator & "Operations_Status_" & Format$(Date, "yyyymmdd") & ".pdf"
    Else
        defaultName = "Operations_Status_" & Format$(Date, "yyyymmdd") & ".pdf"
    End If
    target = Application.GetSaveAsFilename(InitialFileName:=defaultName, FileFilter:="PDF Files (*.pdf), *.pdf")
    If VarType(target) = vbBoolean And target = False Then Exit Sub

    On Error GoTo ExportFail
    ThisWorkbook.Worksheets("Ops Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=CStr(target), Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "Dashboard PDF exported.", vbInformation, "Operations Portfolio"
    Exit Sub
ExportFail:
    MsgBox "The PDF could not be exported: " & Err.Description, vbExclamation, "Operations Portfolio"
End Sub

Public Sub CreateOverdueReport()
    Dim taskTable As ListObject, invoiceTable As ListObject
    Dim reportBook As Workbook, reportSheet As Worksheet
    Dim rowItem As ListRow, nextRow As Long, lastRow As Long
    Dim dueValue As Variant, statusValue As String
    Dim saveTarget As Variant, fileName As String

    On Error GoTo ReportFail
    Set taskTable = ThisWorkbook.Worksheets("Task Tracker").ListObjects("tblTasks")
    Set invoiceTable = ThisWorkbook.Worksheets("Invoice Control").ListObjects("tblInvoices")
    Set reportBook = Workbooks.Add(xlWBATWorksheet)
    Set reportSheet = reportBook.Worksheets(1)
    reportSheet.Name = "Overdue Actions"
    reportSheet.Range("A1:H1").Value = Array("RECORD TYPE", "ID", "TASK / CLIENT", "OWNER / NEXT ACTION", "DUE DATE", "AMOUNT / PRIORITY", "DAYS OVERDUE", "STATUS")
    nextRow = 2

    For Each rowItem In taskTable.ListRows
        dueValue = rowItem.Range.Cells(1, taskTable.ListColumns("DUE DATE").Index).Value
        statusValue = CStr(rowItem.Range.Cells(1, taskTable.ListColumns("STATUS").Index).Value)
        If IsDate(dueValue) And CDate(dueValue) < Date And LCase$(statusValue) <> "completed" Then
            reportSheet.Cells(nextRow, 1).Value = "Task"
            reportSheet.Cells(nextRow, 2).Value = rowItem.Range.Cells(1, taskTable.ListColumns("TASK ID").Index).Value
            reportSheet.Cells(nextRow, 3).Value = rowItem.Range.Cells(1, taskTable.ListColumns("TASK / DELIVERABLE").Index).Value
            reportSheet.Cells(nextRow, 4).Value = rowItem.Range.Cells(1, taskTable.ListColumns("OWNER").Index).Value
            reportSheet.Cells(nextRow, 5).Value = CDate(dueValue)
            reportSheet.Cells(nextRow, 6).Value = rowItem.Range.Cells(1, taskTable.ListColumns("PRIORITY").Index).Value
            reportSheet.Cells(nextRow, 7).Value = Date - CDate(dueValue)
            reportSheet.Cells(nextRow, 8).Value = statusValue
            nextRow = nextRow + 1
        End If
    Next rowItem

    For Each rowItem In invoiceTable.ListRows
        statusValue = CStr(rowItem.Range.Cells(1, invoiceTable.ListColumns("COLLECTION STATUS").Index).Value)
        dueValue = rowItem.Range.Cells(1, invoiceTable.ListColumns("DUE DATE").Index).Value
        If LCase$(statusValue) = "overdue" And IsDate(dueValue) Then
            reportSheet.Cells(nextRow, 1).Value = "Invoice"
            reportSheet.Cells(nextRow, 2).Value = rowItem.Range.Cells(1, invoiceTable.ListColumns("INVOICE ID").Index).Value
            reportSheet.Cells(nextRow, 3).Value = rowItem.Range.Cells(1, invoiceTable.ListColumns("CLIENT").Index).Value
            reportSheet.Cells(nextRow, 4).Value = rowItem.Range.Cells(1, invoiceTable.ListColumns("NEXT ACTION").Index).Value
            reportSheet.Cells(nextRow, 5).Value = CDate(dueValue)
            reportSheet.Cells(nextRow, 6).Value = rowItem.Range.Cells(1, invoiceTable.ListColumns("BALANCE DUE").Index).Value
            reportSheet.Cells(nextRow, 7).Value = Date - CDate(dueValue)
            reportSheet.Cells(nextRow, 8).Value = statusValue
            nextRow = nextRow + 1
        End If
    Next rowItem

    lastRow = nextRow - 1
    With reportSheet
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Color = RGB(255, 255, 255)
        .Rows(1).Interior.Color = RGB(18, 52, 90)
        .Columns("A:H").AutoFit
        .Columns("E:E").NumberFormat = "mmm d, yyyy"
        .Columns("F:F").NumberFormat = "$#,##0.00"
        .Columns("G:G").NumberFormat = "0"
        If lastRow >= 2 Then
            .ListObjects.Add(xlSrcRange, .Range("A1:H" & lastRow), , xlYes).Name = "tblOverdueActions"
            .Range("A1:H" & lastRow).AutoFilter
        End If
        .Range("A1").Select
    End With
    fileName = "Overdue_Action_Report_" & Format$(Date, "yyyymmdd") & ".xlsx"
    saveTarget = Application.GetSaveAsFilename(InitialFileName:=fileName, FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If VarType(saveTarget) = vbBoolean And saveTarget = False Then Exit Sub
    Application.DisplayAlerts = False
    reportBook.SaveAs Filename:=CStr(saveTarget), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    MsgBox "Overdue report created with " & Application.Max(0, lastRow - 1) & " action item(s).", vbInformation, "Operations Portfolio"
    Exit Sub
ReportFail:
    Application.DisplayAlerts = True
    MsgBox "The overdue report could not be created: " & Err.Description, vbExclamation, "Operations Portfolio"
End Sub
