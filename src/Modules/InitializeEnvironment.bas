Option Explicit

' Global collections for demand and production rows
Dim DemandList As Collection
Dim ProductionList As Collection
Dim wsReport As Worksheet

Sub InitializeEnvironment()
    On Error GoTo ErrorHandler
    
    ' Define sheets
    Dim wsDemand As Worksheet, wsProduction As Worksheet
    Set wsDemand = GetOrCreateSheet("Demand")
    Set wsProduction = GetOrCreateSheet("Production")
    Set wsReport = PrepareReconciliationWorksheet()
    
    ' Initialize global collections
    Set DemandList = New Collection
    Set ProductionList = New Collection
    
    MsgBox "Environment successfully initialized. Sheets and global variables are ready.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during initialization: " & Err.Description, vbCritical
End Sub

Function GetOrCreateSheet(sheetName As String) As Worksheet
    ' Check if a sheet exists; create it if it does not
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Sheet does not exist; create it
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
        Call InitializeDefaultHeaders(ws, sheetName)
    End If
    Set GetOrCreateSheet = ws
End Function

Sub InitializeDefaultHeaders(ws As Worksheet, sheetType As String)
    ' Add default headers to new sheets based on their type
    Select Case sheetType
        Case "Demand"
            ws.Range("A1:C1").Value = Array("Demand ID", "Quantity", "Due Date")
        Case "Production"
            ws.Range("A1:C1").Value = Array("Production ID", "Quantity", "Completion Date")
    End Select
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

Function PrepareReconciliationWorksheet() As Worksheet
    ' Creates or prepares the reconciliation worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Reconciliation Report").Delete ' Delete existing report if any
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Reconciliation Report"
    ws.Range("A1:D1").Value = Array("Demand ID", "Original Quantity", "Fulfilled Quantity", "Matched Production ID")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").AutoFit
    Set PrepareReconciliationWorksheet = ws
End Function