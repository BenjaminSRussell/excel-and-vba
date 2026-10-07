Attribute VB_Name = "QueueModule"
' FIFO queue helpers for the sample workbook.
Option Explicit

Public Sub Enqueue(ByVal payload As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Queue")
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = nextRow - 1
    ws.Cells(nextRow, 2).Value = payload
    ws.Cells(nextRow, 3).Value = "pending"
    ws.Cells(nextRow, 4).Value = Now
End Sub

Public Sub ProcessNext()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Queue")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 3).Value = "pending" Then
            ws.Cells(r, 3).Value = "done"
            Exit Sub
        End If
    Next r
End Sub
