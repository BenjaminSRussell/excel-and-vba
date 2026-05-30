Option Explicit

Sub MatchProductionToDemand()
    ' Define variables
    Dim wsDemand As Worksheet, wsProduction As Worksheet, wsReport As Worksheet
    Dim lastRowDemand As Long, lastRowProduction As Long
    Dim rngDemand As Range, rngProduction As Range
    Dim demandRow As Range, productionRow As Range
    Dim matchedQuantity As Double, remainingQuantity As Double
    Dim reportNextRow As Long

    ' Set worksheets
    Set wsDemand = ThisWorkbook.Worksheets("Demand")
    Set wsProduction = ThisWorkbook.Worksheets("Production")
    Set wsReport = ThisWorkbook.Sheets.Add
    wsReport.Name = "Reconciliation Report"

    ' Set headers for reconciliation report
    With wsReport
        .Range("A1:E1").Value = Array("Demand ID", "Demand Quantity", "Fulfillment Quantity", "Production ID", "Production Quantity")
        .Columns("A:E").AutoFit
    End With
    reportNextRow = 2

    ' Get last rows to determine ranges
    lastRowDemand = wsDemand.Cells(wsDemand.Rows.Count, "A").End(xlUp).Row
    lastRowProduction = wsProduction.Cells(wsProduction.Rows.Count, "A").End(xlUp).Row

    ' Define ranges for Demand and Production data
    Set rngDemand = wsDemand.Range("A2:D" & lastRowDemand) ' Assuming columns: A=ID, B=Quantity, C=Date, D=Matched Production ID
    Set rngProduction = wsProduction.Range("A2:D" & lastRowProduction) ' Assuming columns: A=ID, B=Quantity, C=Date, D=Remaining Quantity

    ' Sort Demand by Due Date (FIFO)
    rngDemand.Sort Key1:=wsDemand.Columns("C"), Order1:=xlAscending, Header:=xlNo

    ' Sort Production by Completion Date (FIFO)
    rngProduction.Sort Key1:=wsProduction.Columns("C"), Order1:=xlAscending, Header:=xlNo

    ' Match Production to Demand (FIFO logic)
    For Each productionRow In rngProduction.Rows
        Dim prodId As String
        Dim prodQuantity As Double
        Dim prodDate As Date

        ' Get production data
        prodId = productionRow.Cells(1, 1).Value
        prodQuantity = productionRow.Cells(1, 2).Value
        prodDate = productionRow.Cells(1, 3).Value

        For Each demandRow In rngDemand.Rows
            Dim demandId As String
            Dim demandQuantity As Double
            Dim demandDate As Date
            Dim fulfillment As Double

            ' Get demand data
            demandId = demandRow.Cells(1, 1).Value
            demandQuantity = demandRow.Cells(1, 2).Value
            demandDate = demandRow.Cells(1, 3).Value

            ' Skip if demand is already fulfilled
            If demandRow.Cells(1, 4).Value <> "" Then GoTo SkipToNextDemand

            ' Match production to demand based on dates and FIFO
            If prodDate <= demandDate And prodQuantity > 0 Then
                ' Calculate fulfillment for current row
                fulfillment = WorksheetFunction.Min(prodQuantity, demandQuantity)

                ' Update demand row with matched production data
                demandRow.Cells(1, 4).Value = prodId ' Matched production ID
                demandRow.Interior.Color = RGB(144, 238, 144) ' Highlight row in green

                ' Deduct from both production and demand quantities
                demandQuantity = demandQuantity - fulfillment
                prodQuantity = prodQuantity - fulfillment

                ' Add to reconciliation report
                wsReport.Cells(reportNextRow, 1).Value = demandId
                wsReport.Cells(reportNextRow, 2).Value = demandRow.Cells(1, 2).Value
                wsReport.Cells(reportNextRow, 3).Value = fulfillment
                wsReport.Cells(reportNextRow, 4).Value = prodId
                wsReport.Cells(reportNextRow, 5).Value = productionRow.Cells(1, 2).Value

                reportNextRow = reportNextRow + 1

                ' If demand is fully satisfied, move to the next row
                If demandQuantity <= 0 Then Exit For
            End If

SkipToNextDemand:
        Next demandRow

        ' Update remaining production quantity
        productionRow.Cells(1, 4).Value = prodQuantity
    Next productionRow

    ' Format reconciliation report
    With wsReport
        .Columns("A:E").AutoFit
        .Rows(1).Font.Bold = True
    End With

    MsgBox "Production-to-demand matching completed and reconciliation report generated.", vbInformation

End Sub