VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ProcessTickerData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim firstOpeningPrice As Double
    Dim lastClosingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long

    ' Set the worksheet to the active sheet
    ' Set ws = ActiveSheet
    
    For Each ws In Worksheets
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Initialize the output row
    outputRow = 2

    ' Loop through the data in column A
    For i = 2 To lastRow
        ' Get the ticker symbol
        ticker = ws.Cells(i, 1).Value

        ' Check if it's a new ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' Output the ticker symbol to column I
            ws.Cells(outputRow, 9).Value = ticker

            ' Calculate total stock volume for the ticker
            totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), ticker, ws.Range("G:G"))
            ws.Cells(outputRow, 12).Value = totalVolume

            ' Get the opening price for the ticker
            openingPrice = ws.Cells(Application.WorksheetFunction.Match(ticker, ws.Range("A:A"), 0), 3).Value

            ' Get the closing price for the ticker
            closingPrice = ws.Cells(Application.WorksheetFunction.Match(ticker, ws.Range("A:A"), 1), 6).Value

            ' Calculate yearly change
            yearlyChange = closingPrice - openingPrice
            ws.Cells(outputRow, 11).Value = yearlyChange

            ' Calculate percentage change
            percentChange = (closingPrice - openingPrice) / openingPrice
            ws.Cells(outputRow, 10).NumberFormat = "0.00%"
            ws.Cells(outputRow, 10).Value = percentChange

            ' Apply conditional formatting based on the value of yearlyChange
            Select Case yearlyChange
                Case Is > 0
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Case Is < 0
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                Case Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = xlNone ' No fill
            End Select

            ' Move to the next output row
            outputRow = outputRow + 1
        End If
    Next i
    Next
End Sub


Sub FindGreatestValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim volume As Double
    Dim percentChange As Double
    Dim i As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Find the last row in column K
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

    ' Initialize variables
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0

    ' Loop through the data in column K
    For i = 2 To lastRow
        ' Get the ticker symbol
        ticker = ws.Cells(i, 9).Value

        ' Get the percentage change
        percentChange = ws.Cells(i, 11).Value

        ' Check for greatest percentage increase
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxIncreaseTicker = ticker
        End If

        ' Check for greatest percentage decrease
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = ticker
        End If

        ' Get the total volume
        volume = ws.Cells(i, 12).Value

        ' Check for greatest total volume
        If volume > maxVolume Then
            maxVolume = volume
            maxVolumeTicker = ticker
        End If
    Next i

    ' Output the results to columns O and P
    ws.Cells(2, 15).Value = maxIncreaseTicker
    ws.Cells(2, 16).Value = maxIncrease

    ws.Cells(3, 15).Value = maxDecreaseTicker
    ws.Cells(3, 16).Value = maxDecrease

    ws.Cells(4, 15).Value = maxVolumeTicker
    ws.Cells(4, 16).Value = maxVolume
End Sub

