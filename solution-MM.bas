Attribute VB_Name = "Module1"
Sub Solution()
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Activate
        Call ProcessWS
    Next ws
End Sub

' Create ProcessWS to force variables out of scope between sheets.
Sub ProcessWS()
    ' Set up headers
    Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    Range("O1:P1") = Array("Ticker", "Value")
    Range("N2:N4") = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    
    Range("J:J").NumberFormat = "0.00"
    Range("K:K, P2:P3").NumberFormat = "0.00%"
    
    ' Define variables
    Dim ticker As String
    Dim openingPrice, yearlyChange, percentChange, totalVolume As Double
    Dim greatestIncreaseTicker, greatestDecreaseTicker, greatestVolumeTicker As String
    Dim greatestIncrease, greatestDecrease, greatestVolume As Double
    Dim inputRow, lastRow As Long
    Dim outputRow As Integer
    
    ' Establish initial conditions
    ticker = Range("A2")
    openingPrice = Range("C2")
    totalVolume = Range("G2")
    outputRow = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For inputRow = 3 To lastRow
        While (ticker = Cells(inputRow, 1)) ' Loop until arriving at the next ticker symbol
            totalVolume = totalVolume + Cells(inputRow, 7)
            inputRow = inputRow + 1
        Wend
        
        ' Assign results to output cells
        yearlyChange = Cells(inputRow - 1, 6) - openingPrice
        percentChange = yearlyChange / openingPrice
        
        Cells(outputRow, 9) = ticker
        
        Cells(outputRow, 10) = yearlyChange
        If (yearlyChange < 0) Then
            Cells(outputRow, 10).Interior.Color = vbRed
        ElseIf (yearlyChange > 0) Then
            Cells(outputRow, 10).Interior.Color = vbGreen
        End If
        
        Cells(outputRow, 11) = percentChange
        If (percentChange < greatestDecrease) Then
            greatestDecrease = percentChange
            greatestDecreaseTicker = ticker
        ElseIf (percentChange > greatestIncrease) Then
            greatestIncrease = percentChange
            greatestIncreaseTicker = ticker
        End If
        
        Cells(outputRow, 12) = totalVolume
        If (totalVolume > greatestVolume) Then
            greatestVolume = totalVolume
            greatestVolumeTicker = ticker
        End If
        
        ticker = Cells(inputRow, 1)
        openingPrice = Cells(inputRow, 3)
        totalVolume = Cells(inputRow, 7)
        outputRow = outputRow + 1
    Next inputRow
    
    ' Assign overall results
    Range("O2") = greatestIncreaseTicker
    Range("P2") = greatestIncrease
    Range("O3") = greatestDecreaseTicker
    Range("P3") = greatestDecrease
    Range("O4") = greatestVolumeTicker
    Range("P4") = greatestVolume
End Sub
