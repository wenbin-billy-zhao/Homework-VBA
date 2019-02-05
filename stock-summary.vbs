Sub WorksheetLoop3()

' Declare Current as a worksheet object variable.
Dim ws As Worksheet

' initializae row count and Report Row count (RptRow)
Dim RowCount As Long
Dim RptRow As Long
Dim Ticker As String
Dim Vol As Double
Dim YrOpen As Double
Dim YrClose As Double
Dim YrChange As Double
Dim PctChange As Double
Dim MaxTicker As String
Dim MinTicker As String
Dim MaxVol As Double
Dim MaxVolRow As Long
Dim MaxIncrease As Double
Dim MaxDecrease As Double

MaxVol = 0
MaxIncrease = 0
MaxDecrease = 0


' always start with first sheet
Sheets(1).Select

' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets

' Insert your code here.
' This line displays the worksheet name in a message box.
    
    Vol = 0
    Ticker = ""
    YrOpen = ws.Cells(2, 3).Value
    
    Application.ScreenUpdating = False
    
    
    RptRow = 2
    
    ' find out the row count
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' MsgBox (RowCount) ' just code testing
    
    ws.Range("I" & RptRow - 1).Value = "Ticker"
    ws.Range("I" & RptRow - 1).Font.Bold = True
    ws.Range("J" & RptRow - 1).Value = "YoY Change"
    ws.Range("J" & RptRow - 1).Font.Bold = True
    ws.Range("K" & RptRow - 1).Value = "% Change"
    ws.Range("K" & RptRow - 1).Font.Bold = True
    ws.Range("L" & RptRow - 1).Value = "Vol Total"
    ws.Range("L" & RptRow - 1).Font.Bold = True
    
    For I = 2 To RowCount
    
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            Ticker = ws.Cells(I, 1).Value
            Vol = Vol + ws.Cells(I, 7).Value
            YrClose = ws.Cells(I, 6).Value
            
            YrChange = YrClose - YrOpen
            
            ' error catch for division by 0
            If YrChange = 0 Or YrOpen = 0 Then
                PctChange = 0
            Else
                PctChange = YrChange / YrOpen
            End If
            
            ws.Range("I" & RptRow).Value = Ticker
            ws.Range("J" & RptRow).Value = YrChange
            ws.Range("J" & RptRow).Font.Color = RGB(255, 255, 255)
            
            If YrChange >= 0 Then
                ws.Range("J" & RptRow).Interior.Color = RGB(0, 135, 0)
            Else
                ws.Range("J" & RptRow).Interior.Color = RGB(135, 0, 0)
            End If
             
            ws.Range("K" & RptRow).NumberFormat = "0.00%"
            
            ws.Range("K" & RptRow).Value = PctChange
            ws.Range("L" & RptRow).Value = Vol
            
            YrOpen = ws.Cells(I + 1, 3).Value
            
            RptRow = RptRow + 1
            
            If PctChange > MaxIncrease Then
                MaxIncrease = PctChange
                MaxTicker = Ticker
            ElseIf PctChange < MaxDecrease Then
                MaxDecrease = PctChange
                MinTicker = Ticker
            End If
            
            If MaxVol < Vol Then
                MaxVol = Vol
                MaxVolTicker = Ticker
            End If
            ws.Columns("L:Q").EntireColumn.AutoFit
        Else
            Vol = Vol + ws.Cells(I, 7).Value
            
        End If
        
        
    Next I
    
    
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    
    ws.Range("P2") = MaxTicker
    ws.Range("Q2") = MaxIncrease
    
    
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("P3") = MinTicker
    ws.Range("Q3") = MaxDecrease
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P4") = MaxVolTicker
    ws.Range("Q4") = MaxVol
    ws.Range("Q4").NumberFormat = "#,###"

    ws.Columns("I:Q").EntireColumn.AutoFit
    
    Application.ScreenUpdating = Ture
    
    Vol = 0

Next ws

End Sub

