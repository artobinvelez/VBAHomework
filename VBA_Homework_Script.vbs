Sub MacroEverySheet()
   
   Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call StockScript
    Next
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------

Sub StockScript():

Dim ticker As String
Dim TotalVolume As Double
Dim i As Long
Dim year_open As Double
Dim year_close As Double

TotalVolume = 0

Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"

Cells(1, 15).Value = "Statistics"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim lastrow As Long
    With ActiveSheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    End With
    
    For i = 2 To lastrow
            If year_open = 0 Then

            year_open = Cells(i, 3).Value
      
            End If
            
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker = Cells(i, 1).Value
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            year_close = Cells(i, 6).Value
            
            YearlyChange = year_close - year_open
            
            PercentChange = YearlyChange / year_open
            
            Range("J" & Summary_Table_Row).Value = YearlyChange
            
            Range("K" & Summary_Table_Row).Value = PercentChange
            
            Range("I" & Summary_Table_Row).Value = ticker
            'It will change every time hence the Summary Table Row
            
            Range("L" & Summary_Table_Row).Value = TotalVolume
            'Program DRY Don't Repeat Yourself
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            TotalVolume = 0
            'Reset Total Volume
            
            Else
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
        End If
        
    Next i
    
    Range("L2:L" & lastrow).NumberFormat = "#,###,###,###"
    Range("K2:K" & lastrow).NumberFormat = "0.00%"
    
    Columns("A:A").Columns.AutoFit
    Columns("B:B").Columns.AutoFit
    Columns("C:C").Columns.AutoFit
    Columns("D:D").Columns.AutoFit
    Columns("E:E").Columns.AutoFit
    Columns("F:F").Columns.AutoFit
    Columns("G:G").Columns.AutoFit
    Columns("I:I").Columns.AutoFit
    Columns("J:J").Columns.AutoFit
    Columns("K:K").Columns.AutoFit
    Columns("L:L").Columns.AutoFit
   
    For x = 2 To lastrow
        
        If Cells(x, 11) > 0 Then
        
        Cells(x, 11).Interior.ColorIndex = 4
        
        ElseIf Cells(x, 11) < 0 Then
        
        Cells(x, 11).Interior.ColorIndex = 3
        
        End If
    
    Next x
    
    Dim MaximumP As Double
    Dim MinimumP As Double
    Dim MaximumV As Double
    
    MaximumP = Application.WorksheetFunction.Max(Range("K2:K" & lastrow))
    MinimumP = Application.WorksheetFunction.Min(Range("K2:K" & lastrow))
    MaximumV = Application.WorksheetFunction.Max(Range("L2:L" & lastrow))
    
        For y = 2 To lastrow
            
            If Cells(y, 11).Value = MaximumP Then
            
                Range("Q2").Value = Cells(y, 11).Value
                Range("P2").Value = Cells(y, 9).Value
                        
            ElseIf Cells(y, 11).Value = MinimumP Then
            
                Range("Q3").Value = Cells(y, 11).Value
                Range("P3").Value = Cells(y, 9).Value
            
            ElseIf Cells(y, 12).Value = MaximumV Then
            
                Range("Q4").Value = Cells(y, 12).Value
                Range("P4").Value = Cells(y, 9).Value
        
            End If
    
        Next y
    
    Range("Q2:Q3").NumberFormat = "0,000.00%"
    Range("Q4").NumberFormat = "#,###,###,###"
    
    Cells(2, 17).Interior.ColorIndex = 4
    Cells(3, 17).Interior.ColorIndex = 3
    
    Columns("O:O").Columns.AutoFit
    Columns("P:P").Columns.AutoFit
    Columns("Q:Q").Columns.AutoFit
    
        
    
End Sub