Attribute VB_Name = "Module1"
Sub Stock_Test()
    
    [i1] = "Ticker"
    [j1] = "Yearly Change"
    [k1] = "Percent Change"
    [l1] = "Total Stock Volume"
    [p1] = "Ticker"
    [q1] = "Value"
    Max_Increase = 0
    Max_Decrease = 0
    Largest_Volume = 0
    Rng = Range("K:K")
    Rng2 = Range("L:L")
    Columns("I:L").AutoFit
    Columns("K").NumberFormat = "0.00%"
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
    LastRow = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    SI = 2
    First_Open = 0
    Total = 0
    For i = 2 To LastRow
    
            Total = Total + Cells(i, "G")
    
            If First_Open = 0 Then
                First_Open = Cells(i, "C")
            End If
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Cells(SI, "I") = Cells(i, 1)
                YearlyCh = Cells(i, "f") - First_Open
                Cells(SI, "J") = YearlyCh
                
                If YearlyCh > 0 Then
                    Cells(SI, "J").Interior.ColorIndex = 4
                    
                Else
                    Cells(SI, "J").Interior.ColorIndex = 3
                End If
                
                Cells(SI, "K") = YearlyCh / First_Open
                Cells(SI, "L") = Total
                Cells(2, 17) = Max_Increase
                Cells(3, 17) = Max_Decrease
                Cells(4, 17) = Largest_Volume
                Largest_Volume = Application.WorksheetFunction.Max(Rng2)
                Max_Increase = Application.WorksheetFunction.Max(Rng)
                Max_Decrease = Application.WorksheetFunction.Min(Rng)
                Total = 0
                First_Open = 0
                SI = SI + 1
            End If
    Next i
    Next ws
End Sub

