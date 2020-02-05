Attribute VB_Name = "Module3"
Sub Stocks()
       
        Dim Open_value As Double
        Dim Close_value As Double
        Dim Value_change As Double
        Dim Summary_table As Long
        Dim lastrow As Long
        Dim percent_change As Double
        Dim Total_Volume As Double
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Open_value = Cells(2, 3).Value
        Close_value = 0
        Total_Volume = 0
        Summary_table = 2
        Range("i1") = "Ticker"
        Range("j1") = "Yearly Change"
        Range("k1") = "Percent change"
        Range("l1") = "Total Volume"
            For i = 2 To lastrow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Range("I" & Summary_table).Value = Cells(i, 1).Value
                    Close_value = Cells(i, 6).Value
                    Value_change = (Close_value) - (Open_value)
                    Range("J" & Summary_table).Value = Value_change
                        If Open_value <> 0 Then
                            percent_change = (Value_change / Open_value)
                        Else: percent_change = 0
                        End If
                        If Range("J" & Summary_table).Value > 0 Then
                            Range("J" & Summary_table).Interior.ColorIndex = 4
                            Else: Range("J" & Summary_table).Interior.ColorIndex = 3
                        End If
                    Range("k" & Summary_table).Value = Format(percent_change, "percent")
                    Total_Volume = ((Total_Volume) + Cells(i, 7).Value)
                    Range("L" & Summary_table).Value = (Total_Volume)
                    Summary_table = Summary_table + 1
                    Value_change = 0
                    Close_value = 0
                    Total_Volume = 0
                    Open_value = Cells(i + 1, 3).Value
                    percent_change = 0
                
                    Else
                    Total_Volume = ((Total_Volume) + Cells(i, 7).Value)
                End If
            
            Next i
End Sub
