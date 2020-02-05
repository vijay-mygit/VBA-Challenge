Attribute VB_Name = "Module2"
Sub Stocks_singlesheet()
         
        Dim Open_value As Double
        Dim Close_value As Double
        Dim Value_change As Double
        Dim Summary_table As Long
        Dim lastrow As Long
        Dim percent_change As Double
        Dim Total_Volume As Double
        Dim Max_volume As Double
        Dim Max_increase As Double
        Dim Max_decrease As Double
        Dim Max_Tag1 As String
        Dim Max_Tag2 As String
        Dim Max_Tag3 As String
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
        
'Challenge 1 : To Display the greatest percentage increase, decrease and volume

        If Rows.Count <= 1 Then
        Max_volume = 0
        Max_increase = 0
        Max_decrease = 0
        Else: Max_volume = Cells(2, 12)
        Max_increase = Cells(2, 11)
        Max_decrease = Cells(2, 11)
        End If
            For i = 2 To lastrow
                If Cells(i, 12) > Max_volume Then
                    Max_volume = Cells(i, 12)
                    Max_Tag3 = Cells(i, 9)
                End If
                If Cells(i, 11) > Max_increase Then
                    Max_increase = Cells(i, 11)
                    Max_Tag1 = Cells(i, 9)
                End If
                If Cells(i, 11) < Max_decrease Then
                    Max_decrease = Cells(i, 11)
                    Max_Tag2 = Cells(i, 9)
                End If
            Next
        
        Range("N2") = ("Greatest % increase")
        Range("N3") = ("Greatest % decrease")
        Range("N4") = ("Greatest Total Volume")
        Range("o1") = ("Ticker")
        Range("P1") = ("Value")
        Range("P4") = Max_volume
        Range("O4") = Max_Tag3
        Range("P2") = Format(Max_increase, "percent")
        Range("O2") = Max_Tag1
        Range("P3") = Format(Max_decrease, "percent")
        Range("O3") = Max_Tag2
    
        
    End Sub
