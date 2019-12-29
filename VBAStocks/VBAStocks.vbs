Sub stock_homework():

' dims
    Dim total as Double
    Dim change as Double
    Dim pcntChange as Double
    Dim dayChange as Double
    Dim avgChange as Double
    Dim i as Long
    Dim start as Long
    Dim count as Long
    Dim j as Integer
    Dim day as Integer

' row titles    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Annual Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"

    j = 0
    total = 0
    change = 0
    start = 2

    count = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To count

        If Cells(i + 1, 1).Value <> Cells(i , 1).Value Then

        total = total + Cells (i , 7).Value

' for zero total vals
            If total = 0 Then

            Range("I" & 2 + j).Value = Cells(i , 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0

            Else

                If Cells(start , 3) = 0 Then
                
                    For find_val = start To i

                        If Cells(find_val, 3).Value <> 0 Then

                            start = find_val

                            Exit For
                
                        End If

                    Next find_val

                End If

' for non zero total vals
                change = (Cells(i , 6) - Cells(start , 3))
                pcntChange = Round((change / Cells(start , 3) * 100), 1)

                start = i + 1

                Range("I" & 2 + j).Value = Cells(i , 1).Value
                Range("J" & 2 + j).Value = Round(change, 1)
                Range("K" & 2 + j).Value = "%" & pcntChange
                Range("L" & 2 + j).Value = total

' conditional formt (+/-) - instead of using several if conditionals for formatting, use select case:
                Select Case change

                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4

                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3

                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0

                End Select 

            End If

' reset
            total = 0
            change = 0
            j = j + 1
            day = 0
        Else

            total = total + Cells(i , 7).Value

        End If

    Next i

End Sub