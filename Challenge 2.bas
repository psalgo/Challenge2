Attribute VB_Name = "Module1"
Sub Challenge2()

'Set everything across the entire spreadsheet
    For Each ws In Worksheets
    ws.Activate



'Put column names and autofit
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest Increase"
    Range("N3").Value = "Greatest Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Columns("I:P").AutoFit
    
    
      


'Need last cell with data for the for loop
'Set totalvolume to zero to keep track of total stock volume for each ticker
'Set open price to respective column
'Set populate to 2 so we can populate data later starting in row 2
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    totalvolume = 0
    openprice = Cells(2, "C").Value
    populate = 2
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0


'Start the for loop
    For I = 2 To lastrow
        totalvolume = totalvolume + Cells(I, "G").Value
'If ticker symbol changes then output
    If Cells(I, "A").Value <> Cells(I + 1, "A").Value Then
        closeprice = Cells(I, "F").Value
        yearlychange = closeprice - openprice
        percentagechange = yearlychange / openprice * 100
        Cells(populate, "I").Value = Cells(I, "A").Value
        Cells(populate, "J").Value = yearlychange
        Cells(populate, "K").Value = "%" & percentagechange
        Cells(populate, "L").Value = totalvolume
'Format cell colors
        If yearlychange > 0 Then
            Range("J" & populate).Interior.Color = vbGreen
        ElseIf yearlychange < 0 Then
            Range("J" & populate).Interior.Color = vbRed
        Else
            Range("J" & populate).Interior.Color = vbWhite
        End If

 
 'Greatest increase, decrease
        If percentagechange > greatestincrease Then
            greatestincrease = percentagechange
                Cells(2, 15).Value = Cells(I, "A").Value
                Cells(2, 16).Value = greatestincrease
        ElseIf percentagechange < greatestdecrease Then
            greatestdecrease = percentagechange
                Cells(3, 15).Value = Cells(I, "A").Value
                Cells(3, 16).Value = greatestdecrease
        End If
'Total Volume
        If totalvolume > greatestvolume Then
            greatestvolume = totalvolume
                Range("O4").Value = Cells(I, "A").Value
                Range("P4").Value = greatestvolume
            End If

   'Reset values
            totalvolume = 0
            openprice = Cells(I + 1, "C").Value
            populate = populate + 1
        End If
        Next I
        Next ws
        MsgBox ("complete")
    End Sub
