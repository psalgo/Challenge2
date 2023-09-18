# Challenge2
VBA Challenge
Paul Salgo Challenge 2 VBA

Sources:
Chat.openai.com
  For autofit function: Columns("I:P")

Study group:
   lastrow = Cells(Rows.Count, "A").End(xlUp).Row

I used a tutor and they gave me the framework for the calculated values:
    If percentagechange > greatestincrease Then
            greatestincrease = percentagechange
                Cells(2, 15).Value = Cells(I, "A").Value
                Cells(2, 16).Value = greatestincrease
        ElseIf percentagechange < greatestdecrease Then
            greatestdecrease = percentagechange
                Cells(3, 15).Value = Cells(I, "A").Value
                Cells(3, 16).Value = greatestdecrease
    
