Attribute VB_Name = "Module2"
Sub Stock_Yearly_Change()

ActiveSheet.Range("J2:N100000").ClearContents
For i = 2 To Range("I1").End(4).Row
    For j = 2 To Range("A1").End(4).Row
        If Cells(i, "I") = Cells(j, "A") Then
            'year open
            If Cells(i, "M") = "" Then
                Cells(i, "M") = Cells(j, "C")
            End If
            'year close
            Cells(i, "N") = Cells(j, "F")
            Cells(i, "J") = Cells(i, "N") - Cells(i, "M")
            Cells(i, "K") = Cells(i, "J") / Cells(i, "M")
            Cells(i, "L") = Cells(i, "L") + Cells(j, "G")
            
            'Colorformat the column J based on the cell value
                If Cells(i, "N") < Cells(i, "M") Then
                    Cells(i, "J").Interior.ColorIndex = 3 'Fill the cell red
                Else
                    Cells(i, "J").Interior.ColorIndex = 4 'Fill the cell green
                End If
            'format column K to percentage
                Cells(i, "K").NumberFormat = "0.00%"
        
        End If
    Next
Next
End Sub

