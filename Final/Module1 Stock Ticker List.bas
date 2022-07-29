Attribute VB_Name = "Module1"
Sub Stock_Ticker_List()
  ' Set an initial variable for holding the brand name
  Dim Ticker_Name As String
  ' Set an initial variable for holding the total per credit card brand
  Dim Ticker_Total As Double
  Ticker_Total = 0
  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all stock transactions
  For i = 2 To lastrow
    ' Check if we are still within the same stock ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the ticker name
      Ticker_Name = Cells(i, 1).Value
      ' Add to the ticker Total volume
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      ' Print the ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      ' Print the volume Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Ticker_Total
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker Total
      Ticker_Total = 0
    ' If the cell immediately following a row is the same ticker...
    Else
      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
    End If
  Next i
End Sub


