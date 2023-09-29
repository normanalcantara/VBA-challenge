Attribute VB_Name = "Module1"
Sub Stock_loop():

  ' Set an initial variable for holding the stock name
  Dim Stock_Name As String
  ' Set an initial variable for holding the Total Stock Volume per Stock
  Dim Stock_Total As Double
  Stock_Total = 0
  Dim Year As String

  ' Keep track of the location for each stock name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock volume values
  For i = 2 To Worksheets("A").UsedRange.Rows.Count
    ' Check if we are still within the same stock name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the stock name
      Stock_Name = Cells(i, 1).Value
      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 12).Value
      ' Print the Stock Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Stock_Name
      ' Print the stock Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Stock_Total
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      ' Reset the Stock Total
      Stock_Total = 0
    ' If the cell immediately following a row is the same stock...
    Else
      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value
    End If
  Next i

Cells(1, 9).Value = "Ticker"
Cells(1, 12).Value = "Total Stock Volume"

End Sub



