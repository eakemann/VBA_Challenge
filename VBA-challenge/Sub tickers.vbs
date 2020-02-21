Sub tickers()

  ' Set an initial variable for holding the ticker name
  Dim ticker_Name As String

  ' Set an initial variable for holding the volume per ticker
  Dim ticker_volume As Double
  ticker_volume = 0

  ' Keep track of the location for each ticker name in the summary column
  Dim ticker_Table_Row As Integer
  ticker_Table_Row = 2

  ' Loop through all tickers
  For i = 2 To 797711

    ' Check if we are still within the same ticker name, if not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      ticker_Name = Cells(i, 1).Value

      ' Add to the ticker volume
      ticker_volume = ticker_volume + Cells(i, 7).Value

      ' Print the ticker name in the summary column I
      Range("I" & ticker_Table_Row).Value = ticker_Name

      ' Print the ticker volume in the summary column L
      Range("L" & ticker_Table_Row).Value = ticker_volume

      ' Add 1 to the summary table row
      ticker_Table_Row = ticker_Table_Row + 1
      
      ' Reset the ticker volume
      ticker_volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker volume
      ticker_volume = ticker_volume + Cells(i, 7).Value

    End If

  Next i

End Sub

