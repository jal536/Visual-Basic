Sub TickerTotal()

  Dim Ticker As String

  Range("I1").Value = "Ticker"
  Range("J1").Value = "Total Stock Volume"

  Dim Vol As Double
  Vol = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  Dim i As Long
  Set sht = ActiveSheet

  Dim lRow As Long
  lRow = sht.Range("A1").CurrentRegion.Rows.Count

  For i = 2 To lRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker = Cells(i, 1).Value

      Vol = Vol + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = Ticker

      Range("J" & Summary_Table_Row).Value = Vol

      Summary_Table_Row = Summary_Table_Row + 1

      Vol = 0

    Else

      Vol = Vol + Cells(i, 7).Value

    End If

  Next i

End Sub


