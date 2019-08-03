Sub TickerTotal()

  Dim Ticker As String

  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"

  Dim Vol As Double
  Vol = 0

  Dim yearly_change As Double
  yearly_change = 0

  Dim percent As String
  percent = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  Dim i As Long
  Set sht = ActiveSheet

  Dim lRow As Long
  lRow = sht.Range("A1").CurrentRegion.Rows.Count

  Dim first_open_price As Double
  first_open_price = Cells(2, 3).Value

  For i = 2 To lRow

    'If next row is a new ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      Dim last_close_price As Double
      last_close_price = Cells(i, 6).Value
      
      Ticker = Cells(i, 1).Value

      Vol = Vol + Cells(i, 7).Value

      yearly_change = last_close_price - first_open_price
      
      percent = FormatPercent(((last_close_price / first_open_price) - 1))
      
      Range("I" & Summary_Table_Row).Value = Ticker
      
      Range("J" & Summary_Table_Row).Value = yearly_change
      
      Dim temp As String
      temp = Range("J" & Summary_Table_Row).Value
      
      If Range("J" & Summary_Table_Row).Value > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      ElseIf Range("J" & Summary_Table_Row).Value < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      Else: Range("J" & Summary_Table_Row).Interior.ColorIndex = 2
      End If
      
      Range("K" & Summary_Table_Row).Value = percent

      Range("L" & Summary_Table_Row).Value = Vol

      Summary_Table_Row = Summary_Table_Row + 1

      yearly_change = 0

      percent = 0

      Vol = 0
      ' Reset firstOpenPrice to row i + 1
      firstOpenPrice = Cells(i + 1, 3).Value

    Else
      
      Vol = Vol + Cells(i, 7).Value

    End If

  Next i

End Sub

