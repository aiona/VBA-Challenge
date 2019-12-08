Sub Stock_Loops()

' Initialize variables for Stock Ticker
Dim ws As Worksheet
Dim ticker As String
Dim percent_change As Double
Dim volume As Double
Dim yearly_change As Double
Dim LastRow As Long
Dim open_price As Double
Dim close_price As Double

' Loop through all sheets
For Each ws In Worksheets
    
    ' Keep track of the location for each stock ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    volume = ws.Cells(2, 7).Value

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add the word State to the First Column Header
    ws.Cells(1, 9).Value = "Ticker"

    ' Add the word State to the First Column Header
    ws.Cells(1, 10).Value = "Yearly Change"

    ' Add the word State to the First Column Header
    ws.Cells(1, 11).Value = "Percent Change"

    ' Add the word State to the First Column Header
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set open price
    open_price = ws.Cells(2, 3).Value
    
    ' Error before ticker changes
    For i = 2 To LastRow

    'If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

    ' Check if Next Stock Ticker Equals Current
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
        volume = volume + ws.Cells(i, 7).Value

        ticker = ws.Cells(i, 1).Value

        ' Print ticker in summary table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        
        'Set close price
        close_price = ws.Cells(i, 6).Value
    
        ' Calculate Yearly Change  and print in summary table
        ' ******Conditional formatting
        ' Round to 2 significant digits
        yearly_change = close_price - open_price
         
         
         'If ws.Range("J" & Summary_Table_Row).Value > 0 Then
          '  ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
         
         'ElseIf Range("J" & Summary_Table_Row).Value < 0 Then
            'ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
         
         'End If
         
         ws.Range("J" & Summary_Table_Row).Value = yearly_change
         
         
         ' Calculate percent_change and print in summary table
         ' Conditional Formatting
         
         If open_price = 0 Then
            percent_change = 0
         
         Else
            percent_change = ((close_price - open_price) / open_price)
         
         End If
         
         ws.Range("K" & Summary_Table_Row).Value = percent_change
         ws.Range("K2:K" & LastRow).Style = "Percent"
         ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
         
         ws.Range("L" & Summary_Table_Row).Value = volume
         
         volume = 0
         
         ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
               
    Else
            
        volume = volume + ws.Cells(i, 7).Value
            
    End If
    
Next i

Next ws

End Sub

