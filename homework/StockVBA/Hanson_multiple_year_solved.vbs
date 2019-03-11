Sub stock()

'Loop through all the worksheets in the workbook
    For Each ws In Worksheets
        'Set one counter variable as double
        Dim i As Double
        
        'Set array to double, ticker to string
        Dim volumeTotal As Double
        volumeTotal = 0
        Dim ticker As String
        
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'loop through all rows in column one to identify each unique ticker
        For i = 2 To 79772

            ' Check if current cell has same ticker as previous cell in column, and if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the ticker name
                tickerName = ws.Cells(i, 1).Value
                ' Add to the dataArray
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                
                ' Print the ticker in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = tickerName
                
                ' Print the volumeTotal Amount to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = volumeTotal

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'reset volumeTotal
                volumeTotal = 0
                
            ' if current cell does have same ticker as previous cell in column...
            Else
            ' Add to volumeTotal
            volumeTotal = volumeTotal + ws.Cells(i, 7).Value
            
            End If
        Next i
    Next ws

End Sub
