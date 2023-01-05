Sub stocks()
    For Each ws In Worksheets

        'declare variables
        Dim ticker As String
        Dim yearchange As Double
        Dim annual_open As Double
        Dim pointer As Integer
        c_pointer = 2

        Dim annual_close As Double
        Dim percentchange As Double

        Dim totalvolume As Double
        totalvolume = 0

        Dim summaryrow As Integer
        summaryrow = 2

        'headers/ columns
        ws.Cells(1, "I").Value = "ticker"
        ws.Cells(1, "J").Value = "Yearly Change($)"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Volume"
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row


        'start loop
        For i = 2 To RowCount

        'if row i does not equal i+1 (i.e ticker changes)
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then

                'annual change formula

                annual_open = ws.Cells(c_pointer, "C").Value

                annual_close = ws.Cells(i, "F").Value

                'yearchange will equal open - close
                yearchange = annual_close - annual_open

                'percentchange formula for ticker
                percentchange = (yearchange / annual_open) * 100

                'total volume summed up
                totalvolume = ws.Cells(i, "G").Value + totalvolume
                
                'ticker name cell location
                ws.Cells(summaryrow, "I").Value = ws.Cells(i, "A").Value

                'yearchange cell location and formating
                ws.Cells(summaryrow, "J").Value = yearchange

                If ws.Cells(summaryrow, "J").Value > 0 Then
                    ws.Cells(summaryrow, "J").Interior.Color = vbGreen
                Else
                ws.Cells(summaryrow, "J").Interior.Color = vbRed
                
                End If

                'percentchange cell location
                ws.Cells(summaryrow, "K").Value = "%" & percentchange

                'Totalvolume cell location
                ws.Cells(summaryrow, "L").Value = totalvolume
                

                'move summary row down
                summaryrow = summaryrow + 1
                
                'move pointer row to next ticker value
                c_pointer = i + 1
                
                'reset total volume for next ticker value
                totalvolume = 0


            Else
                totalvolume = ws.Cells(i, "G").Value + totalvolume

            End If

        Next i
        
        'column names for greatest increase/decrease tables
        ws.Range("O2").Value = "Greatest Percent Increase"
        ws.Range("O3").Value = "Greatest Percent Decrease"
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        RowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'greatest and smallest decrease
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        
        For K = 2 To RowCount
           
           If ws.Cells(K, "K").Value = ws.Range("Q2") Then
              ws.Range("P2").Value = ws.Cells(K, "I").Value
          
          
           ElseIf ws.Cells(K, "K").Value = ws.Range("Q3") Then
              ws.Range("P3").Value = ws.Cells(K, "I").Value
           
           
          End If
          
          If ws.Cells(K, "L").Value = ws.Range("Q4") Then
            ws.Range("P4").Value = ws.Cells(K, "I")
            
          End If
           
           
         
      Next K
    
    Next ws
 
 End Sub










