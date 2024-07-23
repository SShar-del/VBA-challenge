Attribute VB_Name = "Module1"
Sub stock_summary()
 ' Create variables
 
Dim ws As Worksheet

Dim r As Long


' LOOP THROUGH ALL SHEETS

   For Each ws In Worksheets
   
    ' Create variables to hold the counter and values
   
        Dim LastRow As Long
        Dim strow As Long
        Dim stcolt As Integer
        Dim stcolqc As Integer
        Dim stcolpc As Integer
        Dim stcoltsv As Integer
        
        strow = 1
        stcolt = 9
        stcolqc = 10
        stcolpc = 11
        stcoltsv = 12
              
         ' Set an initial variables
        Dim ticker As String
        ticker = ""
        Dim qopenval As Double
        Dim qcloseval As Double
        Dim qchg As Double
        
        Dim percchg As Double
        
        Dim stkvol As Double
        Dim totalstkvol As Double
         
         ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
         ' Set summary table column headers and format
         
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Quarterly Change"
         ws.Range("K1").Value = "Percentage Change"
         ws.Range("L1").Value = "Total Stock Volume"
         ws.Columns("J").NumberFormat = "0.00"
         ws.Columns("K").NumberFormat = "0.00%"
            

    ' Loop through all stock entries

     For r = 2 To LastRow

        ' Check if we are still within the same ticker name, if we are not...
         
         If ws.Cells(r, 1).Value <> ticker Then
              strow = strow + 1
              totalstkvol = 0
              ticker = Cells(r, 1).Value
              ws.Cells(strow, stcolt).Value = ticker
              qchg = 0
              qopenval = ws.Cells(r, 3).Value
              qcloseval = ws.Cells(r, 6).Value
              qchg = qcloseval - qopenval
              ws.Cells(strow, stcolqc).Value = qchg
              
              percchg = 0
              percchg = (qcloseval - qopenval) / qopenval
              ws.Cells(strow, stcolpc).Value = percchg
             
              totalstkvol = 0
              stkvol = ws.Cells(r, 7).Value
              totalstkvol = totalstkvol + stkvol
              ws.Cells(strow, stcoltsv).Value = totalstkvol
              'Cells(strow, stcolpc).Activate
          Else
                If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                    qcloseval = ws.Cells(r, 6).Value
                    qchg = qcloseval - qopenval
                    ws.Cells(strow, stcolqc).Value = qchg
                    percchg = (qcloseval - qopenval) / qopenval
                    ws.Cells(strow, stcolpc).Value = percchg
                End If
             
              stkvol = Cells(r, 7).Value
              totalstkvol = totalstkvol + stkvol
              ws.Cells(strow, stcoltsv).Value = totalstkvol
              
          End If
       
    Next r

' Autofit columns for better readability
ws.Columns("J").AutoFit
ws.Columns("K").AutoFit
ws.Columns("L").AutoFit


       ' Set variables for more analysis
            
            
            Dim maxPerInc As Double
            Dim maxPerDec As Double
            Dim maxTotVol As Double
            Dim sumTabLastRow As Long
            Dim ColumnRangeI As Range
            Dim ColumnRangeJ As Range
            Dim ColumnRangeK As Range
            Dim ColumnRangeL As Range
            Dim tickerMatchRow As Long
            Dim tickerMatch As String
       ' Find last row of summary table
       
            sumTabLastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
          
       ' Set range for performing functions
            Set ColumnRangeI = ws.Range("I2:I" & sumTabLastRow)
            Set ColumnRangeJ = ws.Range("J2:J" & sumTabLastRow)
            Set ColumnRangeK = ws.Range("K2:K" & sumTabLastRow)
            Set ColumnRangeL = ws.Range("L2:L" & sumTabLastRow)
            
        ' Perform functions and display formatted values in worksheet
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            ws.Range("O2").Value = "Greatest % Increase"
            maxPerInc = Application.WorksheetFunction.Max(ColumnRangeK)
            ws.Range("Q2").Value = maxPerInc
            tickerMatchRow = Application.WorksheetFunction.Match(maxPerInc, ColumnRangeK, 0)
            tickerMatch = Application.WorksheetFunction.Index(ColumnRangeI, tickerMatchRow)
            ws.Range("P2").Value = tickerMatch
            ws.Range("Q2").NumberFormat = "0.00%"
            
            ws.Range("O3").Value = "Greatest % Decrease"
            maxPerDec = Application.WorksheetFunction.Min(ColumnRangeK)
            ws.Range("Q3").Value = maxPerDec
            tickerMatchRow = Application.WorksheetFunction.Match(maxPerDec, ColumnRangeK, 0)
            tickerMatch = Application.WorksheetFunction.Index(ColumnRangeI, tickerMatchRow)
            ws.Range("P3").Value = tickerMatch
            ws.Range("Q3").NumberFormat = "0.00%"
            
            ws.Range("O4").Value = "Greatest Total Volume"
            maxTotVol = Application.WorksheetFunction.Max(ColumnRangeL)
            ws.Range("Q4").Value = maxTotVol
            tickerMatchRow = Application.WorksheetFunction.Match(maxTotVol, ColumnRangeL, 0)
            tickerMatch = Application.WorksheetFunction.Index(ColumnRangeI, tickerMatchRow)
            ws.Range("P4").Value = tickerMatch
            
           
            ws.Columns("O").AutoFit
            ws.Columns("P").AutoFit
            ws.Columns("Q").AutoFit
            
          ' Set variables/object expressions for conditional formatting based on positive and negative values
          
          
            Dim condition1J As FormatCondition, condition2J As FormatCondition
            Dim condition1K As FormatCondition, condition2K As FormatCondition
            
            Set condition1J = ColumnRangeJ.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
            With condition1J
            .Interior.ColorIndex = 4 ' green color index
            End With
            
            
            Set condition2J = ColumnRangeJ.FormatConditions.Add(xlCellValue, xlLess, "=-0")
            With condition2J
            .Interior.ColorIndex = 3 ' red color index
            End With
            
            Set condition1K = ColumnRangeK.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
            With condition1K
            .Interior.ColorIndex = 4 ' green color index
            End With
            
            
            Set condition2K = ColumnRangeK.FormatConditions.Add(xlCellValue, xlLess, "=-0")
            With condition2K
            .Interior.ColorIndex = 3 ' red color index
            End With

Next ws

End Sub








