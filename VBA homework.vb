Sub GenerateStockResults()

  ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' INSERT THE STATE
        ' --------------------------------------------

       
        Dim WorksheetName As String
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name

        ' headers
        ws.Cells(1, 8).Value = " "
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim Maxopen As Double
        Dim Minopen As Double
        Dim Ticker  As String
        Dim YearlyChange As Double
        Dim PerChange As Double
        Dim TotalStVol As Double
        
       
        Minopen = Cells(2, 3).Value
        Maxopen = Cells(2, 3).Value
        
        Ticker = ""
        YearlyChange = 0
        PerChange = 0
        TotalStVol = 0
        ResultIndex = 2
        
        For i = 2 To LastRow
         
        
         'check for Ticker if we are looping through same Ticker
         If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            If Cells(i + 1, 2).Value > Cells(i, 2).Value Then
                Maxopen = Cells(i + 1, 3).Value
                TotalStVol = TotalStVol + Cells(i, 7).Value
            Else
                Maxopen = Cells(i, 3).Value
                TotalStVol = TotalStVol + Cells(i, 7).Value
            End If
           
        Else
          TotalStVol = TotalStVol + Cells(i, 7).Value
          Cells(ResultIndex, 9).Value = Cells(i, 1).Value
          Cells(ResultIndex, 10).Value = Maxopen - Minopen
          If Minopen > 0 Then
             Cells(ResultIndex, 11).Value = FormatPercent((Maxopen - Minopen) / Minopen)
          Else
          Cells(ResultIndex, 11).Value = FormatPercent(0)
          End If
          
          Cells(ResultIndex, 12).Value = TotalStVol
          If (Minopen) > 0 Then
           If ((Maxopen - Minopen) / Minopen) < 0 Then
              Cells(ResultIndex, 11).Interior.ColorIndex = 3 ' Red
           Else
            Cells(ResultIndex, 11).Interior.ColorIndex = 4 ' geen
           End If
           
           Else
            Cells(ResultIndex, 11).Interior.ColorIndex = 4 ' geen
            End If
            
          
          
          
          Minopen = Cells(i + 1, 3).Value
          Maxopen = 0
          ResultIndex = ResultIndex + 1
          TotalStVol = 0
        End If
        Next i
        
      '  MsgBox (Cells(i - 1, 1).Value)
       ' MsgBox ("Min " + Str(Minopen))
        'MsgBox ("max" + Str(Maxopen))
     
    Next ws

   ' MsgBox ("Fixes Complete")
End Sub
