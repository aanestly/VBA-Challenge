Attribute VB_Name = "Module3"
Sub Stock_Data()
For Each ws In Worksheets

Dim WorksheetName As String
Dim ticker As String
Dim i As Long
Dim j As Long
Dim TickCount As Long

'Last row column A
Dim LastRowA As Long

'last row column I
Dim LastRowI As Long

'percent change double
 Dim PerChange As Double
 
'greatest % increase
Dim GreatPerIncr As Double

'greatest % decrease
Dim GreatPerDecr As Double

'greatest total volume
Dim GreatVol As Double
        
'Get the WorksheetName
WorksheetName = ws.Name
        
'set tick_counter to first row
TickCount = 2

'start row at 2
 j = 2
        
'last cell in column A
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
'Loop through all rows
For i = 2 To LastRowA
            
'Nsme of ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
    'name of ticker in column I
    ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
    'yearly change calculation
     ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
If ws.Cells(TickCount, 10).Value < 0 Then
                
    'cell background color red
     ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
          Else
          'cell background color green
           ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
'percent change in column K
If ws.Cells(j, 3).Value <> 0 Then
    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
     'Percent formating
      ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
           Else
            ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                 End If
                    
'total volume in column L
ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

   TickCount = TickCount + 1
                
     'Set new start row of the ticker block
      j = i + 1
                
                 End If
            
            Next i
            
'last cell in Column I
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
'Loop for summary
For i = 2 To LastRowI
      
If ws.Cells(i, 12).Value > GreatVol Then
    GreatVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
        Else
        GreatVol = GreatVol
                
                End If
                
If ws.Cells(i, 11).Value > GreatIncr Then
    GreatIncr = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
             Else
             GreatIncr = GreatIncr
                    
                    End If
                    
If ws.Cells(i, 11).Value < GreatDecr Then
    GreatDecr = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                End If
                
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
    Next i
    Next ws
End Sub
