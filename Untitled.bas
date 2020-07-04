Attribute VB_Name = "Module1"

Sub taskone():
For Each ws In Worksheets
   'Set Loop Variables
   
    Dim ticker As String
    Dim Summary_Table_Row As Integer
    Dim openp As Double
    Dim closep As Double
    Dim OpenRowTracker As Long
    Dim yoym As Double
    Dim yoyp As Double
    Dim volume As LongLong
    Dim RowCount As Long
    Dim RowCount2 As Long
   
    
    
    
    
    'Create New Table headers
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "YoY CHG $"
    ws.Range("L1").Value = "YoY CHG %"
    ws.Range("M1").Value = "Yearly Volume"
    
    'Set Indexes for Row  Tracking
    Summary_Table_Row = 2
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        OpenRowTracker = 2
        For i = 2 To RowCount
               
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                    
                    
                    'Calculate the YoY Cahnge in $ and %
                    
                    openp = ws.Cells(OpenRowTracker, 3).Value
                    closep = ws.Cells(i, 6).Value
                    yoym = closep - openp
                    If openp > 0 Then
                    yoyp = (closep - openp) / openp
                    Else
                    yoyp = 0
                    End If
                    
                    
                    ws.Range("K" & Summary_Table_Row).Value = yoym
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "$0.00"
                    ws.Range("L" & Summary_Table_Row).Value = yoyp
                    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    OpenRowTracker = i
                
                    
                    'Print Values
                    
                    ticker = ws.Cells(i, 1).Value
                    Summary_Table_Row = Summary_Table_Row + 1
                    ws.Range("J" & Summary_Table_Row - 1).Value = ticker
                    
                    'Calculate volume
                    
                    volume = volume + ws.Cells(i, 7).Value
                    ws.Range("M" & Summary_Table_Row - 1).Value = volume
                    volume = 0
                  Else
                    volume = volume + ws.Cells(i, 7).Value
                    
                End If
            Next i
                
                 RowCount2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
                   
                For j = 2 To RowCount2
                
                    If ws.Cells(j, 11).Value > 0 Then
                        ws.Cells(j, 11).Interior.ColorIndex = 4
                    Else
                        ws.Cells(j, 11).Interior.ColorIndex = 3
                    End If
   
                Next j
                
            ws.Cells(1, 16).Value = "Greatest increase %"
            ws.Cells(2, 16).Value = "Greatest decrease %"
            ws.Cells(3, 16).Value = "Greatest Yearly Volume"
            
            ws.Cells(1, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
            ws.Cells(1, 17).NumberFormat = "0.00%"
            ws.Cells(2, 17).Value = WorksheetFunction.Min(ws.Range("L:L"))
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = WorksheetFunction.Max(ws.Range("M:M"))
        
         
        Next ws
      
    
End Sub
