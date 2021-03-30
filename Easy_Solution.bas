Attribute VB_Name = "Module1"
Sub Testrun()

      
    'Dim ws As Worksheet
    Dim i As Long
    Dim total As Double
    Dim counter As Double
    Dim LastRow As Long
    Dim ws As Worksheet
   
    For Each ws In ActiveWindow.SelectedSheets
    'Name column to keep the data
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Total Stock Volume"
    
    'create unique list from unsorted column. only works on sorted data
        total = 0
        counter = 2
        
    ' Find LastRow in column A
        
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        
    ' Sum TotalStockVolume based on the 1st ticker starting from row#2
        For i = 2 To LastRow
            
                          
            total = total + Cells(i, 7).Value
                    
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            'input Ticker name
                    Cells(counter, 9).Value = Cells(i, 1).Value
                        
            'input totalStockVolume that summarised
                    Cells(counter, 10).Value = total
            
            'reset totalStockVolume to 0 and then count again from the next ticker
                    total = 0
            
            'increase counter
                    counter = counter + 1
            End If
        
        Next i
           
    
    Next ws

End Sub
