Sub VBA_HW_StockMarketAnalysis()

    ' Loop through all the sheets in this workbook

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Heading
        Cells(1, "I").Value = "Ticker"

        Cells(1, "J").Value = "Total Stock Volume"
        
        ' Variable declarations
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TickerName As String
 
        Dim Volume As Double
        Volume = 0
        
        Dim Row As Double
        Row = 2
        
        Dim Column As Integer
        Column = 1
        
        Dim i As Long
        
        'Set Open Price
        OpenPrice = Cells(2, Column + 2).Value
        
         ' Looping through all ticker symbol
        
        For i = 2 To LastRow
        
        
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' Ticker name
                TickerName = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = TickerName
                
                ' Close Price
                ClosePrice = Cells(i, Column + 5).Value
                
 
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 9).Value = Volume
                
                ' Add one to the summary table row
                Row = Row + 1
                
                ' reset the Open Price
                
                Open_Price = Cells(i + 1, Column + 2)
                
                ' reset the Volume Total
                Volume = 0
                
            
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
  
        
    Next WS
        
End Sub

