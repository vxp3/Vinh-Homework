VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub sheet()

'Declare This work sheet as ws for worksheet

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets

'add ws. in front cell and range for it to loop all worksheet

'declare variable names

Dim Percent_change As Double
Dim Price_open As Double
Dim Price_close As Double
Dim Ticker_name As String
Dim Total_stock_volume As Double
Dim Yearly_change As Double

     Total_stock_volume = 0
        
' Determine the Last Row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long

'set cell value name as ticker, yearly change, percent change, and total stock volume in
            ws.Range("I1").Value = "ticker"
            ws.Range("J1").Value = "yearly change"
            ws.Range("K1").Value = "percent change"
            ws.Range("L1").Value = "total stock volume"

'auto fit columns

ws.Range("I1:L1").Columns.AutoFit

'set a starting value of price open and price close

Price_open = ws.Cells(2, 3).Value

'start looping in each worksheet for ticker from start to end

Row = 2

For i = 2 To Lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_name = ws.Cells(i, 1).Value

        'yearly_change = price_close - price_open
        
        Price_close = ws.Cells(i, 6).Value

        Yearly_change = (ws.Cells(i, 6).Value) - ((ws.Cells(2, 3).Value))
        
        'Add total stock volume
        
        Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
        
        ws.Cells(i, 12).Value = Total_stock_volume

             
        'calculate % and add % to column K
        
        If Price_open <> 0 Then
            yearly_percent = ((Yearly_change / ws.Cells(2, 3).Value))
            
            ws.Range("k2" & Row).NumberFormat = "0.00%"
                     
        End If
        
        'add color into yearly change, Green is good and Red is bad
        
        If yearly_percent > 0 Then
            
            ws.Range("J" & Row).Interior.ColorIndex = 4
        
        ElseIf yearly_percent <= 0 Then
            
            ws.Range("J" & Row).Interior.ColorIndex = 3
        
        End If
        
    'store info into cells
    
        ws.Cells(Row, 9).Value = Ticker_name
        ws.Cells(Row, 10).Value = Yearly_change
        ws.Cells(Row, 11).Value = yearly_percent
        ws.Cells(Row, 12).Value = Total_stock_volume
    
        Row = Row + 1
             
        Price_open = ws.Cells(i + 1, 3)
    
        Else
        
        Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
        
        End If
        
    Next i


 Next ws


End Sub

