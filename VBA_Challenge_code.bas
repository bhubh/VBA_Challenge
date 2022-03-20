Attribute VB_Name = "Module1"
Sub Stocks():



    For Each ws In Worksheets
    
        'Set a variable representing the stock ticker
        Dim Ticker As String
        
        'Setting track of Ticker_Symbol_row in summary table
        Dim Ticker_Symbol_Row As Integer
        
        Ticker_Symbol_Row = 2    'starting row in summary table
        
        'Setting variable for representing stock  opening and closing price
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Price_Change As Double
        Dim Percent_Change As Double
        
        'Setting variable representing total volume for particular ticker
        Dim Total_Volume As Double
        Total_Volume = 0         'Starting point for volume variable.
        
        Open_Price = ws.Cells(2, 3).Value   'Open price for first ticker
        
        Dim I As Integer
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'summary table columns
        ws.Cells(Ticker_Symbol_Row - 1, 12).Value = "Ticker"               'ticker column header in summary table
        ws.Cells(Ticker_Symbol_Row - 1, 13).Value = "Yearly change"        'Yearly cahnge column header in summary table
        ws.Cells(Ticker_Symbol_Row - 1, 14).Value = "% Change"             'Percentage change column header
        ws.Cells(Ticker_Symbol_Row - 1, 15).Value = "Total Volume"          'toal volume column header
        
    
    
    
    
            'Loop through all the ticker symbols
            For I = 2 To RowCount
            
                If (ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value) Then
            
                    'Set the ticker symbol
            
                    Ticker = ws.Cells(I, 1).Value
             
                      
                    Close_Price = ws.Cells(I, 6).Value                  'Close price for previous ticker
             
                    Price_Change = Close_Price - Open_Price           'Yearly price change
                    Percent_Change = 100 * (Price_Change / Open_Price)
                    
                    Percent_Change = Round(Percent_Change, 2)          'Round to two decimal places
                    Total_Volume = Total_Volume + ws.Cells(I, 7)
             
                    'Printing ticker_symbol, Price change and total volume in summary table
                    ws.Cells(Ticker_Symbol_Row, 12).Value = Ticker         ' Ticker symbol
                    ws.Cells(Ticker_Symbol_Row, 13).Value = Price_Change
                    ws.Cells(Ticker_Symbol_Row, 14).Value = Percent_Change
                    ws.Cells(Ticker_Symbol_Row, 15).Value = Total_Volume
                
                        If Price_Change >= 0 Then
                            ws.Cells(Ticker_Symbol_Row, 13).Interior.ColorIndex = 4
                            ws.Cells(Ticker_Symbol_Row, 14).Interior.ColorIndex = 4
                   
                        Else
                            ws.Cells(Ticker_Symbol_Row, 13).Interior.ColorIndex = 3
                            ws.Cells(Ticker_Symbol_Row, 14).Interior.ColorIndex = 3
                        End If
                
                    Total_Volume = 0                             'Reset total volume for new ticker
                    Open_Price = ws.Cells(I + 1, 3).Value            'Reset open price for new ticker
                    Ticker_Symbol_Row = Ticker_Symbol_Row + 1      'Next row for next ticker in summary table
                
                 Else
                
                    Total_Volume = Total_Volume + ws.Cells(I, 7)
               
                End If
            
            
            
            Next I
    Next ws


End Sub
