Attribute VB_Name = "Module1"
Sub Alphabetical_testing()

'Creating Variables
Dim Ticker As String
Dim Ticker_Symbol As Double
Dim Yearly_Change As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double


'Looping through Sheets
For Each ws In Worksheets

'Appending Column names to sheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'finding all columns
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Ticker = ""
Ticker_Symbol = 0
Yearly_Change = 0
Percent_Change = 0
Total_Stock_Volume = 0

    For i = 2 To lastrow
        Ticker = Cells(i, 1).Value
        
        If Opening_Price = 0 Then
            Opening_Price = Cells(i, 3).Value
            End If
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
        If Cells(i + 1, 1).Value <> Ticker Then
        Ticker_Symbol = Ticker_Symbol + 1
        ws.Cells(Ticker_Symbol + 1, 9).Value = Ticker
        
        Closing_Price = Cells(i, 6).Value
        
        Yearly_Change = Closing_Price - Opening_Price
        ws.Cells(Ticker_Symbol + 1, 10).Value = Yearly_Change
        
        If Opening_Price = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = (Yearly_Change / Opening_Price)
            End If
            ws.Cells(Ticker_Symbol + 1, 11).Value = Format(Percent_Change, "Percent")
            
            ws.Cells(Ticker_Symbol + 1, 12).Value = Total_Stock_Volume
            
            Opening_Price = 0
            Total_Stock_Volume = 0
        
        End If
        
    Next i

'Set color formatting for Yearly Change
lastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    For j = 2 To lastrow
    If ws.Cells(j, 10) > 0 Then
    ws.Cells(j, 10).Interior.Color = vbGreen
    Else
    ws.Cells(j, 10).Interior.Color = vbRed
    End If
    Next j

Next ws

End Sub
