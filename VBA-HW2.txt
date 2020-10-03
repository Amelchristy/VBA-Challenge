Attribute VB_Name = "Module1"
Sub ticker_stock()

'Define

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim max_ticker As String
Dim min_ticker As String
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume_ticker As String
Dim max_volume As Double

max_volume = 0
max_ticker = " "
min_ticker = " "
max_percent = 0
min_percent = 0
max_volume_ticker = " "

'Avoid overflow error
On Error Resume Next

'Run through each worksheet
For Each ws In ThisWorkbook.Worksheets

'Titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
    
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
     
'setup integers for loop
Summary_Table_Row = 2

'put year open here before the loop

year_open = ws.Cells(2, 3).Value
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop

For i = 2 To Lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(Summary_Table_Row, 9).Value = ticker

'find values
            
ticker = ws.Cells(i, 1).Value
vol = vol + ws.Cells(i, 7).Value
ws.Range("L" & Summary_Table_Row).Value = vol
           
year_close = ws.Cells(i, 6).Value
yearly_change = year_close - year_open
ws.Cells(Summary_Table_Row, 10).Value = yearly_change
If (yearly_change > 0) Then
'Fill column green
 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf (yearly_change <= 0) Then
'Fill column red
ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If

              
percent_change = (year_close - year_open) / year_close
ws.Cells(Summary_Table_Row, 11).Value = percent_change
         
Summary_Table_Row = Summary_Table_Row + 1
vol = 0
     
'Challenge
   
If (percent_change > max_percent) Then
    max_percent = percent_change
    max_ticker = Ticker_Name
ElseIf (percent_change < min_percent) Then
    min_percent = percent_change
    min_ticker = Ticker_Name
    End If
                         
 If (Total_Ticker_Volume > max_volume) Then
    max_volume = Total_Ticker_Volume
    max_volume_ticker = Ticker_Name
    End If
                
'Reset Counters
percent_change = 0
Total_Ticker_Volume = 0
                

'If the cell immediately following a row is still the same ticker name,add to Total Ticker Volume
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            End If
      
        Next i

            'Check if it is not the first spreadsheet
            'Print the Values
            
            If Not COMMAND_SPREADSHEET Then
            
               ws.Range("Q2").Value = (CStr(max_percent) & "%")
               ws.Range("Q3").Value = (CStr(min_percent) & "%")
               ws.Range("P2").Value = MAX_TICKER_NAME
               ws.Range("P3").Value = MIN_TICKER_NAME
               ws.Range("Q4").Value = max_volume
               ws.Range("P4").Value = max_volume_ticker
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next ws
        
    
ws.Columns("K").NumberFormat = "0.00%"


End Sub
