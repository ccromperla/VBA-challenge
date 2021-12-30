Attribute VB_Name = "Module1"
Sub Stock_Ticker_For_Each():

 'Set the dimensions
 Dim Ticker As String
 Dim Opening As Double
 Dim Closing As Double
 Dim Percent_Change As Double
 Dim Yearly_Change As Double
 Dim Row_Table As Long
 Dim Last_Row As Long
 Dim Stock_Volume As Double
 Dim ws As Worksheet

For Each ws In Worksheets
 
 'Assign values
 Opening = 0
 Closing = 0
 Row_Table = 2
 Stock_Volume = 0
 Yearly_Change = 0
 Percent_Change = 0
 
 'Determine the last row for i loop
 Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 'Create Headers for the Table
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Stock Volume"
 
 
 'Start for loop
 For i = 2 To Last_Row
  
  'Identify the first opening value and store it for..later
  If Opening = 0 Then
   Opening = ws.Cells(i, 3).Value
  End If
  
  'Check it the value in next cell is the same, if not..
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
   'Ticker Value id'd and added to the ticker
    Ticker = ws.Cells(i, 1).Value
    ws.Range("I" & Row_Table).Value = Ticker
   
    'Identify the closing value
    Closing = ws.Cells(i, 6).Value
   
   'Calculate Yearly Change and Add these values to the ticker
    Yearly_Change = Closing - Opening
    ws.Range("J" & Row_Table).Value = Yearly_Change
    ws.Range("J" & Row_Table).NumberFormat = "$0.00"
   
   'Format Yearly Change so the increase or same value is green and decrease in value is red
     If Closing > Opening Then
     ws.Range("J" & Row_Table).Interior.ColorIndex = 50
     ElseIf Closing < Opening Then
      ws.Range("J" & Row_Table).Interior.ColorIndex = 53
     Else: ws.Range("J" & Row_Table).Interior.ColorIndex = 44
     End If
    
    'Calculate Percent Change add it to the ticker and format it to a percent
    'Dividing by a 0 value breaks the code so need to account for a 0 opening value (!!!!!!!)
    If Opening = 0 Then
     Percent_Change = 0
    Else
     Percent_Change = (Yearly_Change / Opening)
    End If
    ws.Range("K" & Row_Table).Value = Percent_Change
    ws.Range("K" & Row_Table).NumberFormat = "0.00%"
   
    'Calculate Stock Volume and Add total volume to the ticker
    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    ws.Range("L" & Row_Table).Value = Stock_Volume
    
    'Add 1 to the Row table for the next value
    Row_Table = Row_Table + 1
    
    'Reset Values
    Opening = 0
    Closing = 0
    Percent_Change = 0
    Yearly_Change = 0
    Stock_Volume = 0
    
   'If values aren't different then add Stock_Volume
   Else: Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
  
  'End if
  End If
   
 'Moving on to the next i until we hit the last row
 Next i

'Move on to the next worksheet
Next ws
  
End Sub
