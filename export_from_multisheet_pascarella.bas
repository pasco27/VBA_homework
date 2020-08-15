Attribute VB_Name = "Module1"
Sub VBA_HW_multiyr()

Dim yearly_volume As Double
  
' Set an initial variable for holding each individual ticker
Dim ticker As String

' Set an initial variables for holding opening and closing prices
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_change As Double
Dim yearly_percent_change As Double
Dim lastRow As Long
Dim increase_percent As Double
Dim decrease_percent As Double
Dim increase_ticker As String
Dim decrease_ticker As String
Dim increase_volume As Double
Dim increase_volume_ticker As String
Dim Summary_Table_Row As Integer
Dim startrow As Long


For Each ws In Worksheets

    yearly_volume = 0
    Summary_Table_Row = 2
    increase_percent = 0
    decrease_percent = 0
    yearly_open = ws.Cells(2, 3).Value
    startrow = 2
     
' Set sheet counter to know it's the end of the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
' Loop through all stock records
    For i = 2 To lastRow
            
' Checker for the next stock ticker -- when we get to a ticker that is not the same, then we..
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Set the ticker stored for dataset
            ticker = ws.Cells(i, 1).Value
        
        ' State ticker yearly close, so at the end of the loop the last price is used.
            yearly_close = ws.Cells(i, 6).Value

        ' State yearly change
            yearly_change = yearly_close - yearly_open
            
        ' If 0 encountered in yearly open (new stock)
        ' make the yearly_open iterable within i
                    
                            
                            ' this will interate from starting point up to next new ticker
                            
                            For j = startrow To i
                    
                                    If ws.Cells(j, 3).Value <> 0 Then
                        
                                            yearly_open = ws.Cells(j, 3).Value

                                            Exit For
                                                    
                                    End If
                                    
                            Next j
                            
                            
                ' do the math for yearly percent change
                yearly_percent_change = ((ws.Cells(i, 6).Value - yearly_open) / yearly_open) * 100
                
                ' formatting to rid of all the decimal places
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00"
       
                ' Setup counter for yearly total volume
                yearly_volume = yearly_volume + ws.Cells(i, 7).Value
       

        ' Print the ticker in the Summary Table and its header
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Cells(1, 9).Value = "Ticker"
    
        ' Print the yearly change to the Summary Table and its header
            ws.Range("J" & Summary_Table_Row).Value = yearly_close - ws.Cells(startrow, 3)
            ws.Cells(1, 10) = "Yearly $ Change"
       
        ' Print the yearly percentage change to the Summary Table and its header
            ws.Range("K" & Summary_Table_Row).Value = yearly_percent_change
            ws.Cells(1, 11) = "Yearly % Change"
            
        ' Print the yearly total volume that has been looping to the Summary Table and its header
            ws.Range("L" & Summary_Table_Row).Value = yearly_volume
            ws.Cells(1, 12) = "Total Volume"
       
       
            startrow = i + 1
       
       
       

       
                ' Set colors for yearly % change - this is for the conditional formatting requirement
                If yearly_percent_change > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
           
                ElseIf yearly_perfect_change <= 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
           
                End If
       
       
       
        'This section is my attempt at the Challenge table
        
                If yearly_percent_change > increase_percent Then
                    increase_percent = yearly_percent_change
                    increase_ticker = ticker

                ElseIf yearly_percent_change < decrease_percent Then
                    decrease_percent = yearly_percent_change
                    decrease_ticker = ticker

                yearly_percent_change = 0
                
                ws.Cells(2, 16).Value = increase_percent
                ws.Cells(3, 16).Value = decrease_percent
                
                ws.Cells(2, 15).Value = ticker
                ws.Cells(3, 15).Value = ticker

                End If


                If yearly_volume > increase_volume Then
                    increase_volume = yearly_volume
                    increase_volume_ticker = ticker
                
                yearly_volume = 0
                
                ws.Cells(4, 16).Value = increase_volume
                ws.Cells(4, 15).Value = ticker
                
                End If
                   
            
        ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

        ' Reset the yearly volume
            yearly_volume = 0


            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"


   
   
    ' If the cell immediately following a row is the same ticker...
        Else

            yearly_volume = yearly_volume + ws.Cells(i, 7).Value
          
            
        End If
   
    Next i
 
  
Next ws

End Sub

