Attribute VB_Name = "Module1"
Sub VBA_HW_r0()

' Set an initial variable for holding each individual ticker
  Dim ticker As String

' Set an initial variables for holding opening and closing prices
  Dim yearly_open As Double
  Dim yearly_close As Double
  Dim yearly_change As Double
  Dim yearly_percent_change As Double
    Dim increase_percent As Double
    Dim decrease_percent As Double
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim increase_volume As Double
    Dim increase_volume_ticker As String
    
 
' Set an initial variables for holding the total volume
  Dim yearly_volume As Double
  yearly_volume = 0
 
  increase_percent = 0
  decrease_perecent = 0
  

' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
  yearly_open = Cells(2, 3).Value
 
' Loop through all stock records
  For i = 2 To 70926
   
' Checker for the next stock ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set ticker as item to be checked/stored
        ticker = Cells(i, 1).Value
       
        ' State ticker yearly open to be pulled - is the correct way to pull the last?
        yealy_open = Cells(i + 1, 3).Value
         
        ' State ticker yearly close to be pulled --- how do Iick particularly the last one?
        yearly_close = Cells(i, 6).Value

        ' State yearly change
        yearly_change = yearly_close - yearly_open
       
        ' do the math for yearly percent change
        yearly_percent_change = (yearly_change / yearly_open) * 100
       
       ' formatting to rid of all the decimal places
        Range("K" & Summary_Table_Row).NumberFormat = "0.00"
       
        ' Setup counter for yearly total volume
        yearly_volume = yearly_volume + Cells(i, 7).Value
       

        ' Print the ticker in the Summary Table and its header
        Range("I" & Summary_Table_Row).Value = ticker
        Cells(1, 9).Value = "Ticker"

        ' Print the yearly change to the Summary Table and its header
        Range("J" & Summary_Table_Row).Value = yearly_change
        Cells(1, 10) = "Yearly $ Change"
       
        ' Print the yearly percentage change to the Summary Table and its header
        Range("K" & Summary_Table_Row).Value = yearly_percent_change
        Cells(1, 11) = "Yearly % Change"
       
                ' Set colors for yearly % change - this is for the conditional formatting requirement
                If yearly_percent_change > 0 Then
                    Range("K" & Summary_Table_Row).Interior.Color = vbGreen
           
                ElseIf yearly_perfect_change <= 0 Then
                    Range("K" & Summary_Table_Row).Interior.Color = vbRed
           
                End If
                
                
                'Run the If's to populate the Challenge table
                If yearly_percent_change > increase_percent Then
                    increase_percent = yearly_percent_change
                    increase_ticker = ticker

                    Cells(2, 16).Value = increase_percent
                    Cells(2, 15).Value = ticker

                ElseIf yearly_percent_change < decrease_percent Then
                    decrease_percent = yearly_percent_change
                    decrease_ticker = ticker

                    Cells(3, 16).Value = decrease_percent
                    Cells(3, 15).Value = ticker

                End If

                If yearly_volume > increase_volume Then
                    increase_volume = yearly_volume
                    increase_volume_ticker = ticker
            
                
                Cells(4, 16).Value = increase_volume
                Cells(4, 15).Value = ticker
                
                End If
                
                
        ' Print the yearly total volume to the Summary Table and its header
        Range("L" & Summary_Table_Row).Value = yearly_volume
        Cells(1, 12) = "Total Volume"

               
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1

        ' Reset the yearly volume
        yearly_volume = 0
        
        ' Label the Challenge Table
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
  
    ' If the cell immediately following a row is the same ticker...
        Else

            yearly_volume = yearly_volume + Cells(i, 7).Value
                 
        End If

   Next i
   
End Sub



