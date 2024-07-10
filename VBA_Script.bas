Attribute VB_Name = "Module1"
Sub Multiple_Year_stock_data()
Dim Current As Worksheet
  For Each Current In ThisWorkbook.Worksheets
  Current.Activate

    ' Set an initial variable for holding the ticker name
    Dim ticker_name As String

    ' Set an initial variable for holding the total per ticker
    Dim total_volume As Double
    total_volume = 0

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    ' Set the header of summary table
    Range("I1") = "Ticker"
    Range("J1") = "Quarterly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % Increase Value"
    Range("O3") = "Greatest % Decrease Value"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
  
    'Set the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Loop through all daily records
       For i = 2 To lastrow
        'If previous ticker and current ticker are not the same, then...
        If Cells(i - 1, 1) <> Cells(i, 1) Then
        opening_price = Cells(i, 3)
        
        'If next ticker and current ticker are not the same, then...
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the ticker name
        ticker_name = Cells(i, 1).Value

        ' Add to the total volume
        total_volume = total_volume + Cells(i, 7).Value
      
        'Set the closing price
        closing_price = Cells(i, 6).Value
      
        'Calculate the change between opening price and closing price
        quarterly_change = closing_price - opening_price
      
        'Calculate the percent change between opening price and closing price
        'If opening_price = 0 Then
        'percentage_change = Null
        percent_change = ((closing_price - opening_price) / opening_price)
        On Error Resume Next

        ' Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker_name
      
        ' Print the quarterly change to the Summary Table
        Range("J" & Summary_Table_Row).Value = quarterly_change
      
        ' Print the percent change to the Summary Table
        Range("K" & Summary_Table_Row).Value = percent_change
        Columns("K:K").NumberFormat = "0.00%"

        ' Print the ticker Amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = total_volume
      
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the total volume
        total_volume = 0

        ' If ticker is the same
        Else
        'Add to the ticker total
        total_volume = total_volume + Cells(i, 7).Value

        End If
             
       Next i
        
        'After the 1st loop is done, set the next loop
       Dim greatest_increase, greatest_decrease As Double
       greatest_increase = Cells(2, 11)
       greatest_decrease = Cells(2, 11)
       greatest_volume = Cells(2, 12)
       lastrow_summary = Cells(Rows.Count, 10).End(xlUp).Row
  
       For j = 2 To lastrow_summary
        'Change the format depending on the value
        If Cells(j, 10) >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
  
        ElseIf Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
   
        End If
        
        'Loop through each row and replace the greatest increase value
        If Cells(j, 11) > greatest_increase Then
        greatest_increase = Cells(j, 11)
        Cells(2, 17) = greatest_increase
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16) = Cells(j, 9)
   
        End If
   
        'Loop through each row and replace the greatest decrease value
        If Cells(j, 11) < greatest_decrease Then
        greatest_decrease = Cells(j, 11)
        Cells(3, 17) = greatest_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16) = Cells(j, 9)
  
        End If
        
        'Loop through each row and replace the greatest total volume
        If Cells(j, 12) > greatest_volume Then
        greatest_volume = Cells(j, 12)
        Cells(4, 17) = greatest_volume
        Cells(4, 16) = Cells(j, 9)
   
        End If
   
       Next j
 
        Columns("I:Q").AutoFit
 Next

End Sub
