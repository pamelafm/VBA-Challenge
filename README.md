# VBA-Challenge
Sub stockdatasummary()
  Dim Current As Worksheet
  For Each Current In ThisWorkbook.Worksheets
  Current.Activate

    ' Set an initial variable for holding the ticker name

    ' Set an initial variable for holding the total per ticker

    ' Keep track of the location for each ticker in the summary table
  
    ' Set the header of summary table
   
    'Set the last row
      
    ' Loop through all daily records
       
               'If previous ticker and current ticker are not the same, then...
                
        'If next ticker and current ticker are not the same, then...
        
        ' Set the ticker name
        
        ' Add to the total volume
              
        'Set the closing price
              
        'Calculate the change between opening price and closing price
              
        'Calculate the percent change between opening price and closing price
        'If opening_price = 0 Then
        'percentage_change = Null
        
        ' Print the ticker in the Summary Table
              
        ' Print the quarterly change to the Summary Table
              
        ' Print the quarterly change to the Summary Table
        
        ' Print the ticker Amount to the Summary Table
              
        ' Add one to the summary table row
             
        ' Reset the total volume
        
        ' If ticker is the same
               'Add to the ticker total
        
              
        'After the 1st loop is done, set the next loop
       
  
               'Change the format depending on the value
          
                'Loop through each row and replace the greatest increase value
      
        
        'Loop through each row and replace the greatest decrease value
  
             
        'Loop through each row and replace the greatest total volume
        
      
 
  End Sub
