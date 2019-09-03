Attribute VB_Name = "Module1"

       Sub Easy_and_Challenge_Homework()
       
       'Loop thru all the 3 sheets(2014,2015,2016) in this Workbook
       
       For Each ws In Worksheets
       
       'Insert the labels for Ticker and Total Stock Volume in designated cells
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"

       'Go thru Column A (<ticker>) in all 3 sheets and find each identical<ticker> and insert in Column I
       'Go thru Column G (<vol>) in all 3sheets and add up the volume for each corresponding ticker and insert in Column J
       'Define our variables
       

        Dim LastRow As Long
        Dim i As Long
        Dim ticker_location As Integer
        Dim volume As Double
        ticker_location = 2
        volume = 0
        
        'Insert formula from class/stack overflow
        
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set each symbol
                ws.Range("I" & ticker_location).Value = ws.Range("A" & i).Value

            ' Add the last row to total volume
                volume = volume + ws.Range("G" & i).Value
            
            ' Total volume in column J that corresponds with its ticker symbol
                ws.Range("J" & ticker_location).Value = volume
            
            ' Add to the symbol and volume location
                ticker_location = ticker_location + 1

            ' Reset the volume
                volume = 0

            ' If the cell immediately following a row is the same stock
            Else
                
                ' Add to the volume for the stock symbol
                volume = volume + ws.Range("G" & i).Value
               

            End If
        
        Next i

  
    Next ws
MsgBox ("Changes Applied")

End Sub
    
    
   
        
        
        
        
      
      
      
      
      
      
 
      

