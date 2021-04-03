Attribute VB_Name = "Module1"
Sub stocks_loop_sheet_2016()
    
    Dim tickername As String
    Dim yearopen As Double
    Dim yearclose As Double
    Dim percentagechange As Double
    
    'yearly change info
    Dim yearlychange As Double
    yearlychange = 0
    
    
    'total volume info
    Dim totalvolume As Double
    totalvolume = 0
    
    'summary table info
    Dim summarytable As Integer
    summarytable = 2
    

    'set last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row


                'create headers
                Cells(1, 10).Value = "Ticker"
                Cells(1, 11).Value = "Yearly Change"
                Cells(1, 12).Value = "Percent Change"
                Cells(1, 13).Value = "Total Stock Volume"
        
                'loop through all rows
                For i = 2 To LastRow
        
                
      
            'check current cell with the cell below it
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'set the ticker name
                tickername = Cells(i, 1).Value
            
                'print the ticker name in ticker column
                Range("J" & summarytable).Value = tickername
                
                'determine yearly change
                yearopen = Cells(i - 261, 3).Value
                yearclose = Cells(i, 6).Value
                yearlychange = yearclose - yearopen
                Range("K" & summarytable).Value = yearlychange
                
                'determine percentagechange
                percentagechange = yearlychange / yearopen
                Range("L" & summarytable).Value = percentagechange
                
                                

                
                
                 'formating percent in L column
                Range("L" & summarytable).NumberFormat = "0.00%"

                'add to volume total
                totalvolume = totalvolume + Cells(i, 7).Value
            
                'print total volume in total volume column
                Range("M" & summarytable).Value = totalvolume
                
               

                'add one to the summarytable row
                summarytable = summarytable + 1
            
                totalvolume = 0
                
        
            
            'if the cell immediately following a row is the same brand
            Else
                
                'add to the volume total
                totalvolume = totalvolume + Cells(i, 7).Value
                
                

            
            End If
            
                
               Next i
               
                       
        LastRow = Cells(Rows.Count, "L").End(xlUp).Row
        
        For i = 2 To LastRow
        
                    
            'conditional formatting
            If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
                
           Else
           
                
                Cells(i, 11).Interior.ColorIndex = 3
                
                End If
                
                Next i
                
   
End Sub


Sub stocks_loop_2015_and_2014()
    
    Dim tickername As String
    Dim yearopen As Double
    Dim yearclose As Double
    Dim percentagechange As Double
    
    'yearly change info
    Dim yearlychange As Double
    yearlychange = 0
    
    
    'total volume info
    Dim totalvolume As Double
    totalvolume = 0
    
    'summary table info
    Dim summarytable As Integer
    summarytable = 2
    

    'set last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row


                'create headers
                Cells(1, 10).Value = "Ticker"
                Cells(1, 11).Value = "Yearly Change"
                Cells(1, 12).Value = "Percent Change"
                Cells(1, 13).Value = "Total Stock Volume"
        
                'loop through all rows
                For i = 2 To LastRow
        
                
      
            'check current cell with the cell below it
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'set the ticker name
                tickername = Cells(i, 1).Value
            
                'print the ticker name in ticker column
                Range("J" & summarytable).Value = tickername
                
                'determine yearly change
                yearopen = Cells(i - 260, 3).Value
                yearclose = Cells(i, 6).Value
                yearlychange = yearclose - yearopen
                Range("K" & summarytable).Value = yearlychange
                
                'determine percentagechange
                If percentagechange <> 0 Then
                percentagechange = yearlychange / yearopen
                Range("L" & summarytable).Value = percentagechange
                End If
                
                
                'formating percent in L column
                Range("L" & summarytable).NumberFormat = "0.00%"

                

                'add to volume total
                totalvolume = totalvolume + Cells(i, 7).Value
            
                'print total volume in total volume column
                Range("M" & summarytable).Value = totalvolume
                
               

                'add one to the summarytable row
                summarytable = summarytable + 1
            
                totalvolume = 0
                
        
            
            'if the cell immediately following a row is the same brand
            Else
                
                'add to the volume total
                totalvolume = totalvolume + Cells(i, 7).Value
            


     
        End If
        

        Next i
        
        LastRow = Cells(Rows.Count, "L").End(xlUp).Row
        
        For i = 2 To LastRow
        
                    
            'conditional formatting
            If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
                
           Else
           
                
                Cells(i, 11).Interior.ColorIndex = 3
                
                End If
        
        Next i
        
     

   
End Sub





