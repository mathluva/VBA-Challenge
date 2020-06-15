Option Compare Text

Sub VBAChallenge()

    Dim i, x, j, k, l, m, lastrow, count As Long
    
    Dim openvalue, closevalue, volumetotal, opencounter, percentmax, percentmin, volumemax As Double

    Dim name As String
    
    Dim mainworkbook As Workbook

    Set mainworkbook = ActiveWorkbook
    
             count = 2
        
              x = 2
        
             opencounter = 2

              lastrow = Cells(Rows.count, 1).End(xlUp).Row
        
    For i = 2 To lastrow
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                name = Cells(i, 1).Value   'Grab the name of first ticker once they become different
                
                openvalue = Cells(opencounter, 3).Value
              
                Range("K" & count).Value = name  'Assign ticker to column K
                
                count = count + 1 'counter to put ticker names in column k, starting with row 2
                
                volumetotal = volumetotal + Cells(i, 7).Value  'volume total for each ticker
            
                        'assign oppening value for the year corresponding to ticker in column L
            
               
                Range("N" & x).Value = volumetotal
                
                Range("N" & x).NumberFormat = "0"
            
             'Yearly change from opening price at the beginning of year to the closing price at the end of the year
             
                Range("L" & x).Value = closevalue - openvalue
                    
                        If Range("L" & x).Value < 0 Then
                             Range("L" & x).Interior.ColorIndex = 3
                             
                             Else
                              Range("L" & x).Interior.ColorIndex = 4
                        End If
                        
             'Percent change from opening price at the beginning of the year to closing price at the end of the year
             
                    If openvalue <> 0 Then
                    
                         Range("M" & x).Value = ((closevalue - openvalue) / openvalue)
                    
                    Else
                        
                        Range("M" & x).Value = (closevalue - openvalue)
                        
                    End If
           
                    x = x + 1 'increase summary counter to mover to next row
                    
                    volumetotal = 0
                    
                    opencounter = i + 1 'as soon as it finds where the names change, that's the i value which needs to increase by 1 to capture the next name
                    
        Else
               
                closevalue = Cells(i + 1, 6).Value
         
                volumetotal = volumetotal + Cells(i, 7).Value
        
    End If
                
        Next i
        
        ' Column Headers
        
                    Range("K1").Value = "Ticker Name"
        
                    Range("L1").Value = "Yearly Change"
        
                    Range("M1").Value = "Percent Change"
                    
                    Range("N1").Value = "Total Stock Volume"
                    
         percentmax = Application.WorksheetFunction.Max(Range("M:M"))
         
         percentmin = Application.WorksheetFunction.Min(Range("M:M"))
         
         volumemax = Application.WorksheetFunction.Max(Range("N:N"))
         
    For j = 2 To lastrow
        
          If Cells(j, 13).Value = percentmax Then
            
                Range("Q2") = Cells(j, 11).Value
            
                Range("R2").Value = percentmax
             
                Range("P2").Value = "Greastest % Increase"
                
            End If
            
            Next j
            
    
     For k = 2 To lastrow
                
         If Cells(k, 13).Value = percentmin Then
            
                   Range("Q3") = Cells(k, 11).Value
            
                   Range("R3").Value = percentmin
            
               Range("P3").Value = "Greatest % Decrease"
                
           End If
            
           Next k
            
                
    For l = 2 To lastrow
        
        If Cells(l, 14).Value = volumemax Then
            
           Range("Q4").Value = Cells(l, 11).Value
            
           Range("R4").Value = volumemax
            
            Range("P4").Value = "Greatest Total Volume"
            
          End If
          
         Next l
        
             Range("M:M").NumberFormat = "0.00%"
    
             Range("R2").NumberFormat = "0.00%"
    
             Range("R3").NumberFormat = "0.00%"
    
            Range("R4").NumberFormat = "0"
    
            Range("Q1").Value = "Ticker"
    
            Range("R1").Value = "Value"
        

End Sub
