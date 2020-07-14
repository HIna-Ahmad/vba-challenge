Attribute VB_Name = "Module1"


Sub GenerateSummary():

    Dim lastRow As Long
    Dim writeCellRow As Integer
    Dim Year_Change As Variant
    Dim Start_Price As Variant
    Dim End_Price As Variant
    Dim Percent_Change As Variant
    Dim Total_Volume As Variant

    
        
       
   'To insert column title
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Year Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
 

   writeCellRow = 1
   lastRow = Cells(Rows.Count, "A").End(xlUp).Row
   Total_Volume = 0
    
    
    For Row = 1 To lastRow
        'Checking for change in stock symbol
        
        
        If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
            'To populate ticker column
            Cells(writeCellRow + 1, 9).Value = Cells(Row + 1, 1).Value
            
            
            'To populate Total_Volume
            If Row > 1 Then
                'Adding Volume from last stock
                Total_Volume = Total_Volume + Cells(Row, 7).Value
                Cells(writeCellRow, 12).Value = Total_Volume
                Total_Volume = 0
            End If
            
            
            'Populating Year_Change and Percent_Change
            If Start_Price Then
                End_Price = Cells(Row, 6)
                Year_Change = Start_Price - End_Price
                Cells(writeCellRow, 10) = Year_Change
                
                If Year_Change > 0 Then
                    Cells(writeCellRow, 10).Interior.ColorIndex = 4
                Else
                    Cells(writeCellRow, 10).Interior.ColorIndex = 3
                End If
                
                Percent_Change = Format((Year_Change / Start_Price), "Percent")
                Cells(writeCellRow, 11) = Percent_Change
            End If
            
            'Updating open price
            Start_Price = Cells(Row + 1, 3)
            'Updating print row
            writeCellRow = writeCellRow + 1
            
        Else
            Total_Volume = Total_Volume + Cells(Row, 7).Value
        End If
                
    Next Row
    
End Sub

Sub LoopWorksheets():

    Dim sht As Worksheet
    
    For Each sht In Worksheets
        sht.Select
        Call GenerateSummary
        
        MsgBox sht.Name
    
    Next
        
End Sub

