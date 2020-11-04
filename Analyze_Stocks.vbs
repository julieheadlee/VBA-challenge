Sub Analyze_Stocks()

    Dim MySheet As Worksheet
    Dim Stock As String
    Dim StartPrice As Double
    Dim LastPrice As Double
    Dim TotalVolume As LongLong
    Dim numRows As Long
    Dim gPerInc As Double
    Dim gPerDec As Double
    Dim gTotalVol As LongLong
    Dim gPerIncStock As String
    Dim gPerDecStock As String
    Dim gTotalVolStock As String
      
    ' Add Write locations
    Dim WriteRow As Long
    
    ' Add Read location
    Dim RdRow As Long
            
    For Each MySheet In Sheets
        ' Put the headers in the worksheets
        Worksheets(MySheet.Name).Activate
        MySheet.Range("I1").Value = "Ticker"
        MySheet.Range("J1").Value = "Yearly Change"
        MySheet.Range("K1").Value = "Percent Change"
        MySheet.Range("L1").Value = "Total Stock Volume"
    
        ' Initialize variables
        Stock = ""
        WrtRow = 2  ' first row to write summary info to
        
        numRows = Cells(Rows.Count, 1).End(xlUp).Row
        
        gPerInc = 0
        gPerDec = 0
        gTotalVol = 0
        gPerIncStock = ""
        gPerDecStock = ""
        gTotalVolStock = ""
        
        ' Cycle through each row in the worksheet to process the stocks
        For RdRow = 2 To numRows + 1
            ' Handle stock ticker change
            If Stock <> Cells(RdRow, 1).Value Then
                ' If this isn't the first stock, then write the contents of Stock to the worksheet then start moving data
                If Stock <> "" Then
                    MySheet.Cells(WrtRow, 9).Value = Stock
                    MySheet.Cells(WrtRow, 10).Value = LastPrice - StartPrice
                    
                    ' format color of yearly change cell
                    If MySheet.Cells(WrtRow, 10).Value > 0 Then
                        Cells(WrtRow, 10).Interior.ColorIndex = 4
                    ElseIf MySheet.Cells(WrtRow, 10).Value < 0 Then
                        Cells(WrtRow, 10).Interior.ColorIndex = 3
                    End If
                    
                    ' what to do about divide by zero problem
                    If StartPrice = 0 Then
                        MySheet.Cells(WrtRow, 11).Value = "N/A"
                    Else
                        MySheet.Cells(WrtRow, 11).Value = (LastPrice - StartPrice) / StartPrice
                        If MySheet.Cells(WrtRow, 11).Value > gPerInc Then
                            gPerInc = MySheet.Cells(WrtRow, 11).Value
                            gPerIncStock = Stock
                        ElseIf MySheet.Cells(WrtRow, 11).Value < gPerDec Then
                            gPerDec = MySheet.Cells(WrtRow, 11).Value
                            gPerDecStock = Stock
                        End If
                    End If
                    
                    MySheet.Cells(WrtRow, 12).Value = TotalVolume
                    If TotalVolume > gTotalVol Then
                        gTotalVol = TotalVolume
                        gTotalVolStock = Stock
                    End If
                    
                    WrtRow = WrtRow + 1
                End If
                
                ' Reinitialize variables for next written row
                Stock = Cells(RdRow, 1).Value
                StartPrice = Cells(RdRow, 3).Value
                LastPrice = 0
                TotalVolume = 0
            End If
            
            ' Process the row
            LastPrice = Cells(RdRow, 6).Value
            TotalVolume = TotalVolume + Cells(RdRow, 7).Value
                
        Next RdRow
        
        ' Format the change columns
        Columns("J").NumberFormat = "0.00"
        Columns("K").NumberFormat = "0.00%"
        
        ' Process greatest increase, decrease and volume on this sheet
        Dim incPrice As Double
        Dim decPrice As Double
        Dim grtVolume As LongLong
        
        ' column 15 starts
        MySheet.Cells(2, 15).Value = "Greatest % Increase"
        MySheet.Cells(3, 15).Value = "Greatest % Decrease"
        MySheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Summary headers
        MySheet.Cells(1, 16).Value = "Ticker"
        MySheet.Cells(1, 17).Value = "Value"
        
        'Add summary values and format
        MySheet.Cells(2, 16).Value = gPerIncStock
        MySheet.Cells(3, 16).Value = gPerDecStock
        MySheet.Cells(4, 16).Value = gTotalVolStock
        MySheet.Cells(2, 17).Value = gPerInc
        MySheet.Cells(2, 17).NumberFormat = "0.00%"
        MySheet.Cells(3, 17).Value = gPerDec
        MySheet.Cells(3, 17).NumberFormat = "0.00%"
        MySheet.Cells(4, 17).Value = gTotalVol
        
    Next MySheet
End Sub