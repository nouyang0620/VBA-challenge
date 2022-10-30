Attribute VB_Name = "Module1"
' Part 1: Loop through all stocks for one year
' Part 2: Output ticker symbol, yearly change from opening price at beginning of given year to closing price at the end of that year,
    ' Percentage change from opening price at beginning of year to closing at end of that year, total stock volume of the stock.
' Part 3: Use conditional formatting to highlight pos change in green and neg change in red
' Part 4: BONUS - Return greatest % increase, greatest % decrease, & greatest total volume





 
Sub MultiYearStockData():

    ' Declare all variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim StartRow As Integer
    Dim LastRow As String
    'Dim PrevRow As Integer
    
    Dim Increase As Double
    Dim Decrease As Double
    Dim GreatVolume As Double
    
    ' Loop through each worksheet
    For Each ws In Worksheets
    
        ' Get the WorksheetName
        WorksheetName = ws.Name
        
        ' Create new column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Assign values
        StartRow = 2
        PrevRow = 1
        
        
        ' Find column A's last row
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Loop through all ticker symbols
        For i = 2 To LastRow
        
            ' When symbols change
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                ' Print ticker symbol
                ws.Range("I" & StartRow).Value = Ticker
                
                ' Add to PrevRow
                PrevRow = PrevRow + 1
                
                ' Calculate yearly change
                OpenYear = ws.Cells(PrevRow, 3).Value
                CloseYear = ws.Cells(i, 6).Value
                YearlyChange = CloseYear - OpenYear
                
                ' Print yearly change
                ws.Range("J" & StartRow).Value = YearlyChange
                
                ' Highlight yearly change cell green or red based on value
                If ws.Range("J" & StartRow).Value >= 0 Then
                
                    ws.Range("J" & StartRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & StartRow).Interior.ColorIndex = 3
                End If
                    
                
                ' Calculate percent change
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(StartRow, 3).Value) / ws.Cells(StartRow, 3).Value)
                
                ' Print percent changee and format to a percentage
                ws.Range("K" & StartRow).Value = PercentChange
                ws.Range("K" & StartRow).Value = Format(PercentChange, "Percent")
                
                ' Calculate total volume
                For j = PrevRow To i
                    
                    TotalVolume = TotalVolume + ws.Cells(j, 7).Value
                    
                Next j
                
                If OpenYear = 0 Then
                
                    PercentChange = CloseYear
                
                Else
                
                    PercentChange = YearlyChange / OpenYear
                End If
                
                ' Print total volume
                ws.Range("L" & StartRow).Value = TotalVolume
                
                
                ' Add one to the start row
                StartRow = StartRow + 1
                
                ' Reset
                YearlyChange = 0
                PercentChange = 0
                TotalVolume = 0
                
                PrevRow = i
                
                
            End If
                
        Next i
    
    Next ws
        
       
End Sub
            
            
            

