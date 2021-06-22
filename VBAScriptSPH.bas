Attribute VB_Name = "Module1"
Sub PopulateSummary():

'Added header for summary
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
'Added in conditional formatting
    Dim ChangeRange As Range
    Set ChangeRange = Range("J:J")
    ChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ChangeRange.FormatConditions(1).Interior.Color = vbRed
    ChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ChangeRange.FormatConditions(2).Interior.Color = vbGreen
    Range("J1:J1").FormatConditions.Delete

'Added in % formatting
    Range("K:K").NumberFormat = "0.00%"
    
'Added in Bonus summary table headers
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0"

    

End Sub



Sub PullData():

'Finds number of rows in sheet
    Dim RowCount As Long
    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    

'Loops through Column A and finds when there are unique values and puts them into Arrays
    Dim i As Long
    Dim ArrayLength As Long
    Dim TickerArray() As String
    Dim OpeningArray() As Double
    Dim ClosingArray()
    
    For i = 2 To RowCount + 1
        
    'This finds when there is a unique tickler and adds to TickerArray() as well as opening price and adds to OpeningArray()
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            ReDim Preserve TickerArray(ArrayLength)
            TickerArray(ArrayLength) = Cells(i, 1).Value
            
            ReDim Preserve OpeningArray(ArrayLength)
            OpeningArray(ArrayLength) = Cells(i, 3).Value
            
            ReDim Preserve ClosingArray(ArrayLength)
            ClosingArray(ArrayLength) = Cells(i - 1, 6).Value
            
            ArrayLength = ArrayLength + 1
        End If
        
    Next i
    
'Places the information into summary tab
    Dim i2 As Long
    
    For i2 = 0 To ArrayLength - 2
        Cells(i2 + 2, 9).Value = TickerArray(i2)
        Cells(i2 + 2, 10).Value = ClosingArray(i2 + 1) - OpeningArray(i2)
        
        'Data had 0s in it and had to add for statement to fix divide by zero error.
        If OpeningArray(i2) = 0 Then
            Cells(i2 + 2, 11).Value = 0
        Else
            Cells(i2 + 2, 11).Value = (ClosingArray(i2 + 1) / OpeningArray(i2)) - 1
        End If
        Cells(i2 + 2, 12).Value = WorksheetFunction.SumIf(Range("A:A"), Cells(i2 + 2, 9), Range("G:G"))
    Next i2

End Sub

Sub BonusDate():
    'This will do a lookup for the min max for volume and % increase and then a lookup for the ticker
    
    Dim RowSumCount As Long
    RowSumCount = Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim GreatInc As Double
    Dim GreatIncT As String
    Dim GreatDec As Double
    Dim GreatDecT As String
    Dim GreatVolume As Double
    Dim GreatVolumeT As String
    Dim i4 As Long
    
    
    
   
    For i4 = 2 To RowSumCount
    'Finds Greatest Increase
    If Cells(i4, 11).Value >= GreatInc Then
        GreatInc = Cells(i4, 11).Value
        GreatIncT = Cells(i4, 9).Value
        Range("Q2").Value = GreatInc
        Range("P2").Value = GreatIncT
        
    End If
    
    'Finds Greatest Decrease
    If Cells(i4, 11).Value <= GreatDec Then
        GreatDec = Cells(i4, 11).Value
        GreatDecT = Cells(i4, 9).Value
        Range("Q3").Value = GreatDec
        Range("P3").Value = GreatDecT
        
    End If
    
    If Cells(i4, 12).Value >= GreatVolume Then
        GreatVolume = Cells(i4, 12).Value
        GreatVolumeT = Cells(i4, 9).Value
        Range("Q4").Value = GreatVolume
        Range("P4").Value = GreatVolumeT
        
    End If
    
    
    Next i4
   
    
    
  
End Sub

Sub AllWorksheets():
   'This will execute the PullData() and Populate Headers through each worksheet
    
    Dim NumSheets As Long
    NumSheets = Application.Sheets.Count
    Dim i3 As Long
    
    
    For i3 = 1 To NumSheets
        Worksheets(i3).Activate
        PopulateSummary
        PullData
        BonusDate
    Next i3
    
End Sub









