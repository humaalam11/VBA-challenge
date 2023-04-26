Attribute VB_Name = "Module2"

Sub Stock():

Dim ws As Worksheet

For Each ws In Worksheets 'Worsheet Loop

 ws.Activate              'Applies to all worsheets

  Dim StockName As String
  Dim StockTotal As Double
  StockTotal = 0


  Dim SummaryTable As Double
  SummaryTable = 2

  Dim OpenPrice, ClosePrice As Double
  OpenPrice = Cells(2, 3).Value

  Dim PercentDecreaseTicker As String
  Dim PercentIncreaseTicker As String
  Dim StockMax As Long

'Creating Column Labels
  Cells(1, "J").Value = "Stock Name"
  Cells(1, "K").Value = "Yearly Change"
  Cells(1, "L").Value = "Percentage Change"
  Cells(1, "M").Value = "Total Stock Volume"
  Cells(1, "Q").Value = "Ticker"
  Cells(1, "R").Value = "Value"
  Cells(2, "P").Value = "Greatest % Increase"
  Cells(3, "P").Value = "Greatest % Decrease"
  Cells(4, "P").Value = "Greatest Total Volume"

'Initiating ForLoop to Calculate Totals
  For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

    StockTotal = StockTotal + Cells(i, 7).Value
    PercentIncreaseNumber = Cells(2, "L").Value
    PercentDecreaseNumber = Cells(2, "L").Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        StockName = Cells(i, 1).Value
        ClosePrice = Cells(i, 6).Value

        Range("J" & SummaryTable).Value = StockName
        Range("M" & SummaryTable).Value = StockTotal
        Range("K" & SummaryTable).Value = ClosePrice - OpenPrice
        
    'Calculating Yearly Change, column "K", Highlight cells with Positive Values with Green and Negative with Red
        If Range("K" & SummaryTable).Value > 0 Then
        
            Range("K" & SummaryTable).Interior.ColorIndex = 4
        Else
            Range("K" & SummaryTable).Interior.ColorIndex = 3

        End If
        
    'Calculating Percentage Change in column "L"
        If OpenPrice <> 0 Then
        
            Range("L" & SummaryTable).Value = FormatPercent((ClosePrice - OpenPrice) / OpenPrice, 2)
            
            Else
            
            Range("L" & SummaryTable).Value = Null
            
        End If
        
    'Calculating Greatest % Increase
        If Range("L" & SummaryTable).Value > PercentIncreaseNumber Then
        
            PercentIncreaseNumber = Range("L" & SummaryTable).Value
            
            PercentIncreaseTicker = Range("J" & SummaryTable).Value
            
        End If
        
    'Calculating Greatest % Decrease
        If Range("L" & SummaryTable).Value < PercentDecreaseNumber Then
        
            PercentDecreaseNumber = Range("L" & SummaryTable).Value
            
            PercentDecreaseTicker = Range("J" & SummaryTable).Value
            
        End If
        
    'Reset Values
        SummaryTable = SummaryTable + 1
        StockTotal = 0
        
        OpenPrice = Cells(i + 1, 3).Value
    
    
    End If


    Next i
  'Recording Totals for Ticker and Value in the 3rd Table
    Range("R2") = "%" & WorksheetFunction.Max(Range("L2:L" & SummaryTable)) * 100
    Range("R3") = "%" & WorksheetFunction.Min(Range("L2:L" & SummaryTable)) * 100
    Range("R4") = WorksheetFunction.Max(Range("M2:M" & SummaryTable))
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & SummaryTable)), Range("L2:L" & SummaryTable), 0)
    Range("Q2") = Cells(increase_number + 1, 10)
    
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & SummaryTable)), Range("L2:L" & SummaryTable), 0)
    Range("Q3") = Cells(decrease_number + 1, 10)
    
    StockMax = WorksheetFunction.Match(WorksheetFunction.Max(Range("M2:M" & SummaryTable)), Range("M2:M" & SummaryTable), 0)
    Range("Q4") = Cells(StockMax + 1, 10)
    
Next ws
   
End Sub

