VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CommandButton1_Click()
'The VBA of Wall Street

     Dim lastrow As Long  'Lastrow used to determine when loop returns to the next step
     Dim lastrow2 As Long
     Dim i As Long
     Dim Ticker As String
     Dim TrdDate As Double
     Dim OpenPx As Double
     Dim HighPx As Double  'Ignored
     Dim LowPx As Double   'Ignored
     Dim ClosePx As Double
     Dim TotalVolume As Double
     Dim YearlyChg As Double
     Dim PctChg As Double
     
     Dim ResultRow As Integer
     
     Dim myrange As Range   'myrange used to place summary total volume and max & min percentage
     Dim myrange2 As Range
     Dim myrange3 As Range

     Dim ws As Worksheet    'tool to iterate calculations across multiple worksheets
     
    For j = 2 To ThisWorkbook.Worksheets.Count   'Iterate through the worksheets
        
        'get the name of each sheet
        NameOfActiveSheet = ThisWorkbook.Worksheets(j).Name
        'set the ws variable to the name
        Set ws = ThisWorkbook.Worksheets(NameOfActiveSheet)
        'update the progress cell on the main sheet
        Cells(10, 3) = NameOfActiveSheet
         
         
        'Creating annual table to provide price change,
        'percentage change, and total trade volume.
        'Working off the ticker provided, the code will
        'run through tickers to calculate the above-mentioned
        'list.
        
        'this gets the bottom of the data
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
        ResultRow = 2
        
        TotalVolume = 0
        
        Ticker = ""
        
        For i = 2 To lastrow
           
           'read in the values that we need later so the code does not waste lower empty cells
           Ticker = ws.Cells(i, 1).Value
           Nextticker = ws.Cells(i + 1, 1).Value
           
           MorningPrice = ws.Cells(i, 3).Value 'first row with ticker chgs
           EveningPrice = ws.Cells(i, 6).Value
           VolumeDaily = ws.Cells(i, 7).Value
           TradeDate = ws.Cells(i, 2).Value
           
           TotalVolume = TotalVolume + VolumeDaily 'cumulative volume by ticker
           
           'It would be nice to get rid of this statement somehow
           If i = 2 Then StartPrice = MorningPrice
           
           If Ticker <> Nextticker Then  'we are on the same ticker
               YearlyChg = EveningPrice - StartPrice
               If StartPrice <> 0 Then
                   PctChg = YearlyChg / StartPrice   'isolates the Jan-1 opening price in order
                                                     'to properly calculate the yearly chang
               Else
                   PctChg = 0
                   ws.Cells(ResultRow, 13) = "Start Price invalid."
                   ws.Cells(ResultRow, 13).Interior.ColorIndex = 6
               End If
               ws.Cells(ResultRow, 9) = Ticker          'Tool to overlook stock that has a zero open
                                                        'price which leads to a division by zero when
                                                        'calculating thestock price percentage price

               If YearlyChg < 0 Then
                   ws.Cells(ResultRow, 10).Interior.ColorIndex = 3
               ElseIf YearlyChg > 0 Then
                   ws.Cells(ResultRow, 10).Interior.ColorIndex = 4  'tool to attach color based
               Else                                                 'on percent change
                   ws.Cells(ResultRow, 10).Interior.ColorIndex = 2
               End If
               
               ws.Cells(ResultRow, 10) = YearlyChg                  'summary performance table
               ws.Cells(ResultRow, 11) = PctChg
               ws.Cells(ResultRow, 12) = TotalVolume
               
               'new start price
               StartPrice = ws.Cells(i + 1, 3)
               TotalVolume = 0
               ResultRow = ResultRow + 1
           End If
           
        Next i
    
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

        Set myrange = ws.Range("k2:k" & Trim(Str(lastrow2))) 'this is the percent change column
        Set myrange2 = ws.Range("i2:i" & Trim(Str(lastrow2))) 'this is the ticker column
        Set myrange3 = ws.Range("l2:l" & Trim(Str(lastrow2))) 'this is the volume column
        
        'get the maximum percent change                       'Use of Excel max and min functions and location
        maxpct = Application.WorksheetFunction.Max(myrange)
        ws.Cells(2, 16) = maxpct
        
        'get the ticker for the maximum percent change
        MaxTicker = Application.WorksheetFunction.XLookup(maxpct, myrange, myrange2)
        ws.Cells(2, 15) = MaxTicker
        
        'get the greatest decrease in percent change
        minpct = Application.WorksheetFunction.Min(myrange)
        ws.Cells(3, 16) = minpct
        
        'get the ticker for the greatest decrease in percent change
        MinTicker = Application.WorksheetFunction.XLookup(minpct, myrange, myrange2)
        ws.Cells(3, 15) = MinTicker
        
        'get the max volume
        maxvol = Application.WorksheetFunction.Max(myrange3)
        ws.Cells(4, 16) = maxvol
        
        'get the ticker for the max volume
        MaxTicker = Application.WorksheetFunction.XLookup(maxvol, myrange3, myrange2)
        ws.Cells(4, 15) = MaxTicker

    Next j  'iterate through the workseets


End Sub
