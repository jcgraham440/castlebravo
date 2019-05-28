
Type Greatest_YearlyChange_Percentage
    GYP As Double
    Ticker As String
End Type

Type Lowest_YearlyChange_Percentage
    LYP As Double
    Ticker As String
End Type

Type Greatest_Yearly_Volume
    GYV As LongLong
    Ticker As String
End Type

Sub Summarize_All()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets()
        With ws
            
            Dim CurrentTicker As String
            Dim NextTicker As String
            Dim Current_Year_OpeningPrice As Double
            Dim Current_Year_CurrentPrice As Double
            Dim Current_Year_EndingPrice As Double
            Dim YearlyChange As Double
            Dim YearlyChangePercentage As Double
            Dim Current_Year_TotalVolume As LongLong



            Dim GP As Greatest_YearlyChange_Percentage
            Dim LP As Lowest_YearlyChange_Percentage
            Dim GV As Greatest_Yearly_Volume

            Dim i As LongLong
            Dim TickerSummaryIndex As Integer
            Dim NewYear As Boolean
            '
            ' Set display names for summary columns
            '
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "YearlyChange"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"

            ' Prime the pump
            i = 2
            TickerSummaryIndex = 2
            NewYear = True
            GP.GYP = 0
            LP.LYP = 0
            GV.GYV = 0

            '
            ' Loop through all rows, as long as there's data
            '
            While (.Cells(i, 1) <> "")

                Current_Year_TotalVolume = Current_Year_TotalVolume + .Cells(i, 7)
                CurrentTicker = .Cells(i, 1)
                NextTicker = .Cells(i + 1, 1)
                If NewYear Then
                    Current_Year_OpeningPrice = .Cells(i, 3)
                    NewYear = False
                End If
    
                ' Detect End of Year
                If (NextTicker <> CurrentTicker) Then
                    ' Do year end summary for this ticker symbol:
                    '   TotalVolume
                    '   YearlyChange from Opening Price
                    '   PercentChange from Opening Price
                    .Cells(TickerSummaryIndex, 9).Value = CurrentTicker
                    Current_Year_Ending_Price = .Cells(i, 6)
                    YearlyChange = Current_Year_Ending_Price - Current_Year_OpeningPrice
                    .Cells(TickerSummaryIndex, 10).Value = YearlyChange
        
                    ' Set Cell Color
                    If YearlyChange > 0 Then
                        .Cells(TickerSummaryIndex, 10).Interior.Color = RGB(0, 255, 0)
                    Else
                        .Cells(TickerSummaryIndex, 10).Interior.Color = RGB(255, 0, 0)
                    End If
        
                    If (Current_Year_OpeningPrice <> 0) Then
                        YearlyChangePercentage = YearlyChange / Current_Year_OpeningPrice
                    Else
                        YearlyChangePercentage = 0
                    End If
                    .Cells(TickerSummaryIndex, 11).Value = YearlyChangePercentage
                    .Cells(TickerSummaryIndex, 12).Value = Current_Year_TotalVolume
        
                    '
                    ' Get/Set superlatives
                    '
                    If (Current_Year_TotalVolume > GV.GYV) Then
                        GV.GYV = Current_Year_TotalVolume
                        GV.Ticker = CurrentTicker
                    End If
                    If (YearlyChangePercentage >= 0) Then
                        If (YearlyChangePercentage > GP.GYP) Then
                            GP.GYP = YearlyChangePercentage
                            GP.Ticker = CurrentTicker
                        End If
                    Else
                        If (YearlyChangePercentage < LP.LYP) Then
                            LP.LYP = YearlyChangePercentage
                            LP.Ticker = CurrentTicker
                        End If
                    End If
        
                    ' Increment Ticker Summary Index
                    TickerSummaryIndex = TickerSummaryIndex + 1
        
                    ' Reset Current_Year_TotalVolume to zero
                    Current_Year_TotalVolume = 0
        
                    ' Declare New Year
                    NewYear = True
                End If
                i = i + 1
            Wend

            ' All done. Set superlative cell values for this sheet
            .Range("P2").Value = GP.Ticker
            .Range("Q2").Value = GP.GYP
            .Range("P3").Value = LP.Ticker
            .Range("Q3").Value = LP.LYP
            .Range("P4").Value = GV.Ticker
            .Range("Q4").Value = GV.GYV

        End With
        ' go to next worksheet
        Next ws
End Sub
