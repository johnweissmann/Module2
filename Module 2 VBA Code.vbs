{\rtf1\ansi\ansicpg1252\cocoartf1671\cocoasubrtf600
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Module2()\
For Each Sheet In Worksheets\
Worksheets(Sheet.Name).Cells.EntireColumn.AutoFit\
Worksheets(Sheet.Name).Cells.EntireRow.AutoFit\
Sheet.Cells(1, 9).Value = "Ticker"\
Sheet.Cells(1, 10).Value = "Yearly Change"\
Sheet.Cells(1, 11).Value = "Percent Change"\
Sheet.Range("K:K").NumberFormat = "0.00%"\
Sheet.Cells(1, 12).Value = "Total Stock Volume"\
Sheet.Cells(2, 15).Value = "Greatest % Increase"\
Sheet.Cells(2, 17).NumberFormat = "0.00%"\
Sheet.Cells(3, 15).Value = "Greatest % Decrease"\
Sheet.Cells(3, 17).NumberFormat = "0.00%"\
Sheet.Cells(4, 15).Value = "Greatest Total Volume"\
Sheet.Cells(1, 16).Value = "Ticker"\
Sheet.Cells(1, 17).Value = "Value"\
Dim Ticker As String\
Dim TotalVolume As Double\
TotalVolume = 0\
Dim TickerRow As Integer\
TickerRow = 2\
Dim OpenValue As Double\
OpenValue = Sheet.Cells(2, 3).Value\
lastrow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row\
For i = 2 To lastrow\
    If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1).Value Then\
    Ticker = Sheet.Cells(i, 1).Value\
    TotalVolume = TotalVolume + Sheet.Cells(i, 7)\
    CloseValue = Sheet.Cells(i, 6).Value\
    YearlyChange = CloseValue - OpenValue\
    PercentChange = ((CloseValue - OpenValue) / OpenValue)\
    Sheet.Cells(TickerRow, 9).Value = Ticker\
    Sheet.Cells(TickerRow, 10).Value = YearlyChange\
        If Sheet.Cells(TickerRow, 10).Value >= 0 Then\
        Sheet.Cells(TickerRow, 10).Interior.ColorIndex = 4\
        ElseIf Sheet.Cells(TickerRow, 10).Value < 0 Then\
        Sheet.Cells(TickerRow, 10).Interior.ColorIndex = 3\
        End If\
    Sheet.Cells(TickerRow, 11).Value = PercentChange\
    Sheet.Cells(TickerRow, 12).Value = TotalVolume\
        If PercentChange = WorksheetFunction.Max(Sheet.Range("K:K")) Then\
        Sheet.Cells(2, 16).Value = Ticker\
        Sheet.Cells(2, 17).Value = PercentChange\
        End If\
         If PercentChange = WorksheetFunction.Min(Sheet.Range("K:K")) Then\
        Sheet.Cells(3, 16).Value = Ticker\
        Sheet.Cells(3, 17).Value = PercentChange\
        End If\
        If TotalVolume = WorksheetFunction.Max(Sheet.Range("L:L")) Then\
        Sheet.Cells(4, 16).Value = Ticker\
        Sheet.Cells(4, 17).Value = TotalVolume\
        End If\
    TickerRow = TickerRow + 1\
    TotalVolume = 0\
    OpenValue = Sheet.Cells(i + 1, 3).Value\
    Else\
    TotalVolume = TotalVolume + Sheet.Cells(i, 7)\
    End If\
Next i\
Next Sheet\
End Sub}