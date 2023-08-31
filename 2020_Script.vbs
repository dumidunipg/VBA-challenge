VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub table1_names()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

End Sub
Sub stocks_begin_end()
Dim year_start As Long
year_start = 20200102
MsgBox (year_start)
Dim year_start_value As Double
year_start_value = 0
MsgBox (year_start_value)
Dim year_start_row As Integer
year_start_row = 2

Dim year_end As Long
year_end = 20201231
MsgBox (year_end)
Dim year_end_value As Double
year_end_value = 0
MsgBox (year_end_value)
Dim year_end_row As Integer
year_end_row = 2

Dim stock_name As String
Dim stock_name_row As Integer
stock_name_row = 2



For i = 2 To 759001
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        stock_name = Cells(i, 1).Value
        Range("I" & stock_name_row).Value = stock_name
        stock_name_row = stock_name_row + 1
    End If
    For j = 1 To 7
        If Cells(i, j).Value = year_start Then
            year_start_value = year_start_value + Cells(i, 3)
            Range("N" & year_start_row).Value = year_start_value
            year_start_row = year_start_row + 1
            year_start_value = 0
        End If
        If Cells(i, j).Value = year_end Then
            year_end_value = year_end_value + Cells(i, 6)
            Range("M" & year_end_row).Value = year_end_value
            year_end_row = year_end_row + 1
            year_end_value = 0
        End If
    Next j
Next i

End Sub


Sub yrly_change()

Dim yearly_change As Double
yearly_change = 0
Dim yearly_change_row As Integer
yearly_change_row = 2


For i = 2 To 3001
    If Cells(i + 1, 13).Value <> Cells(i, 14).Value Then
        yearly_change = Cells(i, 13).Value - Cells(i, 14).Value
        Range("J" & yearly_change_row).Value = yearly_change
        yearly_change_row = yearly_change_row + 1
        yearly_change = 0
    End If
Next i


End Sub
Sub color_yrly_change()


For i = 2 To 3001
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.Color = vbGreen
    ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.Color = vbRed
    End If
Next i

End Sub

Sub percent_change()

Dim percent_change As Double
percent_change = 0
Dim percent_change_row As Integer
percent_change_row = 2


For i = 2 To 3001
    If Cells(i + 1, 13).Value <> Cells(i, 14).Value Then
        percent_change = (Cells(i, 13).Value - Cells(i, 14).Value) / (Cells(i, 14).Value)
        Range("K" & percent_change_row).Value = percent_change
        percent_change_row = percent_change_row + 1
        percent_change = 0
    End If
Next i

Range("K2:K3001").NumberFormat = "0.00%"

End Sub

Sub total_stock()

Dim stock_volume As LongLong
stock_volume = 0
Dim stock_volume_row As Integer
stock_volume_row = 2

For i = 2 To 759001
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        stock_volume = stock_volume + Cells(i, 7).Value
        Range("L" & stock_volume_row).Value = stock_volume
        stock_volume_row = stock_volume_row + 1
        stock_volume = 0
    Else
        stock_volume = stock_volume + Cells(i, 7).Value
    End If
Next i

Range("M2:N3001").Clear

End Sub
Sub table2_names()

Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

End Sub


Sub greatest_percent_increase()

Dim maxval As Double

Range("K2:K3001").NumberFormat = "0.0000000"

maxval = WorksheetFunction.Max(Range("K2:K3001"))
Cells(2, 16).Value = maxval


For i = 2 To 3001
    If Cells(i, 11).Value = maxval Then
    MsgBox (maxval)
            Cells(2, 15).Value = Cells(i, 9).Value
    End If
Next i
    
Cells(2, 16).NumberFormat = "0.00%"
Range("K2:K3001").NumberFormat = "0.00%"
   
End Sub

Sub greatest_percent_decrease()

Dim minval As Double

Range("K2:K3001").NumberFormat = "0.0000000"

minval = WorksheetFunction.Min(Range("K2:K3001"))
Cells(3, 16).Value = minval

For i = 2 To 3001
    If Cells(i, 11).Value = minval Then
    MsgBox (minval)
        Cells(3, 15).Value = Cells(i, 9).Value
    End If
Next i

Cells(3, 16).NumberFormat = "0.00%"
Range("K2:K3001").NumberFormat = "0.00%"

End Sub


Sub greatest_total_volume()

Dim max_vol As Double

max_vol = WorksheetFunction.Max(Range("L2:L3001"))
Cells(4, 16).Value = max_vol

For i = 2 To 3001
    If Cells(i, 12).Value = max_vol Then
    MsgBox (max_vol)
        Cells(4, 15).Value = Cells(i, 9).Value
    End If
Next i


End Sub


Sub Run_everyMacro()

Call table1_names
Call stocks_begin_end
Call yrly_change
Call color_yrly_change
Call percent_change
Call total_stock
Call table2_names
Call greatest_percent_increase
Call greatest_percent_decrease
Call greatest_total_volume

End Sub


