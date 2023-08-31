Attribute VB_Name = "Module1"

Sub worksheet_loop()

Dim current_sheet As Worksheet

For Each current_sheet In Worksheets
    If current_sheet.Name = "2018" Then
        Sheet1.Activate
        MsgBox current_sheet.Name
        Call Sheet1.Run_everyMacro
    End If
    If current_sheet.Name = "2019" Then
        Sheet2.Activate
        MsgBox current_sheet.Name
        Call Sheet2.Run_everyMacro
    End If
    If current_sheet.Name = "2020" Then
        Sheet3.Activate
        MsgBox current_sheet.Name
        Call Sheet3.Run_everyMacro
    End If
Next
End Sub

