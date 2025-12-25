Attribute VB_Name = "Module2"
Option Explicit
'====================================================
' Project     : EMI Calculator
' Module      : ClearContents
' Description : Clears EMI output range and borders
' Author      : Blaaji
' Created     : 2024-11-10
' Version     : 1.0.0
' Last Update : 2025-12-25
'====================================================
Sub Clearcontents()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableCancelKey = xlDisabled

Dim lr As Long, Rng As Range, nr As Long, fr As Long
fr =  11
nr = fr + 1
lr = Sheet1.Range("B" & Rows.Count).End(xlUp).Row
If lr > fr Then
    Set Rng = Sheet1.Range("B" & nr, "G" & lr)
        With Rng
        .Clearcontents
        .Borders.LineStyle = xlNone
        End With
End If

End Sub
