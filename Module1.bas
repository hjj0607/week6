Attribute VB_Name = "Module1"
Option Explicit

Sub ¬d¸ß¤f¸n()
Dim qMan As String
Dim rownum As Integer
Dim content  As String
Dim paystatus As Boolean
qMan = Range("G1").Value
For rownum = 2 To 7
If (Cells(rownum, "A").Value = qMan) Then
Range("G2").Value = Cells(rownum, "B").Value
If (Cells(rownum, 3).Value = 0) Then
paystatus = False
Else
paystatus = True
End If
MsgBox qMan & "¥I´Úª¬ºA" & paystatus
Else
End If
Next

End Sub
