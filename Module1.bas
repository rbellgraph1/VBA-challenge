Attribute VB_Name = "Module1"
Option Explicit

Sub clearcontents():

Dim ws As Worksheet
For Each ws In Worksheets

     ws.Range("J:m").clearcontents
     ws.Range("P:R").clearcontents
     ws.Range("K:K").ClearFormats
Next ws
  

End Sub


