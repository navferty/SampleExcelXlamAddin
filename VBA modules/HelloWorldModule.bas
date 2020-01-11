Attribute VB_Name = "HelloWorldModule"
Option Explicit

Public Sub SayHelloWorld()
    Dim firstCellValue As String
    
    firstCellValue = ThisWorkbook.Worksheets(1).Cells(1, 1).Value
    MsgBox "Hello, world! Cell A1 value is " & firstCellValue
End Sub
