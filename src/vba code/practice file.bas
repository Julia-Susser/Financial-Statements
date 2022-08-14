Attribute VB_Name = "Module1"
Sub Macro1()
'
' Macro3 Macro
'
' Keyboard Shortcut: Ctrl+l
'
    Range("A175").Select
    ActiveCell.FormulaR1C1 = "hi"
    Range("A176").Select
End Sub

Sub Macro2()
'
' Macro2 Macro
'

'
    Columns("E:E").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("G16").Select
End Sub



Function MessageBox_Demo()
 'Message Box with just prompt message
 MsgBox ("Welcome")

 'Message Box with title, yes no and cancel Butttons
 A = MsgBox("Do you like blue color?", 3, "Choose options")
 ' Assume that you press No Button
 MsgBox ("The Value of a is " & A)
 
End Function

Private Sub Variables_demo_Click()
    Dim password As String
    password = "Admin#1"
    Dim num As Integer
    num = 1234
    Dim BirthDay As Date
    BirthDay = 30 / 10 / 2020
    MsgBox ("Passowrd is " & password & Chr(10) & "Value of num is " & num & Chr(10) & "Value of Birthday is " & BirthDay)
    
End Sub


Sub Macro()
    
    Worksheets(1).Activate
End Sub


Sub Test()
    Dim value1 As String
    Dim value2 As String
    value1 = ThisWorkbook.Sheets("Data").Range("A1").Value 'value from sheet1
    value2 = ThisWorkbook.Sheets(2).Range("A1").Value 'value from sheet2
    If value1 = value2 Then ThisWorkbook.Sheets(2).Range("L1").Value = value1 'or 2
    MsgBox ("hi")
    Dim MyVar
    MyVar = "Come see me in the Immediate pane."
    Debug.Print MyVar
End Sub


Function getColumns()
    Dim dataArea As Excel.Range
   Set dataArea = ThisWorkbook.Worksheets(1).Range("A:A")
   
   Dim valuesArray() As Variant
   valuesArray = dataArea.Value
   
   
End Function


Function PrintRange()
    Dim dataArea As Excel.Range
    Dim delim As String
    delim = ","
    Set dataArea = ThisWorkbook.Worksheets(1).Range("A1:B3")
    Dim myRow As Range, v As Variant, i As Long
    For Each myRow In dataArea.Rows
        ReDim v(1 To myRow.Cells.Count)
        For i = 1 To myRow.Cells.Count
            v(i) = myRow.Cells(1, i).Value
        Next i
        Debug.Print Join(v, delim)
    Next myRow
    
End Function




Function Hi()
    Hi = 10
End Function
    

