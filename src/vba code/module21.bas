Attribute VB_Name = "RearrangeCols"

Sub Separate()
Attribute Separate.VB_ProcData.VB_Invoke_Func = "S\n14"
    Call RearrangeCols("separated", 3)
End Sub


Sub Order()
Attribute Order.VB_ProcData.VB_Invoke_Func = "O\n14"
    Call Delete_Columns
    Call RearrangeCols("order", 3)
End Sub

Sub RearrangeCols(Optional how As String = "order", Optional counter = 1, Optional SheetName As String)
    Dim sheet As Worksheet
    If SheetName = "" Then
        SheetName = ActiveSheet.Name
    End If
    Set sheet = Worksheets(SheetName)
    Dim Headers As Variant, Header As Variant, Found As Range, FoundIndx As Integer
    If how = "order" Then
        Headers = getPossibleHeaders(sheet)
    Else
        Headers = getPossibleHeadersSeparated(sheet)
    End If
    
    For i = LBound(Headers) To UBound(Headers)
        Header = Headers(i)
        If Header = "space" Then
            Columns(counter).insert Shift:=xlToRight
            Columns(counter).insert Shift:=xlToRight
            counter = counter + 2
        End If
        Set Found = Rows(1).Find(Header, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not Found Is Nothing Then
            FoundIndx = Found.Column
            If FoundIndx <> counter Then
                Found.EntireColumn.Cut
                Columns(counter).insert Shift:=xlToRight
            End If
            counter = counter + 1
        End If
    Next i
End Sub


Sub Delete_Columns()
Dim C As Integer
C = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Do Until C = 0
    If WorksheetFunction.CountA(Columns(C)) = 0 Then
        Columns(C).Delete
    End If
    C = C - 1
Loop
End Sub

Function getHeaders(sheet As Worksheet)
    Dim LastCol As String, rng As Variant, cell As Range, Headers As Variant
    LastCol = LastColumn(sheet)
    rng = "A1:" & LastCol
    Set rng = ActiveSheet.Range(rng)
    For Each cell In rng
        Headers = addValueToArray(Headers, cell.Value)
    Next cell
    getHeaders = Headers
End Function

Function LastColumn(sheet)
    ' LastColumn = ActiveSheet.Range("A1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    LastColumn = sheet.Cells(1, Columns.Count).End(xlToLeft).Address
End Function

Function Contains(strIn As Variant, arrList As Variant) As Boolean
Contains = Not (IsError(Application.Match(strIn, arrList, 0)))
End Function

Function getArrayLength(arrList As Variant) As Integer
    If IsEmpty(arrList) Then
      getArrayLength = 0
   Else
      getArrayLength = UBound(arrList) - LBound(arrList) + 1
   End If
End Function

Function addValueToArray(arrList As Variant, v As Variant) As Variant
    Dim l As Integer
    l = getArrayLength(arrList) + 1
    If l = 1 Then
        ReDim arrList(1 To 1)
    End If
    ReDim Preserve arrList(1 To l)

    arrList(l) = v
    addValueToArray = arrList
End Function


Function getYears(Headers As Variant)
    Dim Header As Variant, Years As Variant
    For Each Header In Headers
        Header = Right(Header, 4)
        If IsNumeric(Header) Then
            If Not Contains(Header, Years) Then
                Years = addValueToArray(Years, Header)
            End If
        End If
        
        
    Next Header
    getYears = Years
End Function


Function reOrderYears(Years As Variant)
    current = LBound(Years)
    For i = LBound(Years) To UBound(Years)
        For C = current To UBound(Years)
        If Years(C) < Years(i) Then
            temp = Years(i)
            Years(i) = Years(C)
            Years(C) = temp
        End If
        Next C
        current = current + 1
    Next i
    reOrderYears = Years
End Function




Function getAllPossibleHeaders(FiscalYear As Variant, sheet As Worksheet) As Variant
    Dim Headers As Variant, h As Variant, Period As Variant, Year As Variant
    Headers = getHeaders(sheet)
    Years = getYears(Headers)
    Years = reOrderYears(Years)
    For Each Year In Years
    For Each Period In FiscalYear
        Period = Period & " " & Year
        h = addValueToArray(h, Period)
    Next Period
    Next Year
    getAllPossibleHeaders = h
End Function




Function getPossibleHeadersSeparated(sheet As Worksheet) As Variant
    Dim Headers1 As Variant, Headers2 As Variant
    FiscalYear = Array("Q1", "Q2", "Q3", "Q4", "Year Ended")
    Headers1 = getAllPossibleHeaders(Array("Year Ended", "E Year Ended"), sheet)
    Headers1 = addValueToArray(Headers1, "space")
    Headers2 = getAllPossibleHeaders(Array("Q1", "Q2", "Q3", "Q4", "E Q1", "E Q2", "E Q3", "E Q4"), sheet)
    For Each h In Headers2
        Headers1 = addValueToArray(Headers1, h)
    Next h
    getPossibleHeadersSeparated = Headers1
End Function



Function getPossibleHeaders(sheet As Worksheet) As Variant
    Dim Headers1 As Variant, FiscalYear As Variant
    FiscalYear = Array("Q1", "Q2", "Q3", "Q4", "Year Ended")
    Headers1 = getAllPossibleHeaders(FiscalYear, sheet)
    getPossibleHeaders = Headers1
End Function

Sub insertEmpty()
Attribute insertEmpty.VB_ProcData.VB_Invoke_Func = " \n14"
'
' insert Macro
'

'
    Selection.insert Shift:=xlToRight
    Selection.insert Shift:=xlToRight
End Sub
