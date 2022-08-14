Attribute VB_Name = "Module2"

Sub RunArrangement()
    RearrangeCols ("separated")
End Sub

Sub RearrangeCols(Optional how As String = "order")
    Dim Headers As Variant, Header As Variant, Found As Range, FoundIndx As Integer, counter As Integer
    If how = "order" Then
        Headers = getPossibleHeaders()
    ElseIf how = "separated" Then
        Headers = getPossibleHeadersSeparated()
    End If
    
    counter = 1
    For i = LBound(Headers) To UBound(Headers)
        Header = Headers(i)
        Set Found = Rows(1).Find(Header, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not Found Is Nothing Then
            Debug.Print (Header)
            Debug.Print (counter)
            FoundIndx = Found.Column
            If FoundIndx <> counter Then
                Found.EntireColumn.Cut
                Columns(counter).Insert Shift:=xlToRight
            End If
            counter = counter + 1
        End If
    Next i
End Sub



Function getHeaders()
    Dim LastCol As String, rng As Variant, cell As Range, Headers As Variant
    LastCol = LastColumn()
    rng = "A1:" & LastCol
    Set rng = ActiveSheet.Range(rng)
    For Each cell In rng
        Headers = addValueToArray(Headers, cell.Value)
    Next cell
    getHeaders = Headers
End Function

Function LastColumn()
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1")
    ' LastColumn = rng.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    LastColumn = rng.End(xlToRight).Address
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
        If Not Contains(Header, Years) Then
            Years = addValueToArray(Years, Header)
        End If
        
    Next Header
    getYears = Years
    ' MsgBox CLng("13.5")
End Function






Function getAllPossibleHeaders(FiscalYear As Variant) As Variant
    Dim Headers As Variant, H As Variant, Period As Variant, Year As Variant
    Headers = getHeaders()
    Years = getYears(Headers)

    For Each Year In Years
    For Each Period In FiscalYear
        Period = Period & " " & Year
        H = addValueToArray(H, Period)
    Next Period
    Next Year
    getAllPossibleHeaders = H
End Function




Function getPossibleHeadersSeparated() As Variant
    Dim Headers1 As Variant, Headers2 As Variant
    FiscalYear = Array("Q1", "Q2", "Q3", "Q4", "Year Ended")
    Headers1 = getAllPossibleHeaders(Array("Year Ended"))
    Headers2 = getAllPossibleHeaders(Array("Q1", "Q2", "Q3", "Q4"))
    For Each H In Headers2
        Headers1 = addValueToArray(Headers1, H)
    Next H
    getPossibleHeadersSeparated = Headers1
End Function



Function getPossibleHeaders() As Variant
    Dim Headers1 As Variant, FiscalYear As Variant
    FiscalYear = Array("Q1", "Q2", "Q3", "Q4", "Year Ended")
    Headers1 = getAllPossibleHeaders(FiscalYear)
    getPossibleHeaders = Headers1
    Debug.Print Join(Headers1, "")
End Function
