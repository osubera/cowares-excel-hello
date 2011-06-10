' ssf_to_html
' convert ssf to html table
' Copyright (C) 2011 Tomizono - kobobau.mocvba.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo ssf_to_html.vbs /t Title /e:Charset FILE

Private Env
Private LastBlockName
Private CellStream

Set Args = WScript.Arguments
Main Args.Named, Args.Unnamed(0)
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Public Sub Main(Opts, File)
    Dim Charset, Caption
    Caption = Opts("t")
    Charset = Opts("e")
    If Charset = "" Then Charset = "utf-8"
    WScript.Echo FileReader(File, Charset, Caption)
End Sub

Private Sub CellsReadFrom(BlockName, Block)
    With CellStream
        .BlockName = BlockName
        .ReadBlock Block
    End With
End Sub

Private Sub ReadFrom(BlockName, Block)
    Select Case BlockName
    Case "cells-text", "cells-formula", "cells-color", "cells-background-color"
        CellsReadFrom BlockName, Block
    End Select
End Sub

Public Function FileReader(FileName, Charset, Caption)
    Const adTypeText = 2
    Dim Stream
    
    If FileName = "" Then Exit Function
    InitializeEnv
    
    Set Stream = CreateObject("ADODB.Stream")
    With Stream
        .Open
        .Type = adTypeText
        .Charset = Charset
        .LoadFromFile FileName
        ReadSsf Stream
        .Close
    End With
    Set Stream = Nothing
    
    If Caption <> "" Then CellStream.Caption = Caption
    FileReader = CellStream.ToHtml
    TerminateEnv
End Function

Public Sub InitializeEnv()
    Set Env = New ArrayBuilder
    Set CellStream = New HtmlCellStream
End Sub

Public Sub TerminateEnv()
    Set CellStream = Nothing
    Set Env = Nothing
End Sub

Public Sub ReadSsf(Stream)
    Dim Finder
    
    Set Finder = New StreamParser
    Set Finder.Stream = Stream
    ParseSsf Finder
    
    Set Finder = Nothing
End Sub

Public Sub ParseSsf(Finder)
    Do
        If Not BeginTheMagic(Finder) Then Exit Sub
        EvalAfter Finder
        EndTheMagic
    Loop While RequireTheMagic And Not Finder.AtEndOfStream
End Sub

Private Function RequireTheMagic()
    RequireTheMagic = False
End Function

Private Sub DoFlush()
    If LastBlockName <> "" Then
        ReadFrom LastBlockName, Env.GetArray()
    End If
End Sub

Private Sub EndTheMagic()
    DoFlush
End Sub

Private Function BeginTheMagic(Finder)
    Dim i, EndBegins, At
    Dim Bch, Ech
    Dim MagicBegin
    
    BeginTheMagic = True
    
    If Not RequireTheMagic Then Exit Function
    MagicBegin = "ssf-begin"
    
    BeginTheMagic = False
    
    At = Finder.FindString(1, MagicBegin, vbBinaryCompare)
    If At <= 1 Then Exit Function
    
    Finder.Text = Mid(Finder.Text, At - 1)
    
    Bch = Left(Finder.Text, 1)
    EndBegins = Len(MagicBegin) + 2
    i = Finder.FindString(EndBegins, Bch, vbBinaryCompare)
    If i <= EndBegins Then Exit Function
    
    Ech = Mid(Finder.Text, EndBegins, i - EndBegins)
    'Env.SetEnv "ssf", "line-begin", Bch
    'Env.SetEnv "ssf", "line-end", Ech
    
    BeginTheMagic = True
End Function

Private Sub SetSpecialChars(ByRef Bch, ByRef Ech, ByRef EscapeBegin, ByRef MagicEnd, ByRef Delimiter)
    EscapeBegin = "{{{"
    MagicEnd = "ssf-end"
    Delimiter = ";"
    Bch = "'"   ' default bch
    EscapeOff Bch, Ech
End Sub

Private Sub EscapeOn(ByRef Bch, ByRef Ech)
    Ech = Ech & Bch & "}}}" & Ech
End Sub

Private Sub EscapeOff(ByRef Bch, ByRef Ech)
    Ech = vbCrLf    ' default ech
End Sub

Private Sub EvalAfter(Finder)
    Dim Bch, Ech
    Dim EscapeBegin, MagicEnd, Delimiter
    Dim BlockName, Key, Value, BeforeTag
    Dim BlockVoid
    Dim EchNextBlock
    
    SetSpecialChars Bch, Ech, EscapeBegin, MagicEnd, Delimiter
    
    BlockName = ""
    EchNextBlock = Ech & Bch
    Do Until Finder.AtEndOfStream
        EvalComma BeforeTag, Finder, Bch, Ech
        EvalEscape BeforeTag, Finder, Bch, Ech, EscapeBegin, Delimiter
        EvalBefore BeforeTag, Bch, Delimiter, BlockName, Key, Value, BlockVoid
        
        If BlockName = "" Then
            Finder.Text = Ech & Finder.Text
            EvalComma BeforeTag, Finder, Bch, EchNextBlock
            If Finder.Text <> "" Then Finder.Text = Bch & Finder.Text
        ElseIf BlockName = MagicEnd Then
            Exit Do
        End If
    Loop
End Sub

Private Sub EvalComma(ByRef BeforeTag, Finder, Bch, Ech)
    Dim At
    
    At = Finder.FindString(1, Ech, vbBinaryCompare)
    If At = 0 Then
        BeforeTag = Finder.Text
        Finder.Text = ""
    Else
        BeforeTag = Left(Finder.Text, At - 1)
        Finder.Text = Right(Finder.Text, Len(Finder.Text) - At + 1 - Len(Ech))
    End If
End Sub

Private Sub EvalEscape(ByRef BeforeTag, Finder, Bch, Ech, EscapeBegin, Delimiter)
    Dim MyBeforeTag
    
    If BeforeTag <> Bch & EscapeBegin Then Exit Sub
    
    EscapeOn Bch, Ech
    EvalComma MyBeforeTag, Finder, Bch, Ech
    BeforeTag = Bch & Delimiter & MyBeforeTag
    EscapeOff Bch, Ech
End Sub

Private Sub EvalBefore(BeforeTag, Bch, Delimiter, ByRef BlockName, ByRef Key, ByRef Value, ByRef BlockVoid)
    Dim KeyValue
    
    If BeforeTag = "" Then
        BlockName = ""
    ElseIf Left(BeforeTag, 1) <> Bch Then
        BlockName = ""
    ElseIf BlockName = "" Then
        DoFlush
        BlockName = Mid(BeforeTag, 2)
        Key = ""
        Value = ""
        BlockVoid = False
        LastBlockName = BlockName
    ElseIf Not BlockVoid Then
        KeyValue = Split(BeforeTag, Delimiter, 2, vbBinaryCompare)
        Key = Trim(Replace(Mid(KeyValue(0), 2), vbTab, ""))
        If UBound(KeyValue) = 1 Then
            Value = KeyValue(1)
        Else
            Value = ""
        End If
        If Key = "void" Then
            BlockVoid = True
        Else
            Env.AddArray Array(Key, Value)
        End If
    End If
End Sub


' Excel Columns are 26 decimal, but each digit begins at 1
' our inside Key is [Row Number],[Col Number]

' map 1..26 into A..Z
Public Function N2A(Number)
    Const Achar = 65    ' A
    N2A = ChrW(Number + Achar - 1)
End Function

' map A..Z into 1..26
Public Function A2N(Alphabet)
    Const Achar = 65    ' A
    A2N = AscW(UCase(Alphabet)) - Achar + 1
End Function

' Column String to number
Public Function Col2Num(ColString)
    Dim i, Num
    Num = 0
    For i = 1 To Len(ColString)
        Num = 26 * Num + A2N(Mid(ColString, i, 1))
    Next
    Col2Num = Num
End Function

' number to Column String
Public Function Num2Col(ByVal Number)
    Dim Col, x
    Col = ""
    Do While Number > 0
        x = ((Number - 1) Mod 26) + 1
        Col = N2A(x) & Col
        Number = (Number - x) / 26
    Loop
    Num2Col = Col
End Function

' A1 to array
Public Function A1RowCol(A1)
    Dim R, m
    Dim Row, Col
    Set R = RegA1
    Set m = R.Execute(A1)
    If m.Count = 0 Then
        A1RowCol = Array(0, 0)
    Else
        Col = Col2Num(m(0).SubMatches(0))
        Row = CLng(m(0).SubMatches(1))
        A1RowCol = Array(Row, Col)
    End If
    Set m = Nothing
    Set R = Nothing
End Function

' array to A1
Public Function RowColA1(RowCol)
    If UBound(RowCol) <> 1 Then
        RowColA1 = ""
    ElseIf RowCol(0) = 0 Then
        RowColA1 = Num2Col(RowCol(1))
    Else
        RowColA1 = Num2Col(RowCol(1)) & RowCol(0)
    End If
End Function

' R1C1 to array
Public Function R1C1RowCol(R1C1)
    Dim R, m
    Dim Row, Col
    Set R = RegR1C1
    Set m = R.Execute(R1C1)
    If m.Count = 0 Then
        R1C1RowCol = Array(0, 0)
    Else
        Row = CLng(m(0).SubMatches(0))
        Col = CLng(m(0).SubMatches(1))
        R1C1RowCol = Array(Row, Col)
    End If
    Set m = Nothing
    Set R = Nothing
End Function

' array to R1C1
Public Function RowColR1C1(RowCol)
    If UBound(RowCol) <> 1 Then
        RowColR1C1 = ""
    Else
        RowColR1C1 = "R" & RowCol(0) & "C" & RowCol(1)
    End If
End Function

' Key to array
Public Function KeyRowCol(Key)
    Dim Row, Col
    Dim x
    x = Split(Key, ",")
    If UBound(x) <> 1 Then
        KeyRowCol = Array(0, 0)
    Else
        KeyRowCol = Array(CLng(x(0)), CLng(x(1)))
    End If
End Function

' array to Key
Public Function RowColKey(RowCol)
    If UBound(RowCol) <> 1 Then
        RowColKey = ""
    Else
        RowColKey = RowCol(0) & "," & RowCol(1)
    End If
End Function

' extract Range size
Public Function RangeSize(Address, ByRef Row1, ByRef Row2, ByRef Col1, ByRef Col2)
    Dim i, y, x(1)
    
    y = RangeStartEnd(Address)
    For i = 0 To 1
        x(i) = R1C1RowCol(y(i))
        If x(i)(1) = 0 Then x(i) = A1RowCol(y(i))
    Next
    
    Row1 = x(0)(0)
    Col1 = x(0)(1)
    Row2 = x(1)(0)
    Col2 = x(1)(1)
    
    If Col1 = 0 Or Col2 = 0 Then
        RangeSize = 0
    Else
        RangeSize = Col2 - Col1 + 1
    End If
End Function

' extract Range
Public Function RangeStartEnd(Address)
    Dim StartAt, EndAt
    Dim x
    x = Split(Address, ":")
    Select Case UBound(x)
    Case -1
        StartAt = ""
        EndAt = ""
    Case 0
        StartAt = Address
        EndAt = StartAt
    Case Else
        StartAt = x(0)
        EndAt = x(1)
    End Select
    RangeStartEnd = Array(StartAt, EndAt)
End Function

' extract A1 format
Private Function RegA1()
    Dim R
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = False
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "([a-zA-Z]+)\$?([0-9]*)"
    
    Set RegA1 = R
End Function

' extract R1C1 absolute format
Private Function RegR1C1()
    Dim R
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = False
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "[rR]([0-9]+)[cC]([0-9]+)"
    
    Set RegR1C1 = R
End Function


' classes

Class ArrayBuilder

Public RawData
Public DefaultValue

Public Function PopData(Key)
    If RawData.Exists(Key) Then
        PopData = RawData(Key)
        RawData.Remove Key
    Else
        PopData = DefaultValue
    End If
End Function

Public Function GetData(Key)
    If RawData.Exists(Key) Then
        GetData = RawData(Key)
    Else
        GetData = DefaultValue
    End If
End Function

Public Sub SetData(Key, Value)
    RawData(Key) = Value
End Sub

Public Sub Clear()
    RawData.RemoveAll
End Sub

Public Function GetArray()
    Dim Key, Count, out()
    
    GetArray = Array()
    
    Count = RawData.Count
    If Count > 0 Then
        ReDim out(Count - 1)
        For Key = 1 To Count
            out(Key - 1) = RawData(Key)
            RawData.Remove Key
        Next
        GetArray = out
    End If
End Function

Public Sub AddArray(Value)
    Dim Key
    
    Key = RawData.Count + 1
    RawData(Key) = Value
End Sub

Private Sub Class_Initialize()
    Const TextCompare = 1
    
    Set RawData = CreateObject("Scripting.Dictionary")
    RawData.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    RawData.RemoveAll
    Set RawData = Nothing
End Sub

End Class

Class StreamParser

Public Text
Public Stream

Public Property Get EOS()
    EOS = Stream.EOS
End Property

Public Property Get AtEndOfStream()
    AtEndOfStream = EOS And (Text = "")
End Property

Public Function MoreText()
    If EOS Then Exit Function
    
    Const BuffSize = 8192
    Dim out
    out = Stream.ReadText(BuffSize)
    Text = Text & out
    MoreText = out
End Function

Public Function FindString(StartAt, Search, CompareMethod)
    Dim out, more, At, Require
    
    Require = StartAt + Len(Search) - 1
    Do While Len(Text) < Require
        more = MoreText
        If more = "" Then Exit Do
    Loop
    
    out = InStr(StartAt, Text, Search, CompareMethod)
    Do While out = 0
        At = Len(Text) - Len(Search) + 2
        more = MoreText
        If more = "" Then Exit Do
        
        out = InStr(At, Text, Search, CompareMethod)
    Loop
    
    FindString = out
End Function

Private Sub Class_Initialize()
    Text = ""
End Sub

End Class

Class HtmlCellStream

Public Text
Public Color
Public BackgroundColor
Public CurrentMatrix
Public Caption

Public LocalKey

Public Property Let BlockName(NewName)
    LocalKey = NewName
    
    Select Case LocalKey
    Case "cells-text", "cells-formula"
        Set CurrentMatrix = Text
    Case "cells-color"
        Set CurrentMatrix = Color
    Case "cells-background-color"
        Set CurrentMatrix = BackgroundColor
    End Select
End Property

Public Sub ReadBlock(Block)
    Dim KeyValue, Key, Value
    
    On Error Resume Next
    
    CurrentMatrix.RepeatCell 1
    
    For Each KeyValue In Block
        ExtractKeyValue KeyValue, Key, Value
        ReadSsfLine Key, Value
        
        If Err.Number <> 0 Then
            WScript.Echo Err.Number & " " & Err.Description & "(" & Key & "," & Value & ")"
            Err.Clear
        End If
    Next
End Sub

Private Sub ReadSsfLine(Key, Value)
    Select Case Key
    Case "address"
        CurrentMatrix.SetRange Value
    Case "repeat"
        CurrentMatrix.RepeatCell Int(CLng(Value))
    Case "skip"
        CurrentMatrix.SkipCell Int(CLng(Value))
    Case ""
        CurrentMatrix.SetCell Value
    End Select
End Sub

Public Function ToHtml()
    Const ExcelTABLE = "border:medium ridge #66cc99;background-color:white;border-collapse:collapse;"
    Const ExcelTH = "border:thin outset white;background-color:#c7c7c7;color:black;padding-left:5px;padding-right:5px;"
    Const ExcelTD = "border:thin solid #c7c7c7;background-color:white;color:black;"
    Const ExcelCAPTION = "border:none;background-color:#66cc99;color:#7e434e;"
    
    Dim HasText, HasColor, HasBackgroundColor
    Dim TextData
    Dim ColumnsHeader, RowsHeader
    Dim R, C
    Dim CellKey, CellStyle, CellColor, CellBackgroundColor
    Dim out
    
    HasText = (Text.ColumnsCount > 0)
    If Not HasText Then Exit Function
    
    HasColor = (Color.ColumnsCount > 0)
    HasBackgroundColor = (BackgroundColor.ColumnsCount > 0)
    
    Text.DefaultValue = " "
    TextData = Text.GetArray
    ColumnsHeader = Text.GetColumnsHeader
    RowsHeader = Text.GetRowsHeader
    
    Set out = New StringStream
    
    out.WriteLine "<table style=""" & ExcelTABLE & """>"
    out.WriteLine " <tr>"
    out.WriteLine "  <td colspan=""" & CStr(UBound(ColumnsHeader) + 2) & """ style=""" & ExcelCAPTION & """>" & Caption & "</td>"
    out.WriteLine " </tr>"
    out.WriteLine " <tr>"
    out.WriteLine "  <th style=""" & ExcelTH & """>&nbsp;</th>"
    For C = 0 To UBound(ColumnsHeader)
        out.WriteLine "  <th style=""" & ExcelTH & """>" & ColumnsHeader(C) & "</th>"
    Next
    out.WriteLine " </tr>"
    For R = 0 To UBound(RowsHeader)
        out.WriteLine " <tr>"
        out.WriteLine "  <th style=""" & ExcelTH & """>" & RowsHeader(R) & "</th>"
        For C = 0 To UBound(ColumnsHeader)
            CellKey = RowColKey(Array(R + Text.Row1, C + Text.Col1))
            CellColor = Color.GetData(CellKey)
            CellBackgroundColor = BackgroundColor.GetData(CellKey)
            CellStyle = ExcelTD
            If CellColor <> "" Then CellStyle = CellStyle & "color:" & CellColor & ";"
            If CellBackgroundColor <> "" Then CellStyle = CellStyle & "background-color:" & CellBackgroundColor & ";"
            out.WriteText "  <td"
            If CellStyle <> "" Then out.WriteText " style=""" & CellStyle & """"
            out.WriteLine ">" & SafeString(TextData(R)(C)) & "</td>"
        Next
        out.WriteLine " </tr>"
    Next
    out.WriteLine "</table>"
    
    ToHtml = out.Text
    Set out = Nothing
End Function

Private Function SafeString(ByVal x)
    x = Replace(x, "&", "&amp;")
    x = Replace(x, "<", "&lt;")
    x = Replace(x, ">", "&gt;")
    x = Replace(x, " ", "&nbsp;")
    SafeString = x
End Function

Private Function ExtractKeyValue(KeyValue, ByRef Key, ByRef Value)
    Key = KeyValue(0)
    Value = KeyValue(1)
    ExtractKeyValue = Key
End Function

Private Sub Class_Initialize()
    Set Text = New MatrixCells
    Set Color = New MatrixCells
    Set BackgroundColor = New MatrixCells
    Caption = "&nbsp;"
End Sub

Private Sub Class_Terminate()
    Set BackgroundColor = Nothing
    Set Color = Nothing
    Set Text = Nothing
End Sub

End Class

Class MatrixCells

Public RawData
Public DefaultValue

Public ColumnsCount
Public Row1
Public Row2
Public Col1
Public Col2
Public CurrentCol1
Public CurrentCol2

Public R
Public C

Private RepeatCount

Public Function PopArray()
    PopArray = GetArray
    Clear
End Function

Public Function GetArray()
    Dim Rs, Cs
    Dim i, j
    Dim EachRow(), AllRows()
    
    Cs = ColumnsCount
    If Cs = 0 Then Cs = 1
    If R < Row2 Then R = Row2
    Rs = R - Row1 + 1
    
    ReDim EachRow(Cs - 1)
    ReDim AllRows(Rs - 1)
    
    For i = 0 To Rs - 1
        For j = 0 To Cs - 1
            EachRow(j) = GetData(RowColKey(Array(i + Row1, j + Col1)))
        Next
        AllRows(i) = EachRow
    Next
    
    GetArray = AllRows
End Function

Public Function GetColumnsHeader()
    Dim Cs, Header()
    Dim i
    Cs = ColumnsCount
    If Cs = 0 Then Cs = 1
    
    ReDim Header(Cs - 1)
    
    For i = 0 To Cs - 1
        Header(i) = RowColA1(Array(0, i + Col1))
    Next
    
    GetColumnsHeader = Header
End Function

Public Function GetRowsHeader()
    Dim Rs, Header()
    Dim i
    If R < Row2 Then R = Row2
    Rs = R - Row1 + 1
    
    ReDim Header(Rs - 1)
    
    For i = 0 To Rs - 1
        Header(i) = i + Row1
    Next
    
    GetRowsHeader = Header
End Function

Public Sub SetRange(Address)
    Dim R1, R2, C1, C2
    If RangeSize(Address, R1, R2, C1, C2) = 0 Then Exit Sub
    
    If ColumnsCount = 0 Then
        Col1 = C1
        Col2 = C2
        Row1 = R1
        Row2 = R2
    Else
        If Col1 > C1 Then Col1 = C1
        If Col2 < C2 Then Col2 = C2
        If Row1 > R1 Then Row1 = R1
        If Row2 < R2 Then Row2 = R2
    End If
    
    CurrentCol1 = C1
    CurrentCol2 = C2
    ColumnsCount = Col2 - Col1 + 1
    R = R1
    C = C1
End Sub

Public Sub SetCell(Value)
    Do While RepeatCount > 0
        SetData RowColKey(Array(R, C)), Value
        NextCell
        RepeatCount = RepeatCount - 1
    Loop
    RepeatCount = 1
End Sub

Public Sub RepeatCell(Count)
    RepeatCount = Count
End Sub

Public Sub SkipCell(ByVal Count)
    Do While Count > 0
        NextCell
        Count = Count - 1
    Loop
End Sub

Public Sub NextCell()
    If C = CurrentCol2 Then
        C = CurrentCol1
        R = R + 1
    Else
        C = C + 1
    End If
End Sub

Public Function PopData(Key)
    If RawData.Exists(Key) Then
        PopData = RawData(Key)
        RawData.Remove Key
    Else
        PopData = DefaultValue
    End If
End Function

Public Function GetData(Key)
    If RawData.Exists(Key) Then
        GetData = RawData(Key)
    Else
        GetData = DefaultValue
    End If
End Function

Public Sub SetData(Key, Value)
    RawData(Key) = Value
End Sub

Public Sub Clear()
    If RawData.Count > 0 Then RawData.RemoveAll
    R = 1
    C = 1
    RepeatCount = 1
    Row1 = 1
    Col1 = 1
    Row2 = 0
    Col2 = 0
    CurrentCol1 = 1
    CurrentCol2 = 0
    ColumnsCount = 0
End Sub

Private Sub Class_Initialize()
    Const TextCompare = 1
    
    Set RawData = CreateObject("Scripting.Dictionary")
    RawData.CompareMode = TextCompare
    Clear
End Sub

Private Sub Class_Terminate()
    RawData.RemoveAll
    Set RawData = Nothing
End Sub

End Class

Class StringStream

Public Text

Public Sub WriteLine(Data)
    Text = Text & Data & vbCrLf
End Sub

Public Sub WriteText(Data)
    Text = Text & Data
End Sub

End Class
