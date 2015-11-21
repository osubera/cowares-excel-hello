

# Introduction #

  * how to parse a text stream in vba

## 概要 ##
  * VBAでテキストストリームを解析するパーサーを作る

# Details #

  * create a parser engine class, to make easy generating other classes for verious rules.
  * added ReadLineParse method into the StringStream class at [hello\_stream](hello_stream.md).
  * UtilText.DivideAtFirstMatch is a string search function. it's simple and dull. we want ths separated, because some parsers may better use the regular expression instead of this one.
```
module: Module1
  executable tests, to see how it works.
module: UtilText
  public functions called by HtmlParser*.
class: HtmlParserT1
  an example parser class for simple html
class: StreamParserEngine
  a parser engine class, to controls streams and the parser class
class: StringStream
  make streams to work with the engine

in additional code

module: Module2
  executable tests, to see how it works.
class: HtmlParserT2
  an example parser class for simple html
```

## 説明 ##
  * パーサーエンジンを作って、ルールに応じたパーサークラスを作りやすくする
  * [hello\_stream](hello_stream.md) で紹介した StringStream クラスに ReadLineParse メソッドを追加した
  * UtilText.DivideAtFirstMatch は文字列検索関数で、非常に単純な総当り検索をしている。パースの文法によっては、正規表現を使うなど別のものに差し替えやすいように切り離している。
```
module: Module1
  test を実行して動作確認をする
module: UtilText
  HtmlParser* から呼ばれる関数
class: HtmlParserT1
  html を使う単純なパーサークラスのサンプル
class: StreamParserEngine
  パーサーエンジンクラスで、ストリームとパーサークラスを制御する
class: StringStream
  エンジンと共に動くストリームを作る

追加コードには次のものがある。

module: Module2
  test を実行して動作確認をする
class: HtmlParserT2
  html を使う単純なパーサークラスのサンプル
```

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. try to run test code in Module1.
  1. include additional codes into the same book.

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. Module1 のテストマクロを試す。
  1. 追加コードは同じブックに読み込む。

# Code #

```
'workbook
'  name;stream_parse_engine.xls

'require
'  ;{420B2830-E718-11CF-893D-00A0C9054228} 1 0 Microsoft Scripting Runtime


'module
'  name;UtilText
'{{{
Option Explicit

Public Function DivideAtFirstMatch( _
    ByRef MatchedTag As String, ByRef BeforeTag As String, ByRef AfterTag As String, _
    Text As String, Tags As Variant, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    ' 配列でもらった Tag のいずれかの、最初の位置に見つかったもので、文字列を前後に分けて返す。
    ' MatchedTag: 実際に見つかったタグ、見つからないときは "" を返す。
    ' BeforeTag:  Text の Tag より前の部分 を返す。
    ' AfterTag:   Text の Tag より後の部分 を返す。
    ' Text:       検索対象の文字列を指定する。
    ' Tags:       検索するタグを指定する。 Array("<b>","<i>") など
    ' 戻り値:     タグが１つも無ければ false
    ' 同一順位なら、先に指定されたタグを優先する。 ("<b","<br" のようにどちらにも一致する場合)
    
    Dim Tag As Variant
    Dim At As Long
    Dim AtFirst As Long
    Dim AtTag As String
    
    If Not IsArray(Tags) Then
        DivideAtFirstMatch = DivideAtFirstMatch(MatchedTag, BeforeTag, AfterTag, Text, Array(CStr(Tags)))
        Exit Function
    End If
    
    AtFirst = Len(Text) + 1
    AtTag = ""
    For Each Tag In Tags
        If Tag = "" Then GoTo Ignore
        At = InStr(1, Text, CStr(Tag), Compare)
        If At = 0 Then GoTo Ignore
        If At >= AtFirst Then GoTo Ignore
        
        AtFirst = At
        AtTag = CStr(Tag)
Ignore:
    Next
    
    MatchedTag = AtTag
    BeforeTag = Left(Text, AtFirst - 1)
    AfterTag = Right(Text, Len(Text) - AtFirst + 1 - Len(AtTag))
    DivideAtFirstMatch = (AtTag <> "")
End Function
'}}}

'class
'  name;StreamParseEngine
'{{{
Option Explicit

Private ComeIn As Object                        ' stream comes in
Private GoOut As Object                         ' stream goes out
Private Parser As Object                        ' parser to be used

Private Sub Class_Initialize()
    Set Parser = Me
End Sub

Private Sub Class_Terminate()
    Set ComeIn = Nothing
    Set GoOut = Nothing
    Set Parser = Nothing
End Sub

Public Function GetStreamIn() As Object
    Set GetStreamIn = ComeIn
End Function

Public Sub SetStreamIn(Stream As Object)
    Set ComeIn = Stream
End Sub

Public Function GetStreamOut() As Object
    Set GetStreamOut = GoOut
End Function

Public Sub SetStreamOut(Stream As Object)
    Set GoOut = Stream
End Sub

Public Function GetParser() As Object
    Set GetParser = Parser
End Function

Public Sub SetParser(UseThisParser As Object)
    Set Parser = UseThisParser
End Sub

Public Function Parse(Optional MyKey As String = "") As Boolean
    Dim FoundTag As String
    Dim FoundKey As String
    Dim Text As String
    Dim ShallExit As Boolean
    Dim FuncStart As String
    Dim FuncContent As String
    Dim FuncEnd As String
    
    If Not Parser.GetFunctionName(FuncStart, FuncContent, FuncEnd, MyKey) Then
        Parse = False
        Exit Function
    End If
    
    FilteredWrite FuncStart, MyKey
    Do Until ComeIn.AtEndOfStream
        ShallExit = ComeIn.ReadLineParse(Text, FoundTag, FoundKey, Parser, MyKey)
        FilteredWrite FuncContent, Text
        If ShallExit Then Exit Do
        If FoundKey <> "" Then Parse FoundKey
    Loop
    FilteredWrite FuncEnd, MyKey
    
    Parse = True
End Function

Private Function FilteredWrite(FunctionName As String, Optional Text As String = "") As Boolean
    Dim Filtered As String
    
    FilteredWrite = False
    If FunctionName <> "" Then
        Filtered = CallByName(Parser, FunctionName, VbMethod, Text)
        If Filtered <> "" Then
            GoOut.WriteLine Filtered
            FilteredWrite = True
        End If
    End If
End Function

' offer the default parser here, as a nop parser

Public Function GetFunctionName( _
        ByRef FuncStart As String, ByRef FuncContent As String, ByRef FuncEnd As String, _
        Key As String) As Boolean
    FuncStart = ""
    FuncContent = "Parse_default"
    FuncEnd = ""
    GetFunctionName = True
End Function

Public Function Parse_default(Text As String) As String
    Parse_default = Text
End Function

Public Function SearchTags( _
        ByRef MatchedTag As String, ByRef BeforeTag As String, ByRef AfterTag As String, _
        ByRef FoundKey As String, ByRef FoundEnd As Boolean, _
        MyKey As String, Text As String) As Boolean
    
    MatchedTag = ""
    BeforeTag = Text
    AfterTag = ""
    FoundKey = ""
    FoundEnd = False
    SearchTags = False
End Function
'}}}

'class
'  name;StringStream
'{{{
Option Explicit

' wrap other modules into a stream class, do like a Scripting.TextStream

Private MyText As Collection
Private MyLineFeed As String
Private RememberFlush As String

Private Sub Class_Initialize()
    Set MyText = New Collection
    MyLineFeed = vbCrLf
    RememberFlush = ""
End Sub

Private Sub Class_Terminate()
    Set MyText = Nothing
End Sub

Public Function OpenTextRead(Optional Text As Variant) As Boolean
    If Not IsMissing(Text) Then Enqueue Text
    RememberFlush = ""
    OpenTextRead = True
End Function

Public Function OpenTextWrite(Optional Append As Boolean = False) As Boolean
    If Not Append Then ClearAll
    RememberFlush = ""
    OpenTextWrite = True
End Function

Public Function OpenText() As Boolean
    ClearAll
    RememberFlush = ""
    OpenText = True
End Function

Public Sub CloseText()
    If RememberFlush <> "" Then
        CallByName Me, RememberFlush, VbMethod
    End If
    ClearAll
    RememberFlush = ""
End Sub

Public Property Get LineFeed() As String
    LineFeed = MyLineFeed
End Property

Public Property Let LineFeed(Text As String)
    MyLineFeed = Text
End Property

Public Property Get AtEndOfStream() As Boolean
    AtEndOfStream = IsEmpty
End Property

Public Function ReadLine() As String
    ReadLine = Dequeue
End Function

Public Function ReadAll() As String
    ReadAll = ToText
    ClearAll
End Function

' "Write" is a reserved word in VBA, so we use this
Public Sub WriteLine(ParamArray Text() As Variant)
    Dim x As Variant
    For Each x In Text
        Enqueue x
    Next
End Sub

Public Sub WriteTextArray(Texts As Variant)
    EnqueueArray Texts
End Sub

Private Sub ClearAll()
    Do While MyText.Count > 0
        MyText.Remove 1
    Loop
End Sub

Private Sub EnqueueArray(Texts As Variant)
    Dim Text As Variant
    For Each Text In Texts
        Enqueue Text
    Next
End Sub

Private Sub Enqueue(Text As Variant)
    Dim Splitted As Variant
    Dim x As Variant
    Splitted = Split(Text, MyLineFeed)
    For Each x In Splitted
        MyText.Add CStr(x)
    Next
End Sub

Private Function Dequeue() As String
    Dequeue = MyText(1)
    MyText.Remove 1
End Function

Private Function CheatQueue() As String
    CheatQueue = MyText(1)
End Function

Private Function EditFirstQueue(Text As String) As String
    MyText.Add Text, After:=1
    EditFirstQueue = Dequeue
End Function

Private Function IsEmpty() As Boolean
    IsEmpty = (MyText.Count = 0)
End Function

Private Function ToText() As String
    Dim Result As String
    Dim i As Long
    For i = 1 To MyText.Count
        Result = Result & MyText(i) & MyLineFeed
    Next
    ToText = Result
End Function

' add methods for StreamParseEngine

Public Function ReadLineParse( _
        ByRef Text As String, ByRef FoundTag As String, ByRef FoundKey As String, _
        Parser As Object, MyKey As String) As Boolean
    
    Dim Cheat As String
    Dim BeforeTag As String
    Dim AfterTag As String
    Dim Found As Boolean
    Dim FoundEnd As Boolean
    
    Cheat = CheatQueue
    Found = Parser.SearchTags(FoundTag, BeforeTag, AfterTag, FoundKey, FoundEnd, MyKey, Cheat)
    If Found Then
        EditFirstQueue AfterTag
        Text = BeforeTag
    Else
        Text = ReadLine
    End If
    
    ReadLineParse = FoundEnd Or AtEndOfStream
End Function

' wrap XXXX: begin
Public Function OpenXXXXRead(Optional Args As Variant) As Boolean
    ClearAll
    RememberFlush = ""
    ' do something to get a text from XXXX, like below
    'Enqueue CopyFromXXXX
    OpenXXXXRead = True
End Function

Public Function OpenXXXXWrite(Optional Args As Variant) As Boolean
    ClearAll
    RememberFlush = "ToXXXX"
    OpenXXXXWrite = True
End Function

Public Sub ToXXXX()
    'do something to save the text into XXXX, like below
    'CopyToXXXX ToText
End Sub

' wrap XXXX: end

'}}}

'module
'  name;Module1
'{{{
Option Explicit

Sub test1()
    Dim p As HtmlParserT1
    Set p = New HtmlParserT1
    
    Debug.Print p.Description
    Debug.Print p.DescriptionEn
    Debug.Print Dump(p.Root)
    Debug.Print Dump(p.Keys)
    Debug.Print Dump(p.StartTag("html"))
    Debug.Print Dump(p.EndTag("html"))
    Debug.Print Dump(p.ChildKeys("html"))
    Debug.Print Dump(p.StartTags(p.Keys))
    
    Set p = Nothing
End Sub

Sub test2()
    Dim e As StreamParseEngine
    Dim StIn As StringStream
    Dim StOut As StringStream
    
    Dim p As HtmlParserT1
    Set p = New HtmlParserT1
    
    Set e = New StreamParseEngine
    Set StIn = New StringStream
    Set StOut = New StringStream
    
    e.SetStreamIn StIn
    e.SetStreamOut StOut
    e.SetParser p
    
    StOut.LineFeed = ""
    StOut.OpenTextWrite
    StIn.OpenTextRead "<html><Body>" & _
        "<table><tr><td><li>irregal li</li></td></tr><tr><td><script>javascript;</script></td></tr></table>" & _
        "<ol><li>#1 hello</li><li>#2 world</li></ol>out sider" & _
        "<ul><li>see 3 items?</li><li id=""w"">has an attribute</li><li/></ul>" & _
        "</boDY></html>"
    e.Parse
    StIn.CloseText
    
    Debug.Print StOut.ReadAll
    StOut.CloseText
    
    e.SetParser Nothing
    e.SetStreamIn Nothing
    e.SetStreamOut Nothing
    Set StIn = Nothing
    Set StOut = Nothing
    Set p = Nothing
    Set e = Nothing
End Sub

Function Dump(x As Variant) As String
    Dim a As Variant
    Dim b As String
    If IsArray(x) Then
        b = "("
        For Each a In x
            b = b & Dump(a) & ","
        Next
        b = b & ")"
    Else
        b = x
    End If
    Dump = CStr(b)
End Function
'}}}

'class
'  name;HtmlParserT1
'{{{
Option Explicit

Public Property Get Description() As String
    Description = _
"HTMLみたいなタグを試す。属性は使えない。"
End Property

Public Property Get DescriptionEn() As String
    DescriptionEn = _
"to begin with a tiny HTML. no attributes supported. "
End Property

Friend Property Get CompareMethod() As VbCompareMethod
    CompareMethod = vbTextCompare
End Property

Friend Property Get Root() As String
    Root = Keys(0)
End Property

Friend Property Get Keys() As Variant
    Keys = Array("html", "ol", "ul", "li")
End Property

Friend Property Get ChildKeys(Key As String) As Variant
    Dim out As Variant
    Select Case Key
    Case "html"
        out = Array("ol", "ul")
    Case "ol"
        out = Array("li")
    Case "ul"
        out = Array("li")
    Case "li"
        out = Array()
    End Select
    ChildKeys = out
End Property

Friend Property Get StartTag(Key As String) As String
    StartTag = "<" & Key & ">"
End Property

Friend Property Get EndTag(Key As String) As String
    EndTag = "</" & Key & ">"
End Property

Friend Property Get StartTags(Keys As Variant) As Variant
    Dim i As Long
    Dim out() As String
    If UBound(Keys) = -1 Then
        StartTags = Array()
    Else
        ReDim out(LBound(Keys) To UBound(Keys))
        For i = LBound(Keys) To UBound(Keys)
            out(i) = StartTag(CStr(Keys(i)))
        Next
        StartTags = out
    End If
End Property

Friend Property Get EndAndChildTags(Key As String) As Variant
    Dim children As Variant
    Dim i As Long
    Dim out() As String
    
    children = StartTags(ChildKeys(Key))
    ReDim out(LBound(children) To UBound(children) + 1)
    out(LBound(out)) = EndTag(Key)
    For i = LBound(out) + 1 To UBound(out)
        out(i) = children(i - 1)
    Next
    
    EndAndChildTags = out
End Property

' called directly by SearchTags() : begin

Friend Function ActiveTags(Key As String) As Variant
    If Key = "" Then
        ActiveTags = Array(StartTag(Root))
    Else
        ActiveTags = EndAndChildTags(Key)
    End If
End Function

Friend Function KeyFromTag(Tag As String, ParentKey As String) As String
    Dim ChildKey As Variant
    
    If ParentKey = "" Then
        KeyFromTag = Root
    Else
        For Each ChildKey In ChildKeys(ParentKey)
            If StartTag(CStr(ChildKey)) = Tag Then
                KeyFromTag = ChildKey
                Exit Function
            End If
        Next
        KeyFromTag = ""
    End If
End Function

Friend Function IsEndTag(Tag As String, Key As String) As Boolean
    IsEndTag = (Tag = EndTag(Key))
End Function

' called directly by SearchTags() : end

Public Function SearchTags( _
        ByRef MatchedTag As String, ByRef BeforeTag As String, ByRef AfterTag As String, _
        ByRef FoundKey As String, ByRef FoundEnd As Boolean, _
        MyKey As String, Text As String) As Boolean
    
    Dim Found As Boolean
    
    Found = DivideAtFirstMatch(MatchedTag, BeforeTag, AfterTag, Text, ActiveTags(MyKey), CompareMethod)
    If Found Then
        FoundEnd = IsEndTag(MatchedTag, MyKey)
        If FoundEnd Then
            FoundKey = MyKey
        Else
            FoundKey = KeyFromTag(MatchedTag, MyKey)
        End If
    Else
        FoundEnd = False
        FoundKey = ""
    End If
    
    SearchTags = Found
End Function

Public Function GetFunctionName( _
        ByRef FuncStart As String, ByRef FuncContent As String, ByRef FuncEnd As String, _
        Key As String) As Boolean
    
    Dim out As String
    
    Select Case Key
    Case "li"
        out = "Parse_default"
    Case "", "html"
        out = ""
    Case Else
        out = "Parse_" & Key
    End Select
    
    FuncStart = ""
    FuncContent = out
    FuncEnd = ""
    GetFunctionName = True
End Function

Public Function Parse_default(Text As String) As String
    Parse_default = Text
End Function

Public Function Parse_ol(Text As String) As String
    Parse_ol = vbCrLf & "[ordered list]"
End Function

Public Function Parse_ul(Text As String) As String
    Parse_ul = vbCrLf & "[unordered list]"
End Function

'}}}

```

### Result ###

  * elements `<li>` with attributes disappeared by design
> > 属性つきの `<li>` が消えたのは仕様による
  * the blank tag `<li/>` disappeared by design
> > 空タグ `<li/>` が消えたのは仕様による
  * the irregal `<li>` directed to `<body>` disappeared, because the restriction of children is working
> > 子要素を限定しているので、 `<body>` 直下の不正な `<li>` は消えた。

```
test1()

HTMLみたいなタグを試す。属性は使えない。
to begin with a tiny HTML. no attributes supported. 
html
(html,ol,ul,li,)
<html>
</html>
(ol,ul,)
(<html>,<ol>,<ul>,<li>,)

test2()

[ordered list]#1 hello
[ordered list]#2 world
[ordered list]
[unordered list]see 3 items?
[unordered list]
```

# More Code #

```
'class
'  name;HtmlParserT2
'{{{
Option Explicit

Public Property Get Description() As String
    Description = _
"HTMLみたいなタグを試す。属性を仮のタグに置き換えてみる。"
End Property

Public Property Get DescriptionEn() As String
    DescriptionEn = _
"to begin with a tiny HTML. make attributes into a virtual tag. "
End Property

Friend Property Get CompareMethod() As VbCompareMethod
    CompareMethod = vbTextCompare
End Property

Friend Property Get Root() As String
    Root = Keys(0)
End Property

Friend Property Get Keys() As Variant
    Keys = Array("html", "body", "head", "br", "div", "p", "span", _
                "dl", "dt", "dd", "ol", "ul", "li", "table", "td", "th", "tr", "script", _
                "attributes")
End Property

Friend Property Get ChildKeys(Key As String) As Variant
    Select Case Key
    Case "attributes"
        ChildKeys = Array()
    Case "script"
        ChildKeys = Array("attributes")
    Case Else
        ' can include every tag
        ChildKeys = Keys
    End Select
End Property

Friend Property Get StartTag(Key As String) As String
    StartTag = "<" & Key & ">"
End Property

Friend Property Get EndTag(Key As String) As String
    EndTag = "</" & Key & ">"
End Property

Friend Property Get StartTags(Keys As Variant) As Variant
    Dim i As Long
    Dim out() As String
    If UBound(Keys) = -1 Then
        StartTags = Array()
    Else
        ReDim out(LBound(Keys) To UBound(Keys))
        For i = LBound(Keys) To UBound(Keys)
            out(i) = StartTag(CStr(Keys(i)))
        Next
        StartTags = out
    End If
End Property

Friend Property Get EndAndChildTags(Key As String) As Variant
    Dim children As Variant
    Dim i As Long
    Dim out() As String
    
    children = StartTags(ChildKeys(Key))
    ReDim out(LBound(children) To UBound(children) + 1)
    out(LBound(out)) = EndTag(Key)
    For i = LBound(out) + 1 To UBound(out)
        out(i) = children(i - 1)
    Next
    
    EndAndChildTags = out
End Property

' add attributes with the virtual key "attr-*"

Private Function AttributesKey() As String
    AttributesKey = Keys(UBound(Keys))
End Function

Private Function AttributesTag() As String
    AttributesTag = "<" & AttributesKey & ">"
End Function

Private Function IsAttr(Key As String) As Boolean
    IsAttr = (Key = AttributesKey)
End Function

Private Function ContainsAttr(Tag As String) As Boolean
    ContainsAttr = (Right(Tag, 1) = " ")
End Function

Private Function ContainsEnd(Tag As String) As Boolean
    ContainsEnd = (Right(Tag, 1) = "/")
End Function

Private Function TagWithAttr(Tag As String) As String
    TagWithAttr = Left(Tag, Len(Tag) - 1) & " "
End Function

Private Function TagWithEnd(Tag As String) As String
    TagWithEnd = Left(Tag, Len(Tag) - 1) & "/"
End Function

Private Function LooksLikeAnEndTag(Tag As String) As Boolean
    LooksLikeAnEndTag = (Left(Tag, 2) = "</")
End Function

Private Function TagWithoutAttrEnd(Tag As String) As String
    Dim out As String
    If ContainsEnd(Tag) Then
        out = Left(Tag, Len(Tag) - 1) & ">"
    ElseIf ContainsAttr(Tag) Then
        out = Left(Tag, Len(Tag) - 1) & ">"
    Else
        out = Tag
    End If
    TagWithoutAttrEnd = out
End Function

Private Function TagsWithAttrEnd(Tags As Variant) As Variant
    Dim i As Long
    Dim out() As String
    Dim x As Collection
    Dim Tag As Variant
    Dim strTag As String
    
    If UBound(Tags) = -1 Then
        TagsWithAttrEnd = Array()
    Else
        Set x = New Collection
        x.Add "</>"  ' wild card
        For Each Tag In Tags
            strTag = CStr(Tag)
            x.Add strTag
            If Not LooksLikeAnEndTag(strTag) Then
                x.Add TagWithAttr(strTag)
                x.Add TagWithEnd(strTag)
            End If
        Next
        ReDim out(0 To x.Count - 1)
        For i = 0 To x.Count - 1
            out(i) = x(i + 1)
        Next
        Set x = Nothing
        TagsWithAttrEnd = out
    End If
    
End Function

' called directly by SearchTags() : begin

Friend Function ActiveTags(Key As String) As Variant
    Dim out As Variant
    If Key = "" Then
        out = TagsWithAttrEnd(Array(StartTag(Root)))
    ElseIf IsAttr(Key) Then
        out = Array("/>", ">")
    Else
        out = TagsWithAttrEnd(EndAndChildTags(Key))
    End If
    ActiveTags = out
End Function

Friend Function KeyFromTag(Tag As String, ParentKey As String) As String
    Dim ChildKey As Variant
    
    If ParentKey = "" Then
        KeyFromTag = Root
    Else
        For Each ChildKey In ChildKeys(ParentKey)
            If StartTag(CStr(ChildKey)) = TagWithoutAttrEnd(Tag) Then
                KeyFromTag = ChildKey
                Exit Function
            End If
        Next
        KeyFromTag = ""
    End If
End Function

Friend Function IsEndTag(Tag As String, Key As String) As Boolean
    If Tag = ">" Or Tag = "/>" Or Tag = "</>" Then
        IsEndTag = True
    ElseIf Tag = EndTag(Key) Then
        IsEndTag = True
    Else
        IsEndTag = False
    End If
End Function

' called directly by SearchTags() : end

Public Function SearchTags( _
        ByRef MatchedTag As String, ByRef BeforeTag As String, ByRef AfterTag As String, _
        ByRef FoundKey As String, ByRef FoundEnd As Boolean, _
        MyKey As String, Text As String) As Boolean
    
    Dim Found As Boolean
    Dim Tag As String
    
    'Debug.Print MyKey, Text
    Found = DivideAtFirstMatch(MatchedTag, BeforeTag, AfterTag, Text, ActiveTags(MyKey), CompareMethod)
    If Found Then
        FoundEnd = IsEndTag(MatchedTag, MyKey)
        If FoundEnd Then
            FoundKey = MyKey
            If MatchedTag = "/>" Then
                ' insert a wild card end tag for parent
                AfterTag = "</>" & AfterTag
            End If
        Else
            FoundKey = KeyFromTag(MatchedTag, MyKey)
            If ContainsAttr(MatchedTag) Then
                ' insert <attributes> as a first child
                AfterTag = AttributesTag & AfterTag
            ElseIf ContainsEnd(MatchedTag) Then
                ' insert <attributes> as a first child, and make an end tag
                AfterTag = AttributesTag & "/" & AfterTag
            End If
        End If
    Else
        FoundEnd = False
        FoundKey = ""
    End If
    
    SearchTags = Found
End Function

Public Function GetFunctionName( _
        ByRef FuncStart As String, ByRef FuncContent As String, ByRef FuncEnd As String, _
        Key As String) As Boolean
    
    Dim outS As String
    Dim out As String
    Dim outE As String
    
    outS = "Parse_tagged_start"
    out = "Parse_default"
    outE = "Parse_tagged_end"
    
    Select Case Key
    Case "", "head", "attributes"
        outS = ""
        out = ""
        outE = ""
    Case "span", "table", "tr"
        outS = ""
        outE = ""
    Case "th", "td", "div", "br"
        outS = ""
        outE = "Parse_br"
    Case "script"
        outS = ""
        out = "Parse_" & Key
        outE = ""
    End Select
    
    FuncStart = outS
    FuncContent = out
    FuncEnd = outE
    GetFunctionName = True
End Function

Public Function Parse_default(Text As String) As String
    Parse_default = Text
End Function

Public Function Parse_linefeed(Text As String) As String
    Parse_linefeed = vbCrLf
End Function

Public Function Parse_tagged_start(Text As String) As String
    Parse_tagged_start = StartTag(Text)
End Function

Public Function Parse_tagged_end(Text As String) As String
    Parse_tagged_end = EndTag(Text) & vbCrLf
End Function

Public Function Parse_slash(Text As String) As String
    Parse_slash = " / "
End Function

Public Function Parse_br(Text As String) As String
    Parse_br = "<br/>" & vbCrLf
End Function

Public Function Parse_script(Text As String) As String
    Parse_script = "[scripts removed]"
End Function

'}}}

'module
'  name;Module2
'{{{
Option Explicit

Sub test3()
    Dim p As HtmlParserT2
    Set p = New HtmlParserT2
    
    Debug.Print p.Description
    Debug.Print p.DescriptionEn
    Debug.Print Dump(p.Root)
    Debug.Print Dump(p.Keys)
    Debug.Print Dump(p.StartTag("html"))
    Debug.Print Dump(p.EndTag("html"))
    Debug.Print Dump(p.ChildKeys("html"))
    Debug.Print Dump(p.StartTags(p.Keys))
    
    Set p = Nothing
End Sub

Sub test4()
    Dim e As StreamParseEngine
    Dim StIn As StringStream
    Dim StOut As StringStream
    
    Dim p As HtmlParserT2
    Set p = New HtmlParserT2
    
    Set e = New StreamParseEngine
    Set StIn = New StringStream
    Set StOut = New StringStream
    
    e.SetStreamIn StIn
    e.SetStreamOut StOut
    e.SetParser p
    
    StOut.LineFeed = ""
    StOut.OpenTextWrite
    StIn.OpenTextRead "<html><Body>" & _
        "<table><tr><td><li>irregal li</li></td></tr><tr><td><script>javascript;</script></td></tr></table>" & _
        "<ol><li>#1 hello</li><li>#2 world</li></ol>out sider" & _
        "<ul><li>see 3 items?</li><li id=""w"">has an attribute</li><li/></ul>" & _
        "</boDY></html>"
    e.Parse
    StIn.CloseText
    
    Debug.Print StOut.ReadAll
    StOut.CloseText
    
    e.SetParser Nothing
    e.SetStreamIn Nothing
    e.SetStreamOut Nothing
    Set StIn = Nothing
    Set StOut = Nothing
    Set p = Nothing
    Set e = Nothing
End Sub

'}}}

```

### Result ###

  * can parse tags with attributes
> > 属性つきのタグが解釈できるようにした
  * can parse the blank tag, like `<li/>`
> > `<li/>` のような空タグを解釈できるようにした

```
test3()

HTMLみたいなタグを試す。属性を仮のタグに置き換えてみる。
to begin with a tiny HTML. make attributes into a virtual tag. 
html
(html,body,head,br,div,p,span,dl,dt,dd,ol,ul,li,table,td,th,tr,script,attributes,)
<html>
</html>
(html,body,head,br,div,p,span,dl,dt,dd,ol,ul,li,table,td,th,tr,script,attributes,)
(<html>,<body>,<head>,<br>,<div>,<p>,<span>,<dl>,<dt>,<dd>,<ol>,<ul>,<li>,<table>,<td>,<th>,<tr>,<script>,<attributes>,)

test4()

<html><body><li>irregal li</li>
<br/>
[scripts removed]<br/>
<ol><li>#1 hello</li>
<li>#2 world</li>
</ol>
out sider<ul><li>see 3 items?</li>
<li>has an attribute</li>
<li></li>
</ul>
</body>
</html>
```
