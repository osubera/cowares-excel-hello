# Introduction #

  * how to create a pseudo HTML DOM tree in vba

## 概要 ##
  * VBAでHTML DOMツリーっぽいものを作成する

# Details #

  * use a recursive model written at [hello\_dom](hello_dom.md) page.
  * use a subclass written at [hello\_class\_override](hello_class_override.md) page.
```
module: test
  test1 is an executable, to see how it works.
  test2 does the same thing in a different way.
class: Node
  a single node class
class: Nodes
  child nodes class, those belong to a single direct parent node
class: X*
  rolls of each tag, to override a Node instance
```

## 説明 ##
  * [hello\_dom](hello_dom.md) ページで紹介した再帰モデル作成のテクニックを使う。
  * [hello\_class\_override](hello_class_override.md) ページで紹介したサブクラス作成のテクニックを使う。
```
module: test
  test1 が実行可能で、どんな動作をするか見れる
  test2 は同じことを別の方法で行う
class: Node
  一つのノードを表すクラス
class: Nodes
  一つの親ノードに直接ぶら下がる、子ノードを表すクラス
class: X*
  ノードのインスタンスをオーバーライドする、各タグのロールを表すクラス
```

# Code #

```
'module
'    name;test
'{{{
Option Explicit

Sub test1()
    Dim TheContent As Scripting.Dictionary
    Set TheContent = New Scripting.Dictionary
    
    TheContent("title") = "we are generating an html document"
    TheContent("keywords") = Array("Fortitudinous", "Free", "Fair")
    
    Dim x As Node
    Dim y As Node
    Dim oTextNode As XTextNode
    Dim oLi As Xli
    Dim a As Variant
    
    Set x = New Node
    Set oTextNode = New XTextNode
    Set oLi = New Xli
    
    With x
        .InitializeAs New Xhtml
        .AddChild New Node
        .AddChild New Node
    End With
    With x.Child(1)
        .InitializeAs New Xhead
        .AddChild New Node
    End With
    With x.Child(1).Child(1)
        .InitializeAs New Xtitle
        .AddChild New Node
    End With
    With x.Child(1).Child(1).Child(1)
        .InitializeAs oTextNode
        .Attributes("text") = TheContent("title")
    End With
    With x.Child(2)
        .InitializeAs New Xbody
        .AddChild New Node
        .AddChild New Node
    End With
    With x.Child(2).Child(1)
        .InitializeAs New Xh1
        .AddChild New Node
    End With
    With x.Child(2).Child(1).Child(1)
        .InitializeAs oTextNode
        .Attributes("text") = TheContent("title")
    End With
    With x.Child(2).Child(2)
        .InitializeAs New Xol
        For Each a In TheContent("keywords")
            Set y = New Node
            y.InitializeAs oLi
            y.AddChild New Node
            y.Child(1).InitializeAs oTextNode
            y.Child(1).Attributes("text") = a
            .AddChild y
            Set y = Nothing
        Next
    End With
    
    Debug.Print x.ToTextHtml
    Debug.Print x.ToTextPlain
    
    TheContent.RemoveAll
    x.Clear
End Sub

Sub test2()
    Dim a As Variant
    Dim x As Node
    Dim y As Node
    Dim oTextNode As XTextNode
    Dim oLi As Xli
    Dim TheStructure As Variant
    Dim TheId As Scripting.Dictionary
    Dim TheContent As Scripting.Dictionary
    Set TheContent = New Scripting.Dictionary
    Set TheId = New Scripting.Dictionary
    Set oTextNode = New XTextNode
    Set oLi = New Xli
    
    TheContent("title") = "we are generating an html document"
    TheContent("keywords") = Array("Fortitudinous", "Free", "Fair")
    
    TheStructure = Array(New Xhtml, 0, Array(New Xhead, 0, Array(New Xtitle, 1)), _
                                Array(New Xbody, 0, Array(New Xh1, 2), Array(New Xol, 3)))
    
    Set x = MakeTree(TheStructure, TheId)
    
    For Each a In Array(TheId(1), TheId(2))
        Set y = New Node
        y.InitializeAs oTextNode
        y.Attributes("text") = TheContent("title")
        a.AddChild y
        Set y = Nothing
    Next
    
    For Each a In TheContent("keywords")
        Set y = New Node
        y.InitializeAs oLi
        y.AddChild New Node
        y.Child(1).InitializeAs oTextNode
        y.Child(1).Attributes("text") = a
        TheId(3).AddChild y
        Set y = Nothing
    Next
    
    Debug.Print x.ToTextHtml
    Debug.Print x.ToTextPlain
    
    TheId.RemoveAll
    TheContent.RemoveAll
    x.Clear
End Sub

Private Function MakeTree(TheStructure As Variant, TheId As Scripting.Dictionary) As Node
    Dim i As Long
    Dim x As Node
    Set x = New Node
    x.InitializeAs TheStructure(0)
    Set TheId(TheStructure(1)) = x
    For i = 2 To UBound(TheStructure)
        x.AddChild MakeTree(TheStructure(i), TheId)
    Next
    Set MakeTree = x
End Function
'}}}
 
'class
'    name;Node
'{{{
Option Explicit

Private TheResources As Scripting.Dictionary
Private TheNode As Scripting.Dictionary
Private TheChild As Nodes
Private TheProperties As Scripting.Dictionary

' make some methods overridable: begin

Private TheRolls As Collection

Private Function TryOverride(ByRef Result As Variant, ProcName As String, CallType As VbCallType _
                            , Optional Args As Variant = 0) As Boolean
    Dim i As Long
    For i = TheRolls.Count To 1 Step -1
        If TheRolls(i).RespondTo(ProcName) Then
            Result = CallByName(TheRolls(i), ProcName, CallType, Array(TheResources, Args))
            TryOverride = True
            Exit Function
        End If
    Next
    TryOverride = False
End Function

Public Sub AddRoll(Roll As Variant)
    TheRolls.Add Roll
End Sub

Public Sub RemoveRoll(i As Long)
    TheRolls.Remove i
End Sub

Private Sub InitializeRolls()
    Set TheRolls = New Collection
End Sub

Private Sub ClearRolls()
    On Error Resume Next
    Do While TheRolls.Count > 0
        TheRolls.Remove 1
    Loop
    Set TheRolls = Nothing
End Sub

' make some methods overridable: end

Private Sub Class_Initialize()
    Set TheNode = New Scripting.Dictionary
    Set TheChild = New Nodes
    Set TheProperties = New Scripting.Dictionary
    Set TheResources = New Scripting.Dictionary
    SetResources
    InitializeRolls
End Sub

Private Sub Class_Terminate()
    ClearRolls
    ClearResources
End Sub

Private Sub SetResources()
    Set TheResources("node") = TheNode
    Set TheResources("children") = TheChild
    Set TheResources("properties") = TheProperties
    Set TheResources("parent") = Nothing
    Set TheResources("this") = Me
    TheResources("tablength") = 2
    TheResources("linefeed") = vbCrLf
    TheResources("lang") = "ja"
    TheResources("encoding") = "UTF-8"
    TheResources("declare_xml") = "<?xml version=""1.0"" encoding=""" & TheResources("encoding") & """?>"
    TheResources("doctype_xhtmlbasic1.1") = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML Basic 1.1//EN"" ""http://www.w3.org/TR/xhtml-basic/xhtml-basic11.dtd"">"
    TheResources("doctype") = TheResources("doctype_xhtmlbasic1.1")
    SetProperties Array(Array("tagname", "o"))
End Sub

Private Sub ClearResources()
    On Error Resume Next
    TheResources.RemoveAll
    Set TheResources = Nothing
    TheProperties.RemoveAll
    Set TheProperties = Nothing
    Clear
    Set TheNode = Nothing
End Sub

Private Sub Alert(Message As Variant)
    Debug.Print Message
End Sub

Public Sub Clear()
    TheChild.Clear
    TheNode.RemoveAll
End Sub

Public Property Get HasParent() As Boolean
    HasParent = IIf(Not TheResources("parent") Is Nothing, True, False)
End Property

Public Property Get Parent() As Node
    Set Parent = TheResources("parent")
End Property

Public Property Set Parent(ParentNode As Node)
    Set TheResources("parent") = ParentNode
End Property

Public Property Get Children() As Nodes
    Set Children = TheChild
End Property

Public Property Get Child(i As Long) As Node
    Set Child = TheChild.Item(i)
End Property

Public Property Get HasChild() As Boolean
    HasChild = IIf(TheChild.Count > 0, True, False)
End Property

Public Sub AddChild(ChildNode As Node)
    Set ChildNode.Parent = Me
    TheChild.Add ChildNode
End Sub

Public Sub InsertChild(ChildNode As Node, Before As Long)
    Set ChildNode.Parent = Me
    TheChild.Insert ChildNode, Before
End Sub

Public Sub RemoveChild(i As Long)
    TheChild.Remove (i)
End Sub

Public Property Get Attributes(Name As String) As String
    If TheNode.Exists(Name) Then
        Attributes = TheNode(Name)
    Else
        Alert "Warning: no attributes for the name. " & Name
        Attributes = ""
    End If
End Property

Public Property Let Attributes(Name As String, Text As String)
    TheNode(Name) = Text
End Property

Public Sub SetAttributes(Settings As Variant)
    Dim Pair As Variant
    For Each Pair In Settings
        TheNode(Pair(0)) = Pair(1)
    Next
End Sub

Public Sub ClearAttributes()
    TheNode.RemoveAll
End Sub

Public Property Get Properties(Name As String) As Variant
    Properties = TheProperties(Name)
End Property

Public Property Let Properties(Name As String, Value As Variant)
    TheProperties(Name) = Value
End Property

Public Sub SetProperties(Settings As Variant)
    Dim Pair As Variant
    For Each Pair In Settings
        TheProperties(Pair(0)) = Pair(1)
    Next
End Sub

Public Sub ClearProperties()
    TheProperties.RemoveAll
End Sub

Public Sub InitializeAs(Roll As Variant)
    AddRoll Roll
    ReInitialize
End Sub

Public Sub ReInitialize()
    ' make it overridable
    Dim Result As Variant
    If TryOverride(Result, "ReInitialize", VbMethod) Then Exit Sub
    
    ' base code begins here
    ClearProperties
    ClearAttributes
    SetProperties Array(Array("tagname", "o"))
End Sub

Public Function ToText(Optional PaddingLeft As Long = 0) As String
    ' overridable begin
    Dim Result As Variant
    If TryOverride(Result, "ToText", VbMethod, Array(PaddingLeft)) Then
        ToText = Result
        Exit Function
    End If
    
    ' base code begins here
    Dim Padding As String
    Padding = Space(PaddingLeft)
    ToText = Padding & "<" & TheProperties("tagname") & ToTextAttributes & ">" & TheResources("linefeed") _
        & ToTextChildren(PaddingLeft + TheResources("tablength")) _
        & Padding & "</" & TheProperties("tagname") & ">" & TheResources("linefeed")
End Function

Public Function ToTextAttributes() As String
    ' overridable begin
    Dim Result As Variant
    If TryOverride(Result, "ToTextAttributes", VbMethod) Then
        ToTextAttributes = Result
        Exit Function
    End If
    
    ' base code begins here
    Dim Text As String
    Dim Key As Variant
    Text = ""
    For Each Key In TheNode.Keys
        Text = Text & " " & CStr(Key) & "=""" & TheNode(Key) & """"
    Next
    ToTextAttributes = Text
End Function

Public Function ToTextChildren(PaddingLeft As Long, Optional ProcName As String = "ToText") As String
    ' overridable begin
    Dim Result As Variant
    If TryOverride(Result, "ToTextChildren", VbMethod, Array(PaddingLeft, ProcName)) Then
        ToTextChildren = Result
        Exit Function
    End If
    
    ' base code begins here
    Dim Text As String
    Dim i As Long
    For i = 1 To TheChild.Count
        TheChild.Item(i).Properties("list_counter") = i
        Text = Text & CallByName(TheChild.Item(i), ProcName, VbMethod, PaddingLeft)
    Next
    ToTextChildren = Text
End Function

Public Function ToTextPlain(Optional PaddingLeft As Long = 0) As String
    ' overridable begin
    Dim Result As Variant
    If TryOverride(Result, "ToTextPlain", VbMethod, Array(PaddingLeft)) Then
        ToTextPlain = Result
        Exit Function
    End If
    
    ' base code begins here
    ToTextPlain = ToTextChildren(PaddingLeft, "ToTextPlain")
End Function

Public Function ToTextHtml(Optional PaddingLeft As Long = 0) As String
    ' overridable begin
    Dim Result As Variant
    If TryOverride(Result, "ToTextHtml", VbMethod, Array(PaddingLeft)) Then
        ToTextHtml = Result
        Exit Function
    End If
    
    ' base code begins here
    Dim Padding As String
    Padding = Space(PaddingLeft)
    ToTextHtml = Padding & "<" & TheProperties("tagname") & ToTextAttributes & ">" & TheResources("linefeed") _
        & ToTextChildren(PaddingLeft + TheResources("tablength"), "ToTextHtml") _
        & Padding & "</" & TheProperties("tagname") & ">" & TheResources("linefeed")
End Function
'}}}
 
'class
'    name;Nodes
'{{{
Option Explicit

Private TheResources As Scripting.Dictionary
Private TheNodes As Collection

Private Sub Class_Initialize()
    Set TheNodes = New Collection
    Set TheResources = New Scripting.Dictionary
    SetResources
End Sub

Private Sub Class_Terminate()
    ClearResources
End Sub

Private Sub SetResources()
    Set TheResources("nodes") = TheNodes
End Sub

Private Sub ClearResources()
    On Error Resume Next
    TheResources.RemoveAll
    Set TheResources = Nothing
    Clear
    Set TheNodes = Nothing
End Sub

Private Sub Alert(Message As Variant)
    Debug.Print Message
End Sub

Private Function IsValidItemNumber(i As Long) As Boolean
    If (i < 1 Or i > TheNodes.Count) Then
        Alert "Warning: not a valid item number. " & i
        IsValidItemNumber = False
        Exit Function
    End If
    IsValidItemNumber = True
End Function

Public Sub Clear()
    Do While TheNodes.Count > 0
        TheNodes(1).Clear
        TheNodes.Remove 1
    Loop
End Sub

Public Property Get Count() As Long
    Count = TheNodes.Count
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = IIf(TheNodes.Count = 0, True, False)
End Property

Public Property Get Item(i As Long) As Node
    If IsValidItemNumber(i) Then
        Set Item = TheNodes(i)
    Else
        Set Item = Nothing
    End If
End Property

Public Sub Add(TheNode As Node)
    TheNodes.Add TheNode
End Sub

Public Sub Insert(TheNode As Node, i As Long)
    If IsValidItemNumber(i) Then
        TheNodes.Add TheNode, Before:=i
    End If
End Sub

Public Sub Remove(i As Long)
    If IsValidItemNumber(i) Then
        TheNodes(i).Clear
        TheNodes.Remove i
    End If
End Sub
'}}}
 
'class
'    name;Xhtml
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize", "ToTextHtml")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "html"))
        .SetAttributes Array(Array("xmlns", "http://www.w3.org/1999/xhtml"), _
            Array("xml:lang", Args(0)("lang")))
    End With
    ReInitialize = 0    ' must return something
End Function

Public Function ToTextHtml(Args As Variant) As String
    Dim TheResources As Scripting.Dictionary
    Dim TheProperties As Scripting.Dictionary
    Dim PaddingLeft As Long
    Dim Padding As String
    Dim DeclareXml As String
    Set TheResources = Args(0)
    Set TheProperties = TheResources("properties")
    PaddingLeft = Args(1)(0)
    Padding = Space(PaddingLeft)
    DeclareXml = TheResources("declare_xml") & TheResources("linefeed") & Padding _
            & TheResources("doctype") & TheResources("linefeed") & Padding
    With TheResources("this")
        ToTextHtml = Padding & DeclareXml _
            & "<" & TheProperties("tagname") & .ToTextAttributes & ">" & TheResources("linefeed") _
            & .ToTextChildren(PaddingLeft + TheResources("tablength"), "ToTextHtml") _
            & Padding & "</" & TheProperties("tagname") & ">" & TheResources("linefeed")
    End With
End Function
'}}}
 
'class
'    name;XTextNode
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize", "ToTextPlain", "ToTextHtml")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "TextNode"))
    End With
    ReInitialize = 0    ' must return something
End Function

Public Function ToTextPlain(Args As Variant) As String
    ToTextPlain = Args(0)("this").Attributes("text")
End Function

Public Function ToTextHtml(Args As Variant) As String
    Dim TheResources As Scripting.Dictionary
    Dim PaddingLeft As Long
    Dim Padding As String
    Set TheResources = Args(0)
    PaddingLeft = Args(1)(0)
    Padding = Space(PaddingLeft)
    With TheResources("this")
        ToTextHtml = Padding & .Attributes("text") & TheResources("linefeed")
    End With
End Function
'}}}
 
'class
'    name;Xhead
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize", "ToTextPlain")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "head"))
    End With
    ReInitialize = 0    ' must return something
End Function

Public Function ToTextPlain(Args As Variant) As String
    ToTextPlain = ""
End Function
'}}}
 
'class
'    name;Xbody
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "body"))
    End With
    ReInitialize = 0    ' must return something
End Function
'}}}
 
'class
'    name;Xtitle
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "title"))
    End With
    ReInitialize = 0    ' must return something
End Function
'}}}
 
'class
'    name;Xh1
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize", "ToTextPlain")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "h1"))
    End With
    ReInitialize = 0    ' must return something
End Function

Public Function ToTextPlain(Args As Variant) As String
    Dim PaddingLeft As Long
    PaddingLeft = Args(1)(0)
    With Args(0)("this")
        ToTextPlain = Space(PaddingLeft) & "= " _
            & .ToTextChildren(PaddingLeft, "ToTextPlain") _
            & " =" & Args(0)("linefeed") & Args(0)("linefeed")
    End With
End Function
'}}}
 
'class
'    name;Xol
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "ol"))
    End With
    ReInitialize = 0    ' must return something
End Function
'}}}
 
'class
'    name;Xli
'{{{
Option Explicit

' this is to override an instance of Node class

Private TheResponses As Variant

Public Function RespondTo(ProcName) As Boolean
    Dim Name As Variant
    For Each Name In TheResponses
        If Name = ProcName Then
            RespondTo = True
            Exit Function
        End If
    Next
    RespondTo = False
End Function

Private Sub Class_Initialize()
    ' keep names to override
    TheResponses = Array("ReInitialize", "ToTextPlain")
End Sub

Public Function ReInitialize(Args As Variant) As Long
    With Args(0)("this")
        .ClearProperties
        .ClearAttributes
        .SetProperties Array(Array("tagname", "li"))
    End With
    ReInitialize = 0    ' must return something
End Function

Public Function ToTextPlain(Args As Variant) As String
    Dim TheResources As Scripting.Dictionary
    Dim PaddingLeft As Long
    Dim Marker As String
    PaddingLeft = Args(1)(0)
    Set TheResources = Args(0)
    Marker = " " & CStr(TheResources("properties")("list_counter")) & ". "
    With TheResources("this")
        ToTextPlain = Space(PaddingLeft) & Marker _
            & .ToTextChildren(PaddingLeft, "ToTextPlain") _
            & TheResources("linefeed")
    End With
End Function
'}}}
```

results

  * we have an html document and a plain text document.
  * the html passes a validation at [w3c unicorn](http://validator.w3.org/unicorn/) .

```
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML Basic 1.1//EN" "http://www.w3.org/TR/xhtml-basic/xhtml-basic11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja">
  <head>
    <title>
      we are generating an html document
    </title>
  </head>
  <body>
    <h1>
      we are generating an html document
    </h1>
    <ol>
      <li>
        Fortitudinous
      </li>
      <li>
        Free
      </li>
      <li>
        Fair
      </li>
    </ol>
  </body>
</html>

= we are generating an html document =

 1. Fortitudinous
 2. Free
 3. Fair
```