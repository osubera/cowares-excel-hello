

# Introduction #

  * show a tiny clock on excel, word and access

## 概要 ##
  * エクセル、ワード、アクセスに小さな時計を表示する

![http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oZy5-i2I/AAAAAAAAACE/mR2QLzZYfow/s1600/shot9.png](http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oZy5-i2I/AAAAAAAAACE/mR2QLzZYfow/s1600/shot9.png)

# Details / 説明 #

  * shows a toolbar with clock.
> > 時計つきツールバーを表示する。
    * when the document is opened for Excel and Word.
> > > エクセルとワードでは、文書を開いたときに
    * when the form HelloNowControler is opned for Access.
> > > アクセスでは、 HelloNowControler フォームを開いたときに
  * uses common core parts for Excel, Word and Access.

> > エクセルとワードとアクセスで、共通の本体コードを使う。
  * common parts / 共通コード
```
class: HelloNow
  core clock functionals
class: ToolBarV2
  tool bar helper
```
  * for Excel / エクセル用
```
module: HelloNowMain
  callback and low level functions to work with excel
code: ThisWorkbook
  handles events on open and close
```
  * for Word / ワード用
```
module: HelloNowMain
  callback and low level functions to work with word
code: ThisDocument
  handles events on open and close
```
  * for Access / アクセス用
```
module: HelloNowMain
  callback and low level functions to work with access
code: Form_HelloNowControler
  control the clock bar on opening, closing and updating
macro: HelloNowMain
  callback from toolbar
```


# How to use / 使い方 #

  1. show the clock bar ( or an addin ribbon for office 2007 and above )
> > 時計バーを表示する (オフィス2007以降ではリボンのアドイン)
    1. a toolbar appears when you open the excel book or the word document
> > > エクセルブックかワード文書を開くと、ツールバーが表示される。
    1. a toolbar appears when you open the `HelloNowControler` form  in access mdb
> > > アクセスMDBファイルでは、 `HelloNowControler` フォームを開くと、ツールバーが表示される。
> > > ![http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oWhYQHwI/AAAAAAAAABs/OTeZX_oLQmg/s1600/shot3.png](http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oWhYQHwI/AAAAAAAAABs/OTeZX_oLQmg/s1600/shot3.png)
  1. work with the clock bar

> > 時計バーの操作
> > ![http://2.bp.blogspot.com/_EUW0nrj9XlM/TT4oX5fCNmI/AAAAAAAAAB0/ogyWE-J9iIY/s1600/shot5.png](http://2.bp.blogspot.com/_EUW0nrj9XlM/TT4oX5fCNmI/AAAAAAAAAB0/ogyWE-J9iIY/s1600/shot5.png)
    1. the time.  a look can be customized.  copy the time formatted into a clipboard by clicking this button.
> > > 今の日時を表示する。見た目は設定変更できる。このボタンを押すと、書式のままの日時をクリップボードにコピーする。
    1. the time formatted is put into active cell, active document, active table, or so on.
> > > 書式通りの日時を、セルや文書やテーブルなどに挿入する。
    1. show or hide buttons for settings, 4th and later.
> > > ４番目以降の設定ボタンを表示または隠す。
> > > ![http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oYSAvrOI/AAAAAAAAAB4/hcrURVtBLyM/s1600/shot6.png](http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oYSAvrOI/AAAAAAAAAB4/hcrURVtBLyM/s1600/shot6.png)
    1. use zenkaku characters in output when turned on.
> > > これがオンのとき、全角文字で出力する。
    1. ignore formats in output when turned on.
> > > これがオンのとき、書式を無視して出力する。
    1. select a format or enter a new one.  accepts date time format strings.
> > > 日時の書式を選ぶか、新たに入力する。 Format 関数の書式文字が使える。
> > > ![http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oY-vFuKI/AAAAAAAAAB8/GBOtDo-8xDw/s1600/shot7.png](http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oY-vFuKI/AAAAAAAAAB8/GBOtDo-8xDw/s1600/shot7.png)
    1. select a cycle of updating the time shown in a bar.
> > > バーの日時を更新する頻度を指定する。
> > > ![http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oZdbNhBI/AAAAAAAAACA/RwUS5zOpRnQ/s1600/shot8.png](http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oZdbNhBI/AAAAAAAAACA/RwUS5zOpRnQ/s1600/shot8.png)
  1. customizing

> > カスタマイズ
    1. you can customize an appearance of the clock bar and initial values at the `ButtonData` property in the `HelloNow` class.
> > > 時計バーの見た目や初期値を変更したい場合、 `HelloNow` クラスにある `ButtonData` プロパティでカスタマイズできる。
```
' people speaking English may prefer this than the original Japanese one.

Public Property Get ButtonData() As Variant
    ButtonData = Array( _
        Array("clock", "copy the date into the clipboard", "now", Empty, 1, 2, Empty, 1), _
        Array("enter", "insert the date into the active document", "enter", Empty, 1, 2), _
        Array("...", "show or hide buttons for settings", "customize", Empty, 1, 2, Empty, 1, Empty, 1), _
        Array("raw", "insert a raw date time value unformatted", "datatype", Empty, 1, 2), _
        Array("format", "choose or enter a format of the date and time", "format", Empty, 4, Empty, 4, Empty, Empty, "dddd am/pmh", "d mmmm", "c", "ddddd", "dddddd", "ttttt", "yyyy/m/d h:nn:ss", "d dd y ww ddd dddd aaa aaaa w", "m mm mmm mmmm oooo q", "g gg ggg e ee yy yyyy", "h hh n nn s ss :/", "AM/PM am/pm A/P a/p AMPM"), _
        Array("cycle", "choose a cycle to update the time", "update", Empty, 3, Empty, Empty, Empty, Empty, "daily", "hourly", "minutely", "secondly") _
        )
End Property
```


# Downloads #

  * [downloads](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=hello_now)

# Code #

### Full Application for Excel ###

```
'workbook
'  name;hello_now.xls

'require
'  ;{0D452EE1-E08F-101A-852E-02608C4D0BB4} 2 0 Microsoft Forms 2.0 Object Library

'worksheet
'  name;Sheet1


'class
'  name;ToolBarV2
'{{{
Option Explicit

' Generate an application toolbar

Private MyBar As Office.CommandBar
Private MyName As String
Private MyApp As Application


'=== main procedures helper begin ===


' this will called by pressing a button
Friend Sub BarMain(Optional oWho As Object = Nothing)
    Dim oAC As Object   ' this is the button itself pressed
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    ' switch to a main menu procedure
    Main oAC, SomebodyOrMe(oWho)
    Set oAC = Nothing
End Sub

' main menu procedure. if you delete this, a public Main in Standard Module will be called, maybe.
Private Sub Main(oAC As Object, Optional oWho As Object = Nothing)
    ' use a button tag to switch a procedure to be called as "Menu_xx"
    CallByName SomebodyOrMe(oWho), "Menu_" & oAC.Tag, VbMethod, oAC
End Sub

Public Sub Menu_about(oAC As Object)
    MsgBox TypeName(Me), vbOKOnly, "Sample of procedure called by the Main"
End Sub

Friend Sub OnButtonToggle()
    Dim oAC As Object   ' toggle this button
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    
    ButtonSwitchToggle oAC
    Set oAC = Nothing
End Sub

Private Function SomebodyOrMe(oWho As Object) As Object
    If oWho Is Nothing Then
        Set SomebodyOrMe = Me
    Else
        Set SomebodyOrMe = oWho
    End If
End Function


'=== main procedures helper end ===
'=== event procedures begin ===


Private Sub Class_Initialize()
    Set MyApp = Application
    MyName = CStr(Timer)    ' random name, maybe uniq
End Sub

Private Sub Class_Terminate()
    Set MyApp = Nothing
End Sub


'=== event procedures end ===
'=== construction and destruction begin ===


Public Sub NewBar(ParamArray Addins() As Variant)
    DelBar
    Set MyBar = CreateBar(MyApp, MyName)
    AddAddins MyBar, CVar(Addins)
    ShowBar MyBar
End Sub

Public Sub DelBar()
    DeleteBar MyBar
    Set MyBar = Nothing
End Sub

Public Sub SetApplication(oApp As Application)
    Set MyApp = oApp
End Sub

Public Sub SetName(NewName As String)
    MyName = NewName
End Sub

Public Property Get Bar() As Office.CommandBar
    Set Bar = MyBar
End Property


'=== construction and destruction end ===
'=== bar generator begin ===


Public Function CreateBar(oApp As Application, BarName As String) As Office.CommandBar
    RemoveExistingBar oApp, BarName
    Set CreateBar = oApp.CommandBars.Add(Name:=BarName, Temporary:=True)
End Function

Public Sub RemoveExistingBar(oApp As Application, BarName As String)
    On Error Resume Next
    oApp.CommandBars(BarName).Delete
End Sub

Public Sub DeleteBar(Bar As Object)
    On Error Resume Next
    Bar.Delete
End Sub

Public Sub ShowBar(Bar As Object, Optional Position As Long = msoBarTop, Optional Height As Long = 0)
    Bar.Visible = True
    Bar.Position = Position
    If Height > 0 Then Bar.Height = Bar.Height * Height
End Sub


'=== bar generator end ===
'=== handle addins begin ===


Public Function WithAddins(ParamArray Addins() As Variant) As Long
    WithAddins = AddAddins(MyBar, CVar(Addins))
End Function

Public Function AddAddins(Bar As Object, Addins As Variant) As Long
    Dim Addin As Variant
    Dim LastButtonIndex As Long
    
    For Each Addin In Addins
        LastButtonIndex = AddButtons(Bar, Addin.ButtonData, Addin.ButtonParent)
    Next
    
    AddAddins = LastButtonIndex
End Function


'=== handle addins end ===
'=== button generator begin ===


Public Function AddButtons(Bar As Object, Data As Variant, Parent As Variant) As Long
    Dim LastButtonIndex As Long
    Dim SingleData As Variant
    
    For Each SingleData In Data
        LastButtonIndex = Add(Bar, MakeAButtonData(SingleData, Parent))
    Next
    
    AddButtons = LastButtonIndex
End Function

Public Function Add(Bar As Object, Data As Variant) As Long
    Dim ButtonA As CommandBarControl
    
    Set ButtonA = Bar.Controls.Add(Type:=ButtonControlType(Data), Temporary:=True)
    With ButtonA
        Select Case ButtonControlType(Data)
        Case msoControlEdit                         '2      ' textbox
        Case msoControlDropdown, msoControlComboBox '3, 4   ' list and combo
            SetButtonItems ButtonA, Data
            SetButtonStyle ButtonA, Data
        Case msoControlPopup                        '10     ' popup
            SetButtonPopup ButtonA, Data
        Case msoControlButton                       '1      ' Button
            SetButtonStyle ButtonA, Data
            SetButtonState ButtonA, Data
        End Select
        SetButtonWidth ButtonA, Data
        SetButtonGroup ButtonA, Data
        .OnAction = ButtonAction(Data)
        .Caption = ButtonCaption(Data)
        .TooltipText = ButtonDescription(Data)
        .Tag = ButtonTag(Data)
        .Parameter = ButtonParameter(Data)
    End With
    
    Add = ButtonA.Index
    Set ButtonA = Nothing
End Function

Public Sub Remove(Bar As Object, Items As Variant)
    On Error Resume Next
    Dim Item As Variant
    
    If IsArray(Item) Then
        For Each Item In Items
            Remove Bar, Item
        Next
    Else
        Bar.Controls(Item).Delete
    End If
End Sub


'=== button generator end ===
'=== button data structure begin ===


' generator / selector

' Data(): Array of button data
' Parent(): Array of button parent information (bar and properties)
'           Parent(0) is reserved for addin key


Public Function MakeAButtonData(Data As Variant, Parent As Variant) As Variant
    MakeAButtonData = Array(NormalizeArray(Data), Parent)
End Function

Public Function DataAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    DataAButtonData = AButtonData(0)
End Function

Public Function ParentAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    ParentAButtonData = AButtonData(1)
End Function

Public Function KeyAButtonData(AButtonData As Variant) As String
    On Error Resume Next
    KeyAButtonData = ParentAButtonData(AButtonData)(0)
End Function

Public Function ItemAButtonData(AButtonData As Variant, ByVal Item As Long, _
            Optional FallBack As Variant = Empty) As Variant
    On Error Resume Next
    Dim out As Variant
    
    out = DataAButtonData(AButtonData)(Item)
    If IsEmpty(out) Then out = FallBack
    
    ItemAButtonData = out
End Function


'=== button data structure end ===
'=== button data struncture detail begin ===


Public Function ButtonCaption(Data As Variant) As String
    ButtonCaption = ItemAButtonData(Data, 0)
End Function

Public Function ButtonDescription(Data As Variant) As String
    ButtonDescription = ItemAButtonData(Data, 1)
End Function

Public Function ButtonTag(Data As Variant) As String
    ButtonTag = ItemAButtonData(Data, 2, ButtonCaption(Data))
End Function

Public Function ButtonParameter(Data As Variant) As String
    ButtonParameter = ItemAButtonData(Data, 3)
End Function

Public Function ButtonControlType(Data As Variant) As Long
    'MsoControlType
    On Error Resume Next
    ButtonControlType = Val(ItemAButtonData(Data, 4, msoControlButton))
End Function

Public Function ButtonStyle(Data As Variant) As Long
    'MsoButtonStyle
    On Error Resume Next
    ButtonStyle = Val(ItemAButtonData(Data, 5, msoButtonCaption))
End Function

Public Function ButtonWidth(Data As Variant) As Long
    ' we use 45 units here
    On Error Resume Next
    Const UnitWidth = 45
    ButtonWidth = Val(ItemAButtonData(Data, 6)) * UnitWidth
End Function

Public Function ButtonGroup(Data As Variant) As Boolean
    ' put group line on its left
    ButtonGroup = Not IsEmpty(ItemAButtonData(Data, 7))
End Function

Public Function ButtonAction(Data As Variant) As String
    On Error Resume Next
    ' Standard Method Name to be kicked with the button
    Const BarMain = "BarMain"
    Dim FullName As String
    
    If KeyAButtonData(Data) = "" Then
        FullName = BarMain
    Else
        FullName = KeyAButtonData(Data) & "." & BarMain
    End If
    
    ButtonAction = ItemAButtonData(Data, 8, FullName)
End Function

Public Function ButtonItems(Data As Variant) As Variant
    Dim pan As Variant
    Dim i As Long
    
    On Error GoTo DONE
    pan = Empty
    i = 9
    
    Do Until IsEmpty(ItemAButtonData(Data, i))
        pan = Array(ItemAButtonData(Data, i), pan)
        i = i + 1
    Loop
    
DONE:
    ButtonItems = pan
End Function


'=== button data struncture detail end ===
'=== button tools for data begin ===


Public Sub SetButtonWidth(ButtonA As CommandBarControl, Data As Variant)
    If ButtonWidth(Data) > 0 Then ButtonA.Width = ButtonWidth(Data)
End Sub

Public Sub SetButtonStyle(ButtonA As Object, Data As Variant)
    On Error Resume Next
    ' Each Button does not accept each style, but we won't check them.
    If ButtonStyle(Data) <> 0 Then ButtonA.Style = ButtonStyle(Data)
End Sub

Public Sub SetButtonGroup(ButtonA As CommandBarControl, Data As Variant)
    If ButtonGroup(Data) Then ButtonA.BeginGroup = True
End Sub

Public Sub SetButtonItems(ButtonA As Object, Data As Variant)
    Dim pan As Variant
    Dim HasItem As Boolean
    
    pan = ButtonItems(Data)
    HasItem = False
    
    Do Until IsEmpty(pan)
        ButtonA.AddItem pan(0), 1
        pan = pan(1)
        HasItem = True
    Loop
    If HasItem Then ButtonA.ListIndex = 1
End Sub

Public Sub SetButtonPopup(ButtonA As CommandBarControl, Data As Variant)
    Dim MyChild As Variant
    
    MyChild = StackToArray(ButtonItems(Data))
    If UBound(MyChild) >= 0 Then Add ButtonA, MyChild
End Sub

Public Sub SetButtonState(ButtonA As Object, Data As Variant)
    If Not IsEmpty(ButtonItems(Data)) Then ButtonA.State = msoButtonDown
End Sub


'=== button tools for data end ===
'=== button tools for control object begin ===


Public Sub ComboAddHistory(oAC As Object, Optional AtBottom As Boolean = False)
    If oAC.ListIndex > 0 Then Exit Sub
    
    If AtBottom Then
        oAC.AddItem oAC.Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem oAC.Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Sub ListAddHistory(oAC As Object, Text As String, Optional AtBottom As Boolean = False)
    If AtBottom Then
        oAC.AddItem Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Function ListFindIndex(oAC As Object, Text As String) As Long
    Dim i As Long
    For i = 1 To oAC.ListCount
        If oAC.List(i) = Text Then
            ListFindIndex = i
            Exit Function
        End If
    Next
    ListFindIndex = 0
End Function

Public Function ControlText(oAC As Object) As String
    ControlText = oAC.Text
End Function

Public Sub ButtonSwitchOn(oAC As Object)
    oAC.State = msoButtonDown
End Sub

Public Sub ButtonSwitchOff(oAC As Object)
    oAC.State = msoButtonUp
End Sub

Public Function ButtonSwitchToggle(oAC As Object) As Boolean
    ButtonSwitchToggle = (Not IsButtonStateOn(oAC))
    If ButtonSwitchToggle Then
        ButtonSwitchOn oAC
    Else
        ButtonSwitchOff oAC
    End If
End Function

Public Function IsButtonStateOn(oAC As Object) As Boolean
    IsButtonStateOn = (oAC.State = msoButtonDown)
End Function

Public Function ButtonFindByTag(oAC As Object, Tag As Variant) As CommandBarControl
    If oAC Is Nothing Then Exit Function
    If TypeName(oAC) = "CommandBar" Then
        Set ButtonFindByTag = oAC.FindControl(Tag:=Tag)
    Else
        Set ButtonFindByTag = oAC.Parent.FindControl(Tag:=Tag)
    End If
End Function


'=== button tools for control object end ===
'=== button tools for mybar begin ===


Public Function GetButton(TagOrIndex As Variant) As Office.CommandBarControl
    On Error Resume Next
    Select Case TypeName(TagOrIndex)
    Case "Long", "Integer", "Byte", "Double", "Single"
        Set GetButton = MyBar.Controls(TagOrIndex)
    Case Else
        Set GetButton = ButtonFindByTag(MyBar, TagOrIndex)
    End Select
End Function

Public Function GetControlText(TagOrIndex As Variant) As String
    Dim out As String
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Select Case oAC.Type
    Case msoControlEdit, msoControlDropdown, msoControlComboBox
        out = oAC.Text
    Case Else   ' msoControlButton, msoControlPopup
        out = oAC.Caption
    End Select
    
    Set oAC = Nothing
    GetControlText = out
End Function

Public Function SetControlText(TagOrIndex As Variant, ByVal Text As String) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    Dim Index As Long
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then
        out = False
    Else
        Select Case oAC.Type
        Case msoControlEdit
            oAC.Text = Text
        Case msoControlDropdown
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                ListAddHistory oAC, Text
            Else
                oAC.ListIndex = Index
            End If
        Case msoControlComboBox
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                oAC.Text = Text
                ComboAddHistory oAC
            Else
                oAC.ListIndex = Index
            End If
        Case Else
            oAC.Caption = Text
        End Select
        Set oAC = Nothing
        out = True
    End If
    
    SetControlText = out
End Function

Public Function GetControlState(TagOrIndex As Variant) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        ' return True when the button is pushed down
        out = IsButtonStateOn(oAC)
    End If
    
    Set oAC = Nothing
    GetControlState = out
End Function

Public Function SetControlState(TagOrIndex As Variant, ByVal State As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        If IsButtonStateOn(oAC) <> State Then
            If State Then
                ButtonSwitchOn oAC
            Else
                ButtonSwitchOff oAC
            End If
            ' return True when the status is strictly changed
            out = True
        End If
    End If
    
    Set oAC = Nothing
    SetControlState = out
End Function

Public Function GetControlVisible(TagOrIndex As Variant) As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    GetControlVisible = oAC.Visible
End Function

Public Function SetControlVisible(TagOrIndex As Variant, ByVal Visible As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Visible <> Visible Then
        oAC.Visible = Visible
        ' return True when the visible is strictly changed
        out = True
    End If
    
    SetControlVisible = out
End Function

Public Function IncControlWidth(TagOrIndex As Variant, ByVal Width As Long) As Long
    Dim out As Long
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    On Error Resume Next
    oAC.Width = oAC.Width + Width
    ' return the width accepted (tips: setting 0 to width makes it becomes default)
    out = oAC.Width
    
    IncControlWidth = out
End Function


'=== button tools for mybar end ===
'=== helper functions begin ===


Public Function NormalizeArray(x As Variant) As Variant
    On Error Resume Next
    Dim out() As Variant
    Dim i As Long
    Dim L1 As Long
    Dim L2 As Long
    Dim U1 As Long
    Dim U2 As Long
    
    L1 = 0
    L2 = 0
    U1 = -1
    U2 = -1
    
    L1 = LBound(x)
    L2 = LBound(x, 2)   ' error unless 2 dimensions
    U1 = UBound(x)
    U2 = UBound(x, 2)   ' error unless 2 dimensions
    
    If U1 < L1 Then
        NormalizeArray = Array()
        Exit Function
    End If
    
    If U2 = -1 Then
        ReDim out(0 To U1 - L1)
        For i = 0 To UBound(out)
            out(i) = x(i + L1)
        Next
    Else
        ReDim out(0 To U2 - L2)
        For i = 0 To UBound(out)
            out(i) = x(L1, i + L2)
            ' we pick up the 1st line only
        Next
    End If
    
    NormalizeArray = out
End Function

Public Function StackToArray(pan As Variant) As Variant
    Dim out() As Variant
    Dim x As Variant
    Dim i As Long
    Dim Counter As Long
    
    x = Empty
    Counter = 0
    Do Until IsEmpty(pan)
        x = Array(pan(0), x)
        pan = pan(1)
        Counter = Counter + 1
    Loop
    
    If Counter = 0 Then
        StackToArray = Array()
        Exit Function
    End If
    
    ReDim out(0 To Counter - 1)
    i = 0
    Do Until IsEmpty(x)
        out(i) = x(0)
        x = x(1)
        i = i + 1
    Loop
    
    StackToArray = out
End Function


'=== helper functions end ===
'}}}

'class
'  name;HelloNow
'{{{
Option Explicit

' sample addin for ToolBarV2

Public Helper As ToolBarV2


'=== button data begin ===

Public Property Get ButtonData() As Variant
    ButtonData = Array( _
        Array("時計", "クリップボードに日時をコピー", "now", Empty, 1, 2, Empty, 1), _
        Array("入力", "文書に日時を挿入する", "enter", Empty, 1, 2), _
        Array("...", "設定項目の表示・非表示を切り替える", "customize", Empty, 1, 2, Empty, 1, Empty, 1), _
        Array("全角", "全角英数を使う", "zenkaku", Empty, 1, 2), _
        Array("素", "加工せずそのままの日付型データを挿入する", "datatype", Empty, 1, 2), _
        Array("時計", "時刻書式を選ぶか、新しく入力する", "format", Empty, 4, Empty, 4, Empty, Empty, "gggee年m月d日 aaaa", "gee/m/d aaa", "c", "ddddd", "dddddd", "ttttt", "yyyy/m/d h:nn:ss", "d dd y ww ddd dddd aaa aaaa w", "m mm mmm mmmm oooo q", "g gg ggg e ee yy yyyy", "h hh n nn s ss :/", "AM/PM am/pm A/P a/p AMPM"), _
        Array("更新", "表示を更新する頻度", "update", Empty, 3, Empty, 1, Empty, Empty, "日", "時", "分", "秒") _
        )
End Property

Public Property Get ButtonParent() As Variant
    ButtonParent = Array("HelloNowMain")
End Property


'=== button data end ===
'=== default main procedures begin ===


' followings need to be public, because they are called from outside by the Helper

Public Sub Menu_now(oAC As Object)
    CopyNow
End Sub

Public Sub Menu_enter(oAC As Object)
    EnterNow
End Sub

Public Sub Menu_customize(oAC As Object)
    ToggleCustomize oAC
End Sub

Public Sub Menu_format(oAC As Object)
    FormatNow oAC
End Sub

Public Sub Menu_update(oAC As Object)
    SetTimer oAC
End Sub

Public Sub Menu_datatype(oAC As Object)
    Helper.OnButtonToggle
End Sub

Public Sub Menu_zenkaku(oAC As Object)
    Helper.OnButtonToggle
End Sub


'=== default main procedures end ===
'=== HelloNow implements begin ===


Friend Sub TimerTask()
    UpdateNow
    SetTimer Helper.GetButton("update")
End Sub

Friend Sub ShowCustomize(Optional ShowPanel As Boolean = True)
    Dim oAC As Office.CommandBarControl
    
    Set oAC = Helper.GetButton("customize")
    If Helper.IsButtonStateOn(oAC) = ShowPanel Then Exit Sub
    
    ToggleCustomize oAC
End Sub

Private Sub UpdateNow()
    Helper.SetControlText "now", NowFormatted
End Sub

Private Sub CopyNow()
    CopyToClipboard NowFormattedWise
End Sub

Private Sub EnterNow()
    HandleEnterNow NowFormattedWise
End Sub

Private Sub FormatNow(oAC As Office.CommandBarControl)
    Helper.ComboAddHistory oAC
    UpdateNow
End Sub

Private Sub ToggleCustomize(oAC As Office.CommandBarControl)
    Dim Visible As Boolean
    Dim TagName As Variant
    Dim Tags As Variant
    
    Tags = Array("zenkaku", "format", "update", "datatype")
    Visible = Helper.ButtonSwitchToggle(oAC)
    For Each TagName In Tags
        Helper.SetControlVisible TagName, Visible
    Next
End Sub

Private Sub SetTimer(oAC As Office.CommandBarControl)
    ResetTimer NextUpdateTime(oAC)
End Sub

Private Function NextUpdateTime(oAC As Office.CommandBarControl) As Variant
    Dim out As Date
    Dim outDelay As Date
    
    Select Case oAC.ListIndex
    Case 1  ' next day
        out = Date + 1
        outDelay = out + 1
    Case 2  ' next hour
        out = Date + TimeSerial(Hour(Now) + 1, 0, 0)
        outDelay = out + TimeValue("1:00:00")
    Case 3  ' next minute
        out = Date + TimeSerial(Hour(Now), Minute(Now) + 1, 0)
        outDelay = out + TimeValue("0:01:00")
    Case 4  ' next second
        out = Now + TimeValue("0:00:01")
        outDelay = out + TimeValue("0:00:01")
    End Select
    
    NextUpdateTime = Array(out, outDelay)
End Function

Private Function NowFormattedWise() As Variant
    Dim out As Variant
    
    If Helper.GetControlState("datatype") Then
        out = Now()
    ElseIf Helper.GetControlState("zenkaku") Then
        out = StrConv(NowFormatted, vbWide)
    Else
        out = NowFormatted
    End If
    
    NowFormattedWise = out
End Function

Private Function NowFormatted() As String
    NowFormatted = Format(Now(), Helper.GetControlText("format"))
End Function

Private Sub CopyToClipboard(Text As String)
    Dim Clip As MSForms.DataObject
    
    Set Clip = New MSForms.DataObject
    Clip.SetText Text
    Clip.PutInClipboard
    
    Set Clip = Nothing
End Sub


'=== HelloNow implements end ===
'=== event procedures begin ===


Private Sub Class_Initialize()
    Dim vMe As Variant
    Set vMe = Me
    Set Helper = New ToolBarV2
    Helper.SetName "HelloNow"
    Helper.NewBar vMe
    
    TimerTask
End Sub

Private Sub Class_Terminate()
    RemoveTimer
    
    Helper.DelBar
    Set Helper = Nothing
End Sub


'=== event procedures end ===
'}}}

'module
'  name;HelloNowMain
'{{{
Option Explicit

Private T1 As HelloNow
Private TimerSubmitted As Variant


Sub ClockInitialize(Optional Reload As Boolean = False)
    Set T1 = New HelloNow
    If Reload Then Exit Sub
    
    T1.ShowCustomize False
End Sub

Sub ClockTerminate()
    Set T1 = Nothing
End Sub


' this will called by pressing a button
Public Sub BarMain(Optional oWho As Object = Nothing)
    On Error GoTo OTL
    T1.Helper.BarMain T1
    Exit Sub
    
OTL:
    ClockInitialize True
End Sub


'=== low level i/o begin ===
' for Microsoft Excel


Public Sub HandleEnterNow(Data As Variant)
    On Error Resume Next
    
    Selection.Value = Data
    If Err.Number = 0 Then Exit Sub
    
    Err.Clear
    Selection.Text = Data
    If Err.Number = 0 Then Exit Sub
    
    Err.Clear
    Selection.Characters.Text = Data
End Sub


'=== low level i/o end ===
'=== timer begin ===


Public Sub Task()
    On Error Resume Next
    T1.TimerTask
End Sub

Public Sub ResetTimer(NewTime As Variant)
    Const TheProc = "Task"
    Dim TheTime As Date
    Dim TheDelay As Date

    TheTime = NewTime(0)
    TheDelay = NewTime(1)
    TimerSubmitted = Array(TheTime, TheProc, TheDelay)
    Application.OnTime TheTime, TheProc, TheDelay
End Sub

Public Function RemoveTimer() As Boolean
    If Not IsArray(TimerSubmitted) Then Exit Function
    On Error Resume Next
    Application.OnTime TimerSubmitted(0), TimerSubmitted(1), TimerSubmitted(2), False
    RemoveTimer = True
End Function


'=== timer end ===
'}}}

'code
'  name;ThisWorkbook
'{{{
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ClockTerminate
End Sub

Private Sub Workbook_Open()
    ClockInitialize
End Sub
'}}}


```

### Common Code ###

```
'class
'  name;HelloNow
'{{{
Option Explicit

' sample addin for ToolBarV2

Public Helper As ToolBarV2


'=== button data begin ===

Public Property Get ButtonData() As Variant
    ButtonData = Array( _
        Array("時計", "クリップボードに日時をコピー", "now", Empty, 1, 2, Empty, 1), _
        Array("入力", "文書に日時を挿入する", "enter", Empty, 1, 2), _
        Array("...", "設定項目の表示・非表示を切り替える", "customize", Empty, 1, 2, Empty, 1, Empty, 1), _
        Array("全角", "全角英数を使う", "zenkaku", Empty, 1, 2), _
        Array("素", "加工せずそのままの日付型データを挿入する", "datatype", Empty, 1, 2), _
        Array("時計", "時刻書式を選ぶか、新しく入力する", "format", Empty, 4, Empty, 4, Empty, Empty, "gggee年m月d日 aaaa", "gee/m/d aaa", "c", "ddddd", "dddddd", "ttttt", "yyyy/m/d h:nn:ss", "d dd y ww ddd dddd aaa aaaa w", "m mm mmm mmmm oooo q", "g gg ggg e ee yy yyyy", "h hh n nn s ss :/", "AM/PM am/pm A/P a/p AMPM"), _
        Array("更新", "表示を更新する頻度", "update", Empty, 3, Empty, 1, Empty, Empty, "日", "時", "分", "秒") _
        )
End Property

Public Property Get ButtonParent() As Variant
    ButtonParent = Array("HelloNowMain")
End Property


'=== button data end ===
'=== default main procedures begin ===


' followings need to be public, because they are called from outside by the Helper

Public Sub Menu_now(oAC As Object)
    CopyNow
End Sub

Public Sub Menu_enter(oAC As Object)
    EnterNow
End Sub

Public Sub Menu_customize(oAC As Object)
    ToggleCustomize oAC
End Sub

Public Sub Menu_format(oAC As Object)
    FormatNow oAC
End Sub

Public Sub Menu_update(oAC As Object)
    SetTimer oAC
End Sub

Public Sub Menu_datatype(oAC As Object)
    Helper.OnButtonToggle
End Sub

Public Sub Menu_zenkaku(oAC As Object)
    Helper.OnButtonToggle
End Sub


'=== default main procedures end ===
'=== HelloNow implements begin ===


Friend Sub TimerTask()
    UpdateNow
    SetTimer Helper.GetButton("update")
End Sub

Friend Sub ShowCustomize(Optional ShowPanel As Boolean = True)
    Dim oAC As Office.CommandBarControl
    
    Set oAC = Helper.GetButton("customize")
    If Helper.IsButtonStateOn(oAC) = ShowPanel Then Exit Sub
    
    ToggleCustomize oAC
End Sub

Private Sub UpdateNow()
    Helper.SetControlText "now", NowFormatted
End Sub

Private Sub CopyNow()
    CopyToClipboard NowFormattedWise
End Sub

Private Sub EnterNow()
    HandleEnterNow NowFormattedWise
End Sub

Private Sub FormatNow(oAC As Office.CommandBarControl)
    Helper.ComboAddHistory oAC
    UpdateNow
End Sub

Private Sub ToggleCustomize(oAC As Office.CommandBarControl)
    Dim Visible As Boolean
    Dim TagName As Variant
    Dim Tags As Variant
    
    Tags = Array("zenkaku", "format", "update", "datatype")
    Visible = Helper.ButtonSwitchToggle(oAC)
    For Each TagName In Tags
        Helper.SetControlVisible TagName, Visible
    Next
End Sub

Private Sub SetTimer(oAC As Office.CommandBarControl)
    ResetTimer NextUpdateTime(oAC)
End Sub

Private Function NextUpdateTime(oAC As Office.CommandBarControl) As Variant
    Dim out As Date
    Dim outDelay As Date
    
    Select Case oAC.ListIndex
    Case 1  ' next day
        out = Date + 1
        outDelay = out + 1
    Case 2  ' next hour
        out = Date + TimeSerial(Hour(Now) + 1, 0, 0)
        outDelay = out + TimeValue("1:00:00")
    Case 3  ' next minute
        out = Date + TimeSerial(Hour(Now), Minute(Now) + 1, 0)
        outDelay = out + TimeValue("0:01:00")
    Case 4  ' next second
        out = Now + TimeValue("0:00:01")
        outDelay = out + TimeValue("0:00:01")
    End Select
    
    NextUpdateTime = Array(out, outDelay)
End Function

Private Function NowFormattedWise() As Variant
    Dim out As Variant
    
    If Helper.GetControlState("datatype") Then
        out = Now()
    ElseIf Helper.GetControlState("zenkaku") Then
        out = StrConv(NowFormatted, vbWide)
    Else
        out = NowFormatted
    End If
    
    NowFormattedWise = out
End Function

Private Function NowFormatted() As String
    NowFormatted = Format(Now(), Helper.GetControlText("format"))
End Function

Private Sub CopyToClipboard(Text As String)
    Dim Clip As MSForms.DataObject
    
    Set Clip = New MSForms.DataObject
    Clip.SetText Text
    Clip.PutInClipboard
    
    Set Clip = Nothing
End Sub


'=== HelloNow implements end ===
'=== event procedures begin ===


Private Sub Class_Initialize()
    Dim vMe As Variant
    Set vMe = Me
    Set Helper = New ToolBarV2
    Helper.SetName "HelloNow"
    Helper.NewBar vMe
    
    TimerTask
End Sub

Private Sub Class_Terminate()
    RemoveTimer
    
    Helper.DelBar
    Set Helper = Nothing
End Sub


'=== event procedures end ===
'}}}


'class
'  name;ToolBarV2
'{{{
Option Explicit

' Generate an application toolbar

Private MyBar As Office.CommandBar
Private MyName As String
Private MyApp As Application


'=== main procedures helper begin ===


' this will called by pressing a button
Friend Sub BarMain(Optional oWho As Object = Nothing)
    Dim oAC As Object   ' this is the button itself pressed
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    ' switch to a main menu procedure
    Main oAC, SomebodyOrMe(oWho)
    Set oAC = Nothing
End Sub

' main menu procedure. if you delete this, a public Main in Standard Module will be called, maybe.
Private Sub Main(oAC As Object, Optional oWho As Object = Nothing)
    ' use a button tag to switch a procedure to be called as "Menu_xx"
    CallByName SomebodyOrMe(oWho), "Menu_" & oAC.Tag, VbMethod, oAC
End Sub

Public Sub Menu_about(oAC As Object)
    MsgBox TypeName(Me), vbOKOnly, "Sample of procedure called by the Main"
End Sub

Friend Sub OnButtonToggle()
    Dim oAC As Object   ' toggle this button
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    
    ButtonSwitchToggle oAC
    Set oAC = Nothing
End Sub

Private Function SomebodyOrMe(oWho As Object) As Object
    If oWho Is Nothing Then
        Set SomebodyOrMe = Me
    Else
        Set SomebodyOrMe = oWho
    End If
End Function


'=== main procedures helper end ===
'=== event procedures begin ===


Private Sub Class_Initialize()
    Set MyApp = Application
    MyName = CStr(Timer)    ' random name, maybe uniq
End Sub

Private Sub Class_Terminate()
    Set MyApp = Nothing
End Sub


'=== event procedures end ===
'=== construction and destruction begin ===


Public Sub NewBar(ParamArray Addins() As Variant)
    DelBar
    Set MyBar = CreateBar(MyApp, MyName)
    AddAddins MyBar, CVar(Addins)
    ShowBar MyBar
End Sub

Public Sub DelBar()
    DeleteBar MyBar
    Set MyBar = Nothing
End Sub

Public Sub SetApplication(oApp As Application)
    Set MyApp = oApp
End Sub

Public Sub SetName(NewName As String)
    MyName = NewName
End Sub

Public Property Get Bar() As Office.CommandBar
    Set Bar = MyBar
End Property


'=== construction and destruction end ===
'=== bar generator begin ===


Public Function CreateBar(oApp As Application, BarName As String) As Office.CommandBar
    RemoveExistingBar oApp, BarName
    Set CreateBar = oApp.CommandBars.Add(Name:=BarName, Temporary:=True)
End Function

Public Sub RemoveExistingBar(oApp As Application, BarName As String)
    On Error Resume Next
    oApp.CommandBars(BarName).Delete
End Sub

Public Sub DeleteBar(Bar As Object)
    On Error Resume Next
    Bar.Delete
End Sub

Public Sub ShowBar(Bar As Object, Optional Position As Long = msoBarTop, Optional Height As Long = 0)
    Bar.Visible = True
    Bar.Position = Position
    If Height > 0 Then Bar.Height = Bar.Height * Height
End Sub


'=== bar generator end ===
'=== handle addins begin ===


Public Function WithAddins(ParamArray Addins() As Variant) As Long
    WithAddins = AddAddins(MyBar, CVar(Addins))
End Function

Public Function AddAddins(Bar As Object, Addins As Variant) As Long
    Dim Addin As Variant
    Dim LastButtonIndex As Long
    
    For Each Addin In Addins
        LastButtonIndex = AddButtons(Bar, Addin.ButtonData, Addin.ButtonParent)
    Next
    
    AddAddins = LastButtonIndex
End Function


'=== handle addins end ===
'=== button generator begin ===


Public Function AddButtons(Bar As Object, Data As Variant, Parent As Variant) As Long
    Dim LastButtonIndex As Long
    Dim SingleData As Variant
    
    For Each SingleData In Data
        LastButtonIndex = Add(Bar, MakeAButtonData(SingleData, Parent))
    Next
    
    AddButtons = LastButtonIndex
End Function

Public Function Add(Bar As Object, Data As Variant) As Long
    Dim ButtonA As CommandBarControl
    
    Set ButtonA = Bar.Controls.Add(Type:=ButtonControlType(Data), Temporary:=True)
    With ButtonA
        Select Case ButtonControlType(Data)
        Case msoControlEdit                         '2      ' textbox
        Case msoControlDropdown, msoControlComboBox '3, 4   ' list and combo
            SetButtonItems ButtonA, Data
            SetButtonStyle ButtonA, Data
        Case msoControlPopup                        '10     ' popup
            SetButtonPopup ButtonA, Data
        Case msoControlButton                       '1      ' Button
            SetButtonStyle ButtonA, Data
            SetButtonState ButtonA, Data
        End Select
        SetButtonWidth ButtonA, Data
        SetButtonGroup ButtonA, Data
        .OnAction = ButtonAction(Data)
        .Caption = ButtonCaption(Data)
        .TooltipText = ButtonDescription(Data)
        .Tag = ButtonTag(Data)
        .Parameter = ButtonParameter(Data)
    End With
    
    Add = ButtonA.Index
    Set ButtonA = Nothing
End Function

Public Sub Remove(Bar As Object, Items As Variant)
    On Error Resume Next
    Dim Item As Variant
    
    If IsArray(Item) Then
        For Each Item In Items
            Remove Bar, Item
        Next
    Else
        Bar.Controls(Item).Delete
    End If
End Sub


'=== button generator end ===
'=== button data structure begin ===


' generator / selector

' Data(): Array of button data
' Parent(): Array of button parent information (bar and properties)
'           Parent(0) is reserved for addin key


Public Function MakeAButtonData(Data As Variant, Parent As Variant) As Variant
    MakeAButtonData = Array(NormalizeArray(Data), Parent)
End Function

Public Function DataAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    DataAButtonData = AButtonData(0)
End Function

Public Function ParentAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    ParentAButtonData = AButtonData(1)
End Function

Public Function KeyAButtonData(AButtonData As Variant) As String
    On Error Resume Next
    KeyAButtonData = ParentAButtonData(AButtonData)(0)
End Function

Public Function ItemAButtonData(AButtonData As Variant, ByVal Item As Long, _
            Optional FallBack As Variant = Empty) As Variant
    On Error Resume Next
    Dim out As Variant
    
    out = DataAButtonData(AButtonData)(Item)
    If IsEmpty(out) Then out = FallBack
    
    ItemAButtonData = out
End Function


'=== button data structure end ===
'=== button data struncture detail begin ===


Public Function ButtonCaption(Data As Variant) As String
    ButtonCaption = ItemAButtonData(Data, 0)
End Function

Public Function ButtonDescription(Data As Variant) As String
    ButtonDescription = ItemAButtonData(Data, 1)
End Function

Public Function ButtonTag(Data As Variant) As String
    ButtonTag = ItemAButtonData(Data, 2, ButtonCaption(Data))
End Function

Public Function ButtonParameter(Data As Variant) As String
    ButtonParameter = ItemAButtonData(Data, 3)
End Function

Public Function ButtonControlType(Data As Variant) As Long
    'MsoControlType
    On Error Resume Next
    ButtonControlType = Val(ItemAButtonData(Data, 4, msoControlButton))
End Function

Public Function ButtonStyle(Data As Variant) As Long
    'MsoButtonStyle
    On Error Resume Next
    ButtonStyle = Val(ItemAButtonData(Data, 5, msoButtonCaption))
End Function

Public Function ButtonWidth(Data As Variant) As Long
    ' we use 45 units here
    On Error Resume Next
    Const UnitWidth = 45
    ButtonWidth = Val(ItemAButtonData(Data, 6)) * UnitWidth
End Function

Public Function ButtonGroup(Data As Variant) As Boolean
    ' put group line on its left
    ButtonGroup = Not IsEmpty(ItemAButtonData(Data, 7))
End Function

Public Function ButtonAction(Data As Variant) As String
    On Error Resume Next
    ' Standard Method Name to be kicked with the button
    Const BarMain = "BarMain"
    Dim FullName As String
    
    If KeyAButtonData(Data) = "" Then
        FullName = BarMain
    Else
        FullName = KeyAButtonData(Data) & "." & BarMain
    End If
    
    ButtonAction = ItemAButtonData(Data, 8, FullName)
End Function

Public Function ButtonItems(Data As Variant) As Variant
    Dim pan As Variant
    Dim i As Long
    
    On Error GoTo DONE
    pan = Empty
    i = 9
    
    Do Until IsEmpty(ItemAButtonData(Data, i))
        pan = Array(ItemAButtonData(Data, i), pan)
        i = i + 1
    Loop
    
DONE:
    ButtonItems = pan
End Function


'=== button data struncture detail end ===
'=== button tools for data begin ===


Public Sub SetButtonWidth(ButtonA As CommandBarControl, Data As Variant)
    If ButtonWidth(Data) > 0 Then ButtonA.Width = ButtonWidth(Data)
End Sub

Public Sub SetButtonStyle(ButtonA As Object, Data As Variant)
    On Error Resume Next
    ' Each Button does not accept each style, but we won't check them.
    If ButtonStyle(Data) <> 0 Then ButtonA.Style = ButtonStyle(Data)
End Sub

Public Sub SetButtonGroup(ButtonA As CommandBarControl, Data As Variant)
    If ButtonGroup(Data) Then ButtonA.BeginGroup = True
End Sub

Public Sub SetButtonItems(ButtonA As Object, Data As Variant)
    Dim pan As Variant
    Dim HasItem As Boolean
    
    pan = ButtonItems(Data)
    HasItem = False
    
    Do Until IsEmpty(pan)
        ButtonA.AddItem pan(0), 1
        pan = pan(1)
        HasItem = True
    Loop
    If HasItem Then ButtonA.ListIndex = 1
End Sub

Public Sub SetButtonPopup(ButtonA As CommandBarControl, Data As Variant)
    Dim MyChild As Variant
    
    MyChild = StackToArray(ButtonItems(Data))
    If UBound(MyChild) >= 0 Then Add ButtonA, MyChild
End Sub

Public Sub SetButtonState(ButtonA As Object, Data As Variant)
    If Not IsEmpty(ButtonItems(Data)) Then ButtonA.State = msoButtonDown
End Sub


'=== button tools for data end ===
'=== button tools for control object begin ===


Public Sub ComboAddHistory(oAC As Object, Optional AtBottom As Boolean = False)
    If oAC.ListIndex > 0 Then Exit Sub
    
    If AtBottom Then
        oAC.AddItem oAC.Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem oAC.Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Sub ListAddHistory(oAC As Object, Text As String, Optional AtBottom As Boolean = False)
    If AtBottom Then
        oAC.AddItem Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Function ListFindIndex(oAC As Object, Text As String) As Long
    Dim i As Long
    For i = 1 To oAC.ListCount
        If oAC.List(i) = Text Then
            ListFindIndex = i
            Exit Function
        End If
    Next
    ListFindIndex = 0
End Function

Public Function ControlText(oAC As Object) As String
    ControlText = oAC.Text
End Function

Public Sub ButtonSwitchOn(oAC As Object)
    oAC.State = msoButtonDown
End Sub

Public Sub ButtonSwitchOff(oAC As Object)
    oAC.State = msoButtonUp
End Sub

Public Function ButtonSwitchToggle(oAC As Object) As Boolean
    ButtonSwitchToggle = (Not IsButtonStateOn(oAC))
    If ButtonSwitchToggle Then
        ButtonSwitchOn oAC
    Else
        ButtonSwitchOff oAC
    End If
End Function

Public Function IsButtonStateOn(oAC As Object) As Boolean
    IsButtonStateOn = (oAC.State = msoButtonDown)
End Function

Public Function ButtonFindByTag(oAC As Object, Tag As Variant) As CommandBarControl
    If oAC Is Nothing Then Exit Function
    If TypeName(oAC) = "CommandBar" Then
        Set ButtonFindByTag = oAC.FindControl(Tag:=Tag)
    Else
        Set ButtonFindByTag = oAC.Parent.FindControl(Tag:=Tag)
    End If
End Function


'=== button tools for control object end ===
'=== button tools for mybar begin ===


Public Function GetButton(TagOrIndex As Variant) As Office.CommandBarControl
    On Error Resume Next
    Select Case TypeName(TagOrIndex)
    Case "Long", "Integer", "Byte", "Double", "Single"
        Set GetButton = MyBar.Controls(TagOrIndex)
    Case Else
        Set GetButton = ButtonFindByTag(MyBar, TagOrIndex)
    End Select
End Function

Public Function GetControlText(TagOrIndex As Variant) As String
    Dim out As String
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Select Case oAC.Type
    Case msoControlEdit, msoControlDropdown, msoControlComboBox
        out = oAC.Text
    Case Else   ' msoControlButton, msoControlPopup
        out = oAC.Caption
    End Select
    
    Set oAC = Nothing
    GetControlText = out
End Function

Public Function SetControlText(TagOrIndex As Variant, ByVal Text As String) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    Dim Index As Long
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then
        out = False
    Else
        Select Case oAC.Type
        Case msoControlEdit
            oAC.Text = Text
        Case msoControlDropdown
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                ListAddHistory oAC, Text
            Else
                oAC.ListIndex = Index
            End If
        Case msoControlComboBox
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                oAC.Text = Text
                ComboAddHistory oAC
            Else
                oAC.ListIndex = Index
            End If
        Case Else
            oAC.Caption = Text
        End Select
        Set oAC = Nothing
        out = True
    End If
    
    SetControlText = out
End Function

Public Function GetControlState(TagOrIndex As Variant) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        ' return True when the button is pushed down
        out = IsButtonStateOn(oAC)
    End If
    
    Set oAC = Nothing
    GetControlState = out
End Function

Public Function SetControlState(TagOrIndex As Variant, ByVal State As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        If IsButtonStateOn(oAC) <> State Then
            If State Then
                ButtonSwitchOn oAC
            Else
                ButtonSwitchOff oAC
            End If
            ' return True when the status is strictly changed
            out = True
        End If
    End If
    
    Set oAC = Nothing
    SetControlState = out
End Function

Public Function GetControlVisible(TagOrIndex As Variant) As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    GetControlVisible = oAC.Visible
End Function

Public Function SetControlVisible(TagOrIndex As Variant, ByVal Visible As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Visible <> Visible Then
        oAC.Visible = Visible
        ' return True when the visible is strictly changed
        out = True
    End If
    
    SetControlVisible = out
End Function

Public Function IncControlWidth(TagOrIndex As Variant, ByVal Width As Long) As Long
    Dim out As Long
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    On Error Resume Next
    oAC.Width = oAC.Width + Width
    ' return the width accepted (tips: setting 0 to width makes it becomes default)
    out = oAC.Width
    
    IncControlWidth = out
End Function


'=== button tools for mybar end ===
'=== helper functions begin ===


Public Function NormalizeArray(x As Variant) As Variant
    On Error Resume Next
    Dim out() As Variant
    Dim i As Long
    Dim L1 As Long
    Dim L2 As Long
    Dim U1 As Long
    Dim U2 As Long
    
    L1 = 0
    L2 = 0
    U1 = -1
    U2 = -1
    
    L1 = LBound(x)
    L2 = LBound(x, 2)   ' error unless 2 dimensions
    U1 = UBound(x)
    U2 = UBound(x, 2)   ' error unless 2 dimensions
    
    If U1 < L1 Then
        NormalizeArray = Array()
        Exit Function
    End If
    
    If U2 = -1 Then
        ReDim out(0 To U1 - L1)
        For i = 0 To UBound(out)
            out(i) = x(i + L1)
        Next
    Else
        ReDim out(0 To U2 - L2)
        For i = 0 To UBound(out)
            out(i) = x(L1, i + L2)
            ' we pick up the 1st line only
        Next
    End If
    
    NormalizeArray = out
End Function

Public Function StackToArray(pan As Variant) As Variant
    Dim out() As Variant
    Dim x As Variant
    Dim i As Long
    Dim Counter As Long
    
    x = Empty
    Counter = 0
    Do Until IsEmpty(pan)
        x = Array(pan(0), x)
        pan = pan(1)
        Counter = Counter + 1
    Loop
    
    If Counter = 0 Then
        StackToArray = Array()
        Exit Function
    End If
    
    ReDim out(0 To Counter - 1)
    i = 0
    Do Until IsEmpty(x)
        out(i) = x(0)
        x = x(1)
        i = i + 1
    Loop
    
    StackToArray = out
End Function


'=== helper functions end ===
'}}}


```

### Code for Excel ###

```
'module
'  name;HelloNowMain
'{{{
Option Explicit

Private T1 As HelloNow
Private TimerSubmitted As Variant


Sub ClockInitialize(Optional Reload As Boolean = False)
    Set T1 = New HelloNow
    If Reload Then Exit Sub
    
    T1.ShowCustomize False
End Sub

Sub ClockTerminate()
    Set T1 = Nothing
End Sub


' this will called by pressing a button
Public Sub BarMain(Optional oWho As Object = Nothing)
    On Error GoTo OTL
    T1.Helper.BarMain T1
    Exit Sub
    
OTL:
    ClockInitialize True
End Sub


'=== low level i/o begin ===
' for Microsoft Excel


Public Sub HandleEnterNow(Data As Variant)
    On Error Resume Next
    
    Selection.Value = Data
    If Err.Number = 0 Then Exit Sub
    
    Err.Clear
    Selection.Text = Data
    If Err.Number = 0 Then Exit Sub
    
    Err.Clear
    Selection.Characters.Text = Data
End Sub


'=== low level i/o end ===
'=== timer begin ===


Public Sub Task()
    On Error Resume Next
    T1.TimerTask
End Sub

Public Sub ResetTimer(NewTime As Variant)
    Const TheProc = "Task"
    Dim TheTime As Date
    Dim TheDelay As Date

    TheTime = NewTime(0)
    TheDelay = NewTime(1)
    TimerSubmitted = Array(TheTime, TheProc, TheDelay)
    Application.OnTime TheTime, TheProc, TheDelay
End Sub

Public Function RemoveTimer() As Boolean
    If Not IsArray(TimerSubmitted) Then Exit Function
    On Error Resume Next
    Application.OnTime TimerSubmitted(0), TimerSubmitted(1), TimerSubmitted(2), False
    RemoveTimer = True
End Function


'=== timer end ===
'}}}


'code
'  name;ThisWorkbook
'{{{
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ClockTerminate
End Sub

Private Sub Workbook_Open()
    ClockInitialize
End Sub
'}}}


```

### Code for Word ###

```
'module
'  name;HelloNowMain
'{{{
Option Explicit

Private T1 As HelloNow


Sub ClockInitialize(Optional Reload As Boolean = False)
    Set T1 = New HelloNow
    If Reload Then Exit Sub
    
    T1.ShowCustomize False
End Sub

Sub ClockTerminate()
    Set T1 = Nothing
End Sub


' this will called by pressing a button
Public Sub BarMain(Optional oWho As Object = Nothing)
    On Error GoTo OTL
    T1.Helper.BarMain T1
    Exit Sub
    
OTL:
    ClockInitialize True
End Sub


'=== low level i/o begin ===
' for Microsoft Word


Public Sub HandleEnterNow(Data As Variant)
    Selection.TypeText Data
End Sub


'=== low level i/o end ===
'=== timer begin ===


Public Sub Task()
    On Error Resume Next
    T1.TimerTask
End Sub

Public Sub ResetTimer(NewTime As Variant)
    Const TheProc = "ProjectHelloNow.HelloNowMain.Task"
    Dim TheTime As Date

    TheTime = NewTime(0)
    Application.OnTime TheTime, TheProc
End Sub

' this will remove the timer, because Word has only one timer at once.
Public Sub RemoveTimer()
    Application.OnTime Now, ""
End Sub


'=== timer end ===
'}}}


'code
'  name;ThisDocument
'{{{
Option Explicit

Private Sub Document_Close()
    ClockTerminate
End Sub

Private Sub Document_Open()
    ClockInitialize
End Sub
'}}}


```

### Code for Access ###

```
'module
'  name;HelloNowMain
'{{{
Option Compare Database
Option Explicit

Private T1 As HelloNow
'Private TimerSubmitted As Variant


Public Function ClockInitialize(Optional Reload As Boolean = False)
    Set T1 = New HelloNow
    If Reload Then Exit Function
    
    T1.ShowCustomize False
End Function

Public Function ClockTerminate()
    Set T1 = Nothing
End Function


' this will called by pressing a button
Public Function BarMain(Optional oWho As Object = Nothing)
    On Error GoTo OTL
    T1.Helper.BarMain T1
    Exit Function
    
OTL:
    ClockInitialize True
End Function


'=== low level i/o begin ===
' for Microsoft Access


Public Sub HandleEnterNow(Data As Variant)
    SendKeys Data
End Sub


'=== low level i/o end ===
'=== timer begin ===


Public Sub Task()
    On Error Resume Next
    T1.TimerTask
End Sub

Public Sub ResetTimer(NewTime As Variant)
    Dim IntervalMilSec As Long
    
    IntervalMilSec = Int((NewTime(1) - NewTime(0)) * 24 * 60 * 60 * 1000)
    If IntervalMilSec > 60 * 60 * 1000& Then
        IntervalMilSec = IntervalMilSec / 24
    ElseIf IntervalMilSec > 10 * 1000& Then
        IntervalMilSec = IntervalMilSec / 10
    End If
    
    Form_HelloNowControler.SetTimerInterval IntervalMilSec
End Sub

Public Sub RemoveTimer()
    Form_HelloNowControler.SetTimerInterval 0
End Sub


'=== timer end ===
'}}}


'code
'  name;Form_HelloNowControler
'{{{
Option Compare Database
Option Explicit

Private Sub Form_Close()
    ClockTerminate
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Minimize
    ClockInitialize
End Sub

Private Sub Form_Timer()
    Task
End Sub

Public Function SetTimerInterval(NewInterval As Long)
    Me.TimerInterval = NewInterval
End Function
'}}}


```


# Snapshots #

  * Excel 2000, Word 2000, Access 2000 toolbar

> > ![http://2.bp.blogspot.com/_EUW0nrj9XlM/TT4oX5fCNmI/AAAAAAAAAB0/ogyWE-J9iIY/s1600/shot5.png](http://2.bp.blogspot.com/_EUW0nrj9XlM/TT4oX5fCNmI/AAAAAAAAAB0/ogyWE-J9iIY/s1600/shot5.png)
> > ![http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oYSAvrOI/AAAAAAAAAB4/hcrURVtBLyM/s1600/shot6.png](http://1.bp.blogspot.com/_EUW0nrj9XlM/TT4oYSAvrOI/AAAAAAAAAB4/hcrURVtBLyM/s1600/shot6.png)

  * Access 2000 From
> > ![http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oWhYQHwI/AAAAAAAAABs/OTeZX_oLQmg/s1600/shot3.png](http://4.bp.blogspot.com/_EUW0nrj9XlM/TT4oWhYQHwI/AAAAAAAAABs/OTeZX_oLQmg/s1600/shot3.png)

  * Access 2000 Macro
> > ![http://3.bp.blogspot.com/_EUW0nrj9XlM/TT4oXfFFysI/AAAAAAAAABw/tOBQ9dhKa1g/s1600/shot4.png](http://3.bp.blogspot.com/_EUW0nrj9XlM/TT4oXfFFysI/AAAAAAAAABw/tOBQ9dhKa1g/s1600/shot4.png)

  * Excel 2007 ribbon
> > ![http://3.bp.blogspot.com/_EUW0nrj9XlM/TT4oWLjZYWI/AAAAAAAAABo/nErWKPgcCF0/s1600/shot1.png](http://3.bp.blogspot.com/_EUW0nrj9XlM/TT4oWLjZYWI/AAAAAAAAABo/nErWKPgcCF0/s1600/shot1.png)
