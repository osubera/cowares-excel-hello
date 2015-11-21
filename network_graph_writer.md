

# Introduction #

  * create a network graph by excel

## 概要 ##
  * エクセルでネットワークグラフを手軽に描く

# Details #

## 説明 ##
  * ブックを開くと、操作用のボタンがツールバー（リボン）に追加される。
    * ❏、➔、↻反転、node TSV、arc TSV、SEM
  * ツールバーの機能
    * `❏` ボタンを押すとアクティブシートに囲み図形のノードを追加する。
    * `➔` ボタンを押すとアクティブシートに矢印の弧を追加する。
      * あらかじめノードを選べば、選択したノードを結合する。
      * ノードが選択されていないときは、最後に追加した２つのノードを結ぶ。
    * `↻反転` ボタンを押すと選択した弧の向きを反転する。
    * `node TSV` ボタンを押すとノード情報をTSV (タブ区切り) 出力する。
    * `arc TSV` ボタンを押すと弧の情報をTSV (タブ区切り) 出力する。
    * `SEM` ボタンを押すとRのsemパッケージ用に出力する。

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=network_graph_writer)

# How to use #


## 使い方 ##

  * 用途
    1. ネットワークグラフをエクセルブックに素早く書く。
    1. グラフの状態を再利用可能なテキストとしてエクスポートする。

  1. 一つの手順
    1. 新規ブックを開く。
    1. `❏` ボタンを押し、ノードを追加する。このとき、ノード名をダイアログで指定する。
    1. `❏` ボタンを押し、２つ目のノードを追加する。
    1. `➔` ボタンを押し、弧を追加する。２つのノードが矢印で結合される。
    1. ２つ目のノードをマウスで少し右に動かす。
    1. `❏` ボタンを押し、３つ目のノードを追加する。
    1. `➔` ボタンを押し、弧を追加する。２つのノードが矢印で結合される。
    1. 矢印の先をマウスで動かし、結合先を最初のノードにする。
    1. `SEM` ボタンを押し、メモ帳に出力される情報を確認する。
  1. 図形について
    1. ノードの名称は、図形の持つ名称 (アドレス欄で表示・変更) を使用する。
    1. ノード図形内のテキストは、コメント情報なので補助的な情報を追加しても動作は変わらない。
    1. エクセルの標準的な四角形や矢印を使用しているため、ツールバーを使わず作成した図形も使用できる。


# Snapshots #

**エクセル2000 ではツールバーを表示する。
> ![http://1.bp.blogspot.com/-8UP1E0fCsRs/UAIK0Vyy3oI/AAAAAAAAAKs/7x8-gjnmyb4/s1600/shot2.png](http://1.bp.blogspot.com/-8UP1E0fCsRs/UAIK0Vyy3oI/AAAAAAAAAKs/7x8-gjnmyb4/s1600/shot2.png)**

**エクセル2007 ではリボンのアドインタブにボタンを表示する。
> ![http://2.bp.blogspot.com/-OUhtRdSf6Eg/UAAdwMoCzAI/AAAAAAAAAKc/yo4eirFnUKU/s1600/shot1.png](http://2.bp.blogspot.com/-OUhtRdSf6Eg/UAAdwMoCzAI/AAAAAAAAAKc/yo4eirFnUKU/s1600/shot1.png)**



# Code #

```
'ssf-begin
';

'workbook
'   name;network_graph_writer.xlsm/F3NetworkGraphWriter

'book-identity

'require
'       ;{420B2830-E718-11CF-893D-00A0C9054228} 1 0 Microsoft Scripting Runtime
'       ;{3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5 5 Microsoft VBScript Regular Expressions 5.5

'cells-name
'       ;=network_graph_writer!R15C1
'       ;network_graph_writer!_ButtonCaption
'       ;=network_graph_writer!R3C2
'       ;network_graph_writer!_Comment
'       ;=network_graph_writer!R6C2
'       ;network_graph_writer!_Contributor
'       ;=network_graph_writer!R4C2
'       ;network_graph_writer!_Copyright
'       ;=network_graph_writer!R5C2
'       ;network_graph_writer!_License
'       ;=network_graph_writer!R2C2
'       ;network_graph_writer!_LocalComment
'       ;=network_graph_writer!R1C2
'       ;network_graph_writer!_PublicName
'       ;=network_graph_writer!R7C2
'       ;network_graph_writer!_Since
'       ;=network_graph_writer!R10C2
'       ;network_graph_writer!_Tag
'       ;=network_graph_writer!R9C2
'       ;network_graph_writer!_Timestamp
'       ;=network_graph_writer!R8C2
'       ;network_graph_writer!_Url

'worksheet
'   name;network_graph_writer/BaumMain

'cells-formula
'  address;A1:B10
'         ;名称
'         ;network_graph_writer
'         ;コメント
'         ;ネットワークグラフを手軽に描く
'         ;comment
'         ;simple network graph writer
'         ;著作権
'         ;="Copyright (C) " &R[3]C & "-" & YEAR(R[5]C) & " " & R[2]C
'         ;ライセンス
'         ;自律, 自由, 公正, http://cowares.nobody.jp
'         ;作者
'         ;Tomizono - mocvba.com
'         ;初版
'         ;2012
'         ;配布元
'         ;http://cowares.blogspot.com/search/label/baum
'         ;更新
'         ;41105.350462963
'         ;keyword
'         ;vba,excel,network,graph
'  address;A13:J13
'         ;ボタンの表示
'         ;ボタンの機能
'         ;Tag
'         ;Parameter
'         ;ControlType
'         ;Style
'         ;Width
'         ;Group
'         ;Action
'         ;Initialize ..
'  address;A15:H20
'         ;❏
'         ;ノードを追加する
'         ;node
'     skip;1
'         ;1
'         ;2
'     skip;1
'         ;1
'         ;➔
'         ;弧を追加する
'         ;arc
'     skip;1
'         ;1
'         ;2
'     skip;2
'         ;↻反転
'         ;弧の向きを反転する
'         ;flip
'     skip;1
'         ;1
'         ;2
'     skip;2
'         ;node TSV
'         ;ノード情報をTSV出力する
'         ;nodetsv
'     skip;1
'         ;1
'         ;2
'     skip;1
'         ;1
'         ;arc TSV
'         ;弧の情報をTSV出力する
'         ;arctsv
'     skip;1
'         ;1
'         ;2
'     skip;2
'         ;SEM
'         ;Rのsemパッケージ用出力
'         ;sem
'     skip;1
'         ;1
'         ;2

'cells-numberformat
'  address;B9
'         ;m/d/yyyy h:mm

'cells-width
'   unit;pt
'  address;B1
'         ;96.75

'cells-v-align
'  address;A1:J20
'   repeat;200
'         ;center

'code
'   name;BaumMain
'{{{
Option Explicit

' BaumMain addin for ToolBarV2

' using a excel worksheet as a property holder

' we do not support popup on excel sheet at this moment
' no ideas how to describe it wisely on 2 dimensional sheet

Private Helper As ToolBarV2
'Public Helper As ToolBarV2
' we cannot have a public variable in Excel Worksheet

Friend Function VBProjectName() As String
    ' VBProject.Name can't be accessed for the default settings.
    VBProjectName = "F3NetworkGraphWriter"
End Function

Friend Function GetHelper() As ToolBarV2
    Set GetHelper = Helper
End Function

'=== default main procedures begin ===


' this will called by pressing a button
Public Sub BarMain(Optional oWho As Object = Nothing)
    If Helper Is Nothing Then
        BarInitialize
        MsgBox "ツールバーを修復しました。もう一度操作してください。", vbExclamation, BaumMain.Name
    Else
        Helper.BarMain Me
    End If
End Sub

Public Sub OnButtonToggle()
    Helper.OnButtonToggle
End Sub

' followings need to be public, because they are called from outside by the Helper
' we also can copy the Helper.BarMain code here, and let the followings be private.

Public Sub Menu_node(oAC As Object)
    NetworkGraphWriter.DoNode
End Sub

Public Sub Menu_arc(oAC As Object)
    NetworkGraphWriter.DoArc
End Sub

Public Sub Menu_flip(oAC As Object)
    NetworkGraphWriter.DoFlip
End Sub

Public Sub Menu_nodetsv(oAC As Object)
    NetworkGraphWriter.DoNodeTsv
End Sub

Public Sub Menu_arctsv(oAC As Object)
    NetworkGraphWriter.DoArcTsv
End Sub

Public Sub Menu_sem(oAC As Object)
    NetworkGraphWriter.DoSem
End Sub


'=== default main procedures end ===
'=== button data begin ===

Public Property Get ButtonData() As Variant
    ButtonData = ConvertRangeToArray(Application.Intersect(GetButtonRow, GetButtonCol))
End Property

Public Property Get ButtonParent() As Variant
    ButtonParent = Array(VBProjectName & "." & Me.CodeName)
End Property

' above simple property codes are supported by the following range helpers

Private Function GetButtonRow(Optional Address As String = "_ButtonCaption") As Range
    Dim Out As Range
    Dim StartAt As Range
    
    Set StartAt = Me.Range(Address)
    If IsEmpty(StartAt.Offset(1, 0).Value) Then
        Set Out = StartAt
    Else
        Set Out = Me.Range(StartAt, StartAt.End(xlDown))
    End If
    
    Set GetButtonRow = Out.EntireRow
End Function

Private Function GetButtonCol(Optional Address As String = "_ButtonCaption") As Range
    Dim StartAt As Range
    Set StartAt = Me.Range(Address)
    Set GetButtonCol = Me.Range(StartAt, StartAt.SpecialCells(xlCellTypeLastCell)).EntireColumn
End Function

Private Function ConvertRangeToArray(Ra As Range) As Variant
    Dim Out() As Variant
    Dim i As Long
    
    ReDim Out(0 To Ra.Rows.Count - 1)
    For i = 0 To UBound(Out)
        Out(i) = Ra.Rows(i + 1).Value
    Next
    
    ConvertRangeToArray = Out
End Function


'=== button data end ===
'=== constructor / destructor begin ===


Private Function BarName() As String
    BarName = Me.Name & Me.Range("_PublicName").Text & Me.Range("_Timestamp").Text
End Function

Public Sub BarInitialize()
    Dim vMe As Variant
    Set vMe = Me
    Set Helper = New ToolBarV2
    Helper.SetName BarName
    Helper.NewBar vMe
End Sub

Public Sub BarTerminate()
    On Error Resume Next
    Helper.DelBar
    Set Helper = Nothing
End Sub


'=== constructor / destructor end ===

'}}}

'class
'   name;ToolBarV2
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
    Dim Out As Variant
    
    Out = DataAButtonData(AButtonData)(Item)
    If IsEmpty(Out) Then Out = FallBack
    
    ItemAButtonData = Out
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
    Dim Out As String
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Select Case oAC.Type
    Case msoControlEdit, msoControlDropdown, msoControlComboBox
        Out = oAC.Text
    Case Else   ' msoControlButton, msoControlPopup
        Out = oAC.Caption
    End Select
    
    Set oAC = Nothing
    GetControlText = Out
End Function

Public Function SetControlText(TagOrIndex As Variant, ByVal Text As String) As Boolean
    Dim Out As Boolean
    Dim oAC As Office.CommandBarControl
    Dim Index As Long
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then
        Out = False
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
        Out = True
    End If
    
    SetControlText = Out
End Function

Public Function GetControlState(TagOrIndex As Variant) As Boolean
    Dim Out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Out = False
    If oAC.Type = msoControlButton Then
        ' return True when the button is pushed down
        Out = IsButtonStateOn(oAC)
    End If
    
    Set oAC = Nothing
    GetControlState = Out
End Function

Public Function SetControlState(TagOrIndex As Variant, ByVal State As Boolean) As Boolean
    Dim Out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Out = False
    If oAC.Type = msoControlButton Then
        If IsButtonStateOn(oAC) <> State Then
            If State Then
                ButtonSwitchOn oAC
            Else
                ButtonSwitchOff oAC
            End If
            ' return True when the status is strictly changed
            Out = True
        End If
    End If
    
    Set oAC = Nothing
    SetControlState = Out
End Function

Public Function GetControlVisible(TagOrIndex As Variant) As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    GetControlVisible = oAC.Visible
End Function

Public Function SetControlVisible(TagOrIndex As Variant, ByVal Visible As Boolean) As Boolean
    Dim Out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Out = False
    If oAC.Visible <> Visible Then
        oAC.Visible = Visible
        ' return True when the visible is strictly changed
        Out = True
    End If
    
    SetControlVisible = Out
End Function

Public Function IncControlWidth(TagOrIndex As Variant, ByVal Width As Long) As Long
    Dim Out As Long
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    On Error Resume Next
    oAC.Width = oAC.Width + Width
    ' return the width accepted (tips: setting 0 to width makes it becomes default)
    Out = oAC.Width
    
    IncControlWidth = Out
End Function


'=== button tools for mybar end ===
'=== helper functions begin ===


Public Function NormalizeArray(x As Variant) As Variant
    On Error Resume Next
    Dim Out() As Variant
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
        ReDim Out(0 To U1 - L1)
        For i = 0 To UBound(Out)
            Out(i) = x(i + L1)
        Next
    Else
        ReDim Out(0 To U2 - L2)
        For i = 0 To UBound(Out)
            Out(i) = x(L1, i + L2)
            ' we pick up the 1st line only
        Next
    End If
    
    NormalizeArray = Out
End Function

Public Function StackToArray(pan As Variant) As Variant
    Dim Out() As Variant
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
    
    ReDim Out(0 To Counter - 1)
    i = 0
    Do Until IsEmpty(x)
        Out(i) = x(0)
        x = x(1)
        i = i + 1
    Loop
    
    StackToArray = Out
End Function


'=== helper functions end ===

'}}}

'module
'   name;DoNotepad
'{{{
Option Explicit
 
Private Const GW_CHILD = 5
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_CLOSE = &H10
Private Const EM_REPLACESEL = &HC2
Private Const EM_SETSEL = &HB1
Private Const EM_SETMODIFY = &HB9
Private Const HWND_BOTTOM = 1
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SW_RESTORE = 9
 
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 
' clear the save me flag
Public Function SetSavedNotepad(hWnd As Long) As Long
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    SendMessage i, EM_SETMODIFY, 0, 0
    SetSavedNotepad = i
End Function
 
' close the notepad
Public Sub CloseNotepad(hWnd As Long)
    SetSavedNotepad hWnd
    SendMessage hWnd, WM_CLOSE, 0, 0
End Sub
 
' kick up a new notepad process, return the hWnd
Public Function OpenNotepad(Optional iWindowState As Long = vbNormalFocus, _
            Optional NameMe As String = "") As Long
    Dim hWnd As Long
    Dim ProcID As Long
    Dim i As Long
    Dim TitleText As String
    Dim ExePath As String
    
    On Error GoTo Err1
    
    TitleText = " - notepad - meets VBA"
    ExePath = "notepad.exe"
    
    ProcID = Shell(ExePath, iWindowState)
    If ProcID = 0 Then GoTo Err1
    
    hWnd = GetWindowByProcessId(ProcID)
    If hWnd = 0 Then GoTo Err1
    
    TitleText = IIf(NameMe = "", ProcID, NameMe) & TitleText
    i = SetWindowText(hWnd, TitleText)
    'MoveWindow hWnd, 0, 50, 300, 200, 1
    ' SetWindowPos can be used to change Z-order
    'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    OpenNotepad = hWnd
    Exit Function
Err1:
    MsgBox "failed to start a notepad", vbExclamation Or vbOKOnly, Err.Number & " " & Err.Description
    OpenNotepad = 0
End Function

' repalce text at the notepad
Public Function WriteNotepad(hWnd As Long, strTextAll As String) As Boolean
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    WriteNotepad = _
        (0 <> SendMessageStr(i, WM_SETTEXT, 0, strTextAll))
End Function
 
' push text into the notepad with a linefeed
' iPos=0: at a cursor position
'     -1: at the first
'      1: at the last
Public Function WriteLineNotepad(hWnd As Long, strText As String, Optional iPos As Long = 0) As Boolean
    WriteLineNotepad = WriteTextNotepad(hWnd, strText & vbNewLine, iPos)
End Function
 
' push text into the notepad without a linefeed
' iPos=0: at a cursor position
'     -1: at the first
'      1: at the last
Public Function WriteTextNotepad(hWnd As Long, strText As String, Optional iPos As Long = 0) As Boolean
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    Select Case iPos
    Case -1
        SendMessage i, EM_SETSEL, 0, 0
    Case 1
        SendMessage i, EM_SETSEL, 0, -1     ' select all
        SendMessage i, EM_SETSEL, -1, 0     ' unselect (let the cursor move to the end of the selection)
    End Select
    WriteTextNotepad = _
        (0 <> SendMessageStr(i, EM_REPLACESEL, 0, strText))
End Function
 
' get text from the notepad
Public Function ReadNotepad(hWnd As Long) As String
    Dim i As Long
    Dim j As Long
    Dim x As String
    i = GetWindow(hWnd, GW_CHILD)
    j = 1 + SendMessage(i, WM_GETTEXTLENGTH, 0, 0)
    x = String(j, Chr(0))
    SendMessageStr i, WM_GETTEXT, j, x
    ReadNotepad = x
End Function

' get text length from the notepad
Public Function ReadNotepadLength(hWnd As Long) As Long
    Dim i As Long
    Dim j As Long
    i = GetWindow(hWnd, GW_CHILD)
    j = SendMessage(i, WM_GETTEXTLENGTH, 0, 0)
    ReadNotepadLength = j
End Function

' get a ProcessID from hWnd
Public Function GetWindowProcessId(hWnd As Long) As Long
    Dim ProcID As Long
    Dim ThreadID As Long
    ThreadID = GetWindowThreadProcessId(hWnd, ProcID)
    GetWindowProcessId = ProcID
End Function

' get an hWnd from ProcessID (Notepad only)
Public Function GetWindowByProcessId(ProcessId As Long, _
        Optional TaskName As String = "Notepad", _
        Optional TitleText As String = vbNullString) As Long
    Dim ProcID As Long
    Dim ThreadID As Long
    Dim hWnd As Long

    hWnd = 0
    Do
        hWnd = FindWindowEx(0, hWnd, TaskName, TitleText)
        If hWnd = 0 Then Exit Do
        ThreadID = GetWindowThreadProcessId(hWnd, ProcID)
    Loop Until ProcessId = ProcID
    
    GetWindowByProcessId = hWnd
End Function

' notify user to see the notepad described with the hWnd
Public Function ShowNotepad(hWnd As Long) As Boolean
    Dim Result As Long
    Result = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    ShowWindow hWnd, SW_RESTORE
    ShowNotepad = (Result <> 0)
End Function

'}}}

'module
'   name;NetworkGraphWriter
'{{{
Option Explicit

Dim NodeBag As Scripting.Dictionary
Dim ArcBag As Scripting.Dictionary
' these bags are just temporary storages,
' must be rebuild from worksheet.shapes on every run.


'=== entry point begin ===


Public Sub DoNode()
    AddNewNode ActiveSheet
End Sub

Public Sub DoArc()
    AddNewArc ActiveSheet, Selection
End Sub

Public Sub DoFlip()
    On Error GoTo NoShape
    FlipArc Selection.ShapeRange
NoShape:
End Sub

Public Sub DoNodeTsv()
    Dim St As StringStreamNotepad
    Set St = New StringStreamNotepad
    St.NameNotepad = "Node TSV " & Now()
    St.OpenWrite
    
    DumpNodeTsv St, ActiveSheet
    
    St.CloseWrite
    Set St = Nothing
End Sub

Public Sub DoArcTsv()
    Dim St As StringStreamNotepad
    Set St = New StringStreamNotepad
    St.NameNotepad = "Arc TSV " & Now()
    St.OpenWrite
    
    DumpArcTsv St, ActiveSheet
    
    St.CloseWrite
    Set St = Nothing
End Sub

Public Sub DoSem()
    Dim St As StringStreamNotepad
    Set St = New StringStreamNotepad
    St.NameNotepad = "SEM " & Now()
    St.OpenWrite
    
    DumpSem St, ActiveSheet
    
    St.CloseWrite
    Set St = Nothing
End Sub


'=== entry point end ===
'=== main functions begin ===


Private Sub AddNewNode(Target As Worksheet)
    Dim x As Shape
    Dim R As Variant
    Dim NewName As String
    Dim Current As Shape
    
    UpdateBags Target
    
    NewName = LastNodeName
    If NewName = "" Then
        Set Current = Nothing
    Else
        Set Current = Target.Shapes(NewName)
    End If
    
    Do
        NewName = SuggestNameNode(NewName)
        If NewName = "" Then Exit Sub
    Loop While NodeBag.Exists(NewName)
    
    R = SuggestRectNode(Current)
    Set x = Target.Shapes.AddShape(msoShapeRectangle, R(0), R(1), R(2), R(3))
    x.Name = NewName
    SetShapeText x, NewName
End Sub

Private Sub AddNewArc(Target As Worksheet, Selected As Object)
    Dim x As Shape
    Dim R As Variant
    Dim Connection As Variant
    Dim NewName As String
    
    UpdateBags Target
    
    NewName = LastArcName
    Do
        NewName = SuggestNameArc(NewName)
        If NewName = "" Then Exit Sub
    Loop While ArcBag.Exists(NewName)
    
    R = SuggestRectArc()
    Set x = Target.Shapes.AddConnector(msoConnectorStraight, R(0), R(1), R(2), R(3))
    x.Name = NewName
    
    Connection = SuggestConnection(Target, Selected)
    
    If Not Connection(0) Is Nothing Then
        x.ConnectorFormat.EndConnect Connection(0), 3
    End If
    If Not Connection(1) Is Nothing Then
        x.ConnectorFormat.BeginConnect Connection(1), 1
    End If
    If x.ConnectorFormat.BeginConnected = msoTrue And x.ConnectorFormat.EndConnected = msoTrue Then
        x.RerouteConnections
    End If
    
    With x.Line
        .EndArrowheadStyle = msoArrowheadStealth
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadWidth = msoArrowheadWide
    End With
End Sub

Private Sub FlipArc(Target As ShapeRange)
    On Error GoTo 0
    Dim x As Shape
    Dim KeepBegin As Shape
    Dim KeepBeginSite As Long
    
    For Each x In Target
        If x.Connector = msoTrue Then
            With x.ConnectorFormat
                If .BeginConnected = msoTrue And .EndConnected = msoTrue Then
                    Set KeepBegin = .BeginConnectedShape
                    KeepBeginSite = .BeginConnectionSite
                    .BeginConnect .EndConnectedShape, .EndConnectionSite
                    .EndConnect KeepBegin, KeepBeginSite
                    Set KeepBegin = Nothing
                End If
            End With
        End If
    Next
End Sub

Private Sub DumpNodeTsv(Stream As Object, Target As Worksheet)
    Const Delimiter = vbTab
    Dim x As Shape
    Dim xName As Variant
    
    UpdateBags Target
    Stream.WriteLine "Name" & Delimiter & "Text"
    
    For Each xName In NodeBag.Keys
        Set x = Target.Shapes(xName)
        Stream.WriteLine x.Name & Delimiter & GetShapeText(x)(0)
    Next
End Sub

Private Sub DumpArcTsv(Stream As Object, Target As Worksheet)
    Const Delimiter = vbTab
    Dim x As Shape
    Dim xName As Variant
    Dim BeginName As String
    Dim EndName As String
    
    UpdateBags Target
    Stream.WriteLine "Name" & Delimiter & "Begin" & Delimiter & "End"
    
    For Each xName In ArcBag.Keys
        Set x = Target.Shapes(xName)
        With x.ConnectorFormat
            If .BeginConnected = msoTrue Then
                BeginName = .BeginConnectedShape.Name
            Else
                BeginName = ""
            End If
            If .EndConnected = msoTrue Then
                EndName = .EndConnectedShape.Name
            Else
                EndName = ""
            End If
        End With
        Stream.WriteLine x.Name & Delimiter & BeginName & Delimiter & EndName
    Next
End Sub

Private Sub DumpSem(Stream As Object, Target As Worksheet)
    Dim x As Shape
    Dim xName As Variant
    Dim BeginName As String
    Dim EndName As String
    
    UpdateBags Target
    
    For Each xName In ArcBag.Keys
        Set x = Target.Shapes(xName)
        With x.ConnectorFormat
            If .BeginConnected = msoTrue And .EndConnected = msoTrue Then
                BeginName = .BeginConnectedShape.Name
                EndName = .EndConnectedShape.Name
                Stream.WriteLine BeginName & "->" & EndName & "," & x.Name & ",NA"
            End If
        End With
    Next
    
    For Each xName In NodeBag.Keys
        Set x = Target.Shapes(xName)
        Stream.WriteLine x.Name & "<->" & x.Name & ",e" & x.Name & ",NA"
    Next
End Sub


'=== main functions end ===
'=== bag manager begin ===


Private Sub NewNodeBag()
    On Error GoTo NewOnFail
    
    NodeBag.RemoveAll
    Exit Sub
    
NewOnFail:
    Set NodeBag = New Scripting.Dictionary
End Sub

Private Sub NewArcBag()
    On Error GoTo NewOnFail
    
    ArcBag.RemoveAll
    Exit Sub
    
NewOnFail:
    Set ArcBag = New Scripting.Dictionary
End Sub

Private Sub UpdateBags(Target As Worksheet)
    Dim x As Shape
    
    NewArcBag
    NewNodeBag
    
    For Each x In Target.Shapes
        Select Case x.Connector
        Case msoTrue    ' is Arc
            ArcBag.Add x.Name, Empty
        Case msoFalse   ' is Node
            NodeBag.Add x.Name, Empty
        End Select
    Next
End Sub


'=== bag manager end ===
'=== shape helper begin ===


Private Sub SetShapeText(Target As Shape, Text As String)
    Target.TextFrame.Characters.Text = Text
End Sub

Private Function GetShapeText(Target As Shape) As Variant
    On Error Resume Next
    Dim Text As String
    
    Text = ""
    ' below fails on excel 2000 rectangle without text
    Text = Target.TextFrame.Characters.Text
    
    GetShapeText = Array(Text)
End Function

Private Function SuggestConnection(Target As Worksheet, Selected As Object) As Variant
    Dim BeginX As Shape
    Dim EndX As Shape
    Dim BeginName As String
    Dim EndName As String
    
    If Not IsShapes(Selected, BeginX, EndX) Then IsShape Selected, BeginX
    
    If BeginX Is Nothing Then
        BeginName = SecondLastNodeName
        If BeginName <> "" Then Set BeginX = Target.Shapes(BeginName)
    End If
    
    If EndX Is Nothing Then
        EndName = LastNodeName
        If EndName <> "" Then Set EndX = Target.Shapes(EndName)
    End If
    
    SuggestConnection = Array(BeginX, EndX)
End Function

Private Function IsShapes(Selected As Object, ByRef First As Shape, Second As Shape) As Boolean
    Dim IsDrawingObjects As Boolean
    Dim i As Long
    Dim x As Object
    
    IsDrawingObjects = (TypeName(Selected) = "DrawingObjects")
    If IsDrawingObjects Then
        i = 1
        For Each x In Selected
            If i = 1 Then
                Set First = x.ShapeRange(1)
            ElseIf i = 2 Then
                Set Second = x.ShapeRange(1)
            Else
                Exit For
            End If
            i = i + 1
        Next
    End If
    
    IsShapes = IsDrawingObjects
End Function

Private Function IsShape(Selected As Object, ByRef First As Shape) As Boolean
    Dim IsSomeShape As Boolean
    
    Select Case TypeName(Selected)
    Case "Rectangle", "Oval", "TextBox"
        IsSomeShape = True
        Set First = Selected.ShapeRange(1)
    Case Else
        IsSomeShape = False
    End Select
    
    IsShape = IsSomeShape
End Function

Private Function SuggestRectNode(Current As Shape) As Variant
    If Current Is Nothing Then
        SuggestRectNode = Array(23.25, 23.25, 50.25, 23.25)
    Else
        SuggestRectNode = Array(Current.Left, Current.Top + 2 * Current.Height, Current.Width, Current.Height)
    End If
End Function

Private Function SuggestRectArc() As Variant
    SuggestRectArc = Array(0, 0, 100, 100)
End Function

Private Function SuggestNameNode(Current As String) As String
    Dim Candidate As String
    
    If Current = "" Then
        Candidate = "node1"
    Else
        Candidate = NextName(Current)
    End If
    
    SuggestNameNode = ShowInputBox("name me", "node name", Candidate)
End Function

Private Function SuggestNameArc(Current As String) As String
    Dim Candidate As String
    
    If Current = "" Then
        Candidate = "arc1"
    Else
        Candidate = NextName(Current)
    End If
    
    SuggestNameArc = ShowInputBox("name me", "arc name", Candidate)
End Function

Private Function ShowInputBox(BoxText As String, BoxTitle As String, Candidate As String) As String
    Dim x As String
    
    x = InputBox(BoxText, BoxTitle, Candidate)
    If LCase(x) = "false" Then x = ""
    ShowInputBox = x
End Function

Private Function LastNodeName() As String
    Dim Counter As Long
    
    Counter = NodeBag.Count
    If Counter = 0 Then
        LastNodeName = ""
    Else
        LastNodeName = NodeBag.Keys(Counter - 1)
    End If
End Function

Private Function SecondLastNodeName() As String
    Dim Counter As Long
    
    Counter = NodeBag.Count
    If Counter <= 1 Then
        SecondLastNodeName = ""
    Else
        SecondLastNodeName = NodeBag.Keys(Counter - 2)
    End If
End Function

Private Function LastArcName() As String
    Dim Counter As Long
    
    Counter = ArcBag.Count
    If Counter = 0 Then
        LastArcName = ""
    Else
        LastArcName = ArcBag.Keys(Counter - 1)
    End If
End Function


'=== shape helper end ===
'=== string helper begin ===


Private Function NextName(CurrentName As String) As String
    Dim R As RegExp
    Dim M As MatchCollection
    Dim Base As String
    Dim Index As String
    Dim Out As String
    
    Set R = RegSplitIndex
    Set M = R.Execute(CurrentName)
    If M.Count = 0 Then
        Out = CurrentName & "2"
    Else
        Base = M(0).SubMatches(0)
        Index = M(0).SubMatches(1)
        If Index = "" Then
            Out = CurrentName & "2"
        Else
            Out = Base + CStr(Val(Index) + 1)
        End If
    End If
    
    NextName = Out
    Set R = Nothing
End Function

Function RegSplitIndex() As RegExp
    Dim R As RegExp
    Set R = New RegExp
    R.Global = True
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "^(.*?)(\d*)$"
    Set RegSplitIndex = R
End Function


'=== string helper end ===

'}}}

'class
'   name;StringStreamNotepad
'{{{
Option Explicit

Public Text As String
Public LineFeed As String
Public hWndNotepad As Long
Public NameNotepad As String

Public Property Get AtEndOfStream()
    AtEndOfStream = IsEmpty
End Property

Public Property Get EOS()
    EOS = IsEmpty
End Property

Public Function ReadAll()
    ReadAll = Dequeue(-1)
End Function

Public Function ReadLine()
    ReadLine = Dequeue(-1)
End Function

Public Function ReadText(Size)
    ReadText = Dequeue(Size)
End Function

Public Sub WriteLine(Data)
    Enqueue Data
    Enqueue LineFeed
End Sub

Public Sub WriteText(Data)
    Enqueue Data
End Sub

Public Sub FlushRead()
End Sub

Public Sub FlushWrite()
    If hWndNotepad = 0 Then Exit Sub
    WriteNotepad hWndNotepad, Dequeue(-1)
End Sub

Public Function OpenRead()
    ClearAll
    Enqueue ReadNotepad(hWndNotepad)
    OpenRead = (hWndNotepad <> 0)
End Function

Public Function OpenWrite()
    ClearAll
    hWndNotepad = OpenNotepad(1, NameNotepad)
    OpenWrite = (hWndNotepad <> 0)
End Function

Public Sub CloseRead()
    FlushRead
    ClearAll
End Sub

Public Sub CloseWrite()
    FlushWrite
    ClearAll
End Sub

Public Function OpenSpawn(SpawnName)
    Set OpenSpawn = New StringStreamNotepad
    With OpenSpawn
        .LineFeed = LineFeed
        .NameNotepad = SpawnName
        If Not .OpenWrite Then Set OpenSpawn = Nothing
    End With
End Function

Public Function CloseSpawn()
    CloseWrite
    If ReadNotepadLength(hWndNotepad) = 0 Then CloseNotepad (hWndNotepad)
    Set CloseSpawn = Nothing
End Function

Private Sub Class_Initialize()
    Text = ""
    LineFeed = vbCrLf
    hWndNotepad = 0
    NameNotepad = ""
End Sub

Private Sub ClearAll()
    Text = ""
End Sub

Private Sub Enqueue(Data)
    Text = Text & Data
End Sub

Private Function Dequeue(Size)
    If Size > 0 Then
        Dequeue = Left(Text, Size)
        Text = Mid(Text, Size + 1)
    Else
        Dequeue = Text
        Text = ""
    End If
End Function

Private Function IsEmpty()
    IsEmpty = (Text = "")
End Function

'}}}

'code
'   name;ThisWorkbook
'{{{
Option Explicit

Private Sub Workbook_Open()
    BaumMain.BarInitialize
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    BaumMain.BarTerminate
End Sub

'}}}

'ssf-end


```