

# Introduction #

  * convert characters between hankaku and zenkaku in excel cells
  * work as an addin

## 概要 ##
  * エクセルでセル文字列を半角または全角に変換する
  * アドインとして動く

# Details #

## 説明 ##
  * アドインを有効にすると、操作用のボタンがツールバー（リボン）に追加される。
    * 全角へ、ｶﾅ全、数全、英全、半角へ、ｶﾅ半、数半、英半
  * ボタンを押すと、それぞれの文字種への変換を行う。
    1. 全角へ、半角へ、のボタンはすべてを変換。
    1. カタカナは、カタカナと関連する記号
    1. 数字は、0から9まで
    1. 英字は、アルファベットと関連する記号
  * 変換の範囲は、選択したセル範囲、選択した列、選択した行など。
  * 変換による副作用。
    1. 選択した範囲は、すべて文字列として変換を行う。
    1. 数値、日付、数式が入っていた場合も、そのまま文字列として変換対象にする。
      * たとえば =Sum(A1:B1) が入っていた場合、この式そのものを文字として全角文字 ＝Ｓｕｍ（Ａ１：Ｂ１） に変換する。
    1. 選択した範囲のセル書式の表示形式（日付や数値桁などの表現）を、すべて `標準` にする。
      * たとえば日付データ 2011/4/4 15:50 が入っていて、表示形式で ４月４日（月）になっていた場合、これを英字全角変換すると、 2011／4／4 15：50 という文字になる。

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=convert_char_zenkaku)

# How to use #


## 使い方 ##
  * アドインのインストールと有効化
    1. アドインファイルを、アドイン用フォルダにコピーする。
      * アドインの設定画面で、参照ボタンを押したときに表示されるフォルダなど、アドイン用に信頼した場所を使う。
    1. アドインの設定画面で、「全角半角変換」にチェックを入れて、有効化する。
      * 一覧に表示されない場合、参照ボタンで読み込む。
      * 有効化しておけば、エクセル起動だけでアドインが自動で読み込まれる。
      * ![http://1.bp.blogspot.com/-IKIU0wRY5j4/TZ_QpcEGJTI/AAAAAAAAAEw/__X7C-dGPTw/s1600/shot3.png](http://1.bp.blogspot.com/-IKIU0wRY5j4/TZ_QpcEGJTI/AAAAAAAAAEw/__X7C-dGPTw/s1600/shot3.png)
    1. アンインストール
      * 上記の逆手順、アドインの無効化、アドインファイルの削除、を行う。
      * 追加修正のバージョンをインストールするとき、同じファイル名で上書きするのでなければ、先に古いバージョンをアンインストールしないといけない。
  * 基本的な使い方
    1. 変換対象のブックで、変換したい範囲を選ぶ。列、行、セル範囲、シート全体など。
      * Ctrl キーを使って、複数箇所を選んでもよい。
      * 何も選ばないと、カーソルのあるセル１つだけを変換する。
    1. 変換したいキーを押す。
      * 正常動作すればメッセージなどは出ず、通常は一瞬で処理が終わる。
  * もっと詳しい情報
    1. 非対称なカナ変換
      * カナ変換では、全角への変換と半角への変換対象が異なるため、行きと帰りで処理が異なる。
      * 半角カナを全角に変換する場合、カナ記号はすべて全角に変換。
      * 片仮名の全角カナを半角に変換する場合、カナ記号はひらがなや漢字として扱うため変換せず、片仮名だけを半角に変換する。
      * カナ長音 `ー` は、直前の文字に応じて変換する。
    1. スペースの変換
      * 半角、全角のスペースは、文字種類によらない変換のときのみ、変換対象となる。
        * 全角へ、半角へ、のボタン　 －－－ 　スペースを変化させる。
        * それ以外のボタン　　　　　　 －－－ 　スペースを変化させない。
    1. 数式の扱い
      * セル数式の扱いを指定できる。
        1. 式文字：　数式を文字として扱い、変換対象とする。
        1. 式無視：　数式セルは変換しない。
        1. 値固定：　数式の計算結果を文字として扱い、変換対象とする。

### 2011/4/13 の修正 ###

  1. 全角カナ長音 `ー` を、直前の文字に応じて変換する。
    * `ｶﾅ半` → 直前の文字が、全角カナの場合のみ、半角カナ長音 `ｰ` に変換する。
    * `半角へ` → 直前の文字が、全角英数カナスペースの場合のみ、半角カナ長音 `ｰ` に変換する。
    * 上記いずれも、そのボタンが変換対象とする文字に続いて現れた長音を変換する。
  1. 全角ハイフン `－` と半角ハイフン `-` を、直前の文字に応じて変換する。
    * `数半` → 直前の文字が、全角数字の場合のみ、半角ハイフン `-` に変換する。
    * `数全` → 直前の文字が、半角数字の場合のみ、全角ハイフン `－` に変換する。
    * `半角へ` `全角へ` `英半` および `英全` ボタンでは、既に、ハイフンの変換は行われており、その動作は変更しない。


# Snapshots #

  * エクセル2007 ではリボンのアドインタブにボタンを表示する。
> > ![http://2.bp.blogspot.com/-CesX5NTgjOs/TZ_QofQspzI/AAAAAAAAAEo/cEl4tcQA4Zc/s1600/shot1.png](http://2.bp.blogspot.com/-CesX5NTgjOs/TZ_QofQspzI/AAAAAAAAAEo/cEl4tcQA4Zc/s1600/shot1.png)

  * セル数式の扱いを指定する。
> > ![http://2.bp.blogspot.com/-n63FTS-IhQU/TZ_Qo9CZCcI/AAAAAAAAAEs/eTQXfGHlPCo/s1600/shot2.png](http://2.bp.blogspot.com/-n63FTS-IhQU/TZ_Qo9CZCcI/AAAAAAAAAEs/eTQXfGHlPCo/s1600/shot2.png)

  * アドインの設定画面で有効にする。
> > ![http://1.bp.blogspot.com/-IKIU0wRY5j4/TZ_QpcEGJTI/AAAAAAAAAEw/__X7C-dGPTw/s1600/shot3.png](http://1.bp.blogspot.com/-IKIU0wRY5j4/TZ_QpcEGJTI/AAAAAAAAAEw/__X7C-dGPTw/s1600/shot3.png)



# Code #

```
'ssf-begin

'workbook
'   name;convert_char_zenkaku.xls/F3ConvertCharZenkaku

'require
'       ;{3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5 5 Microsoft VBScript Regular Expressions 5.5

'worksheet
'   name;全角半角変換/BaumMain

'cells-formula
'  address;A1:M23
'         ;名称
'         ;convert_char_zenkaku
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;コメント
'         ;エクセルでセル文字列を半角または全角に変換する
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;comment
'         ;convert characters between hankaku and zenkaku in excel cells
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;著作権
'         ;="Copyright (C) " &R[3]C & "-" & YEAR(R[5]C) & " " & R[2]C
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;ライセンス
'         ;自律, 自由, 公正, http://cowares.nobody.jp
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;作者
'         ;Tomizono - kobobau.com
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;初版
'         ;2011
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;配布元
'         ;http://cowares.blogspot.com/search/label/baum
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;更新
'         ;40646.625
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;keyword
'         ;vba,excel
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
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
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;全角へ
'         ;選択した範囲を、すべて全角にする。
'         ;zen_all
'         ;
'         ;1
'         ;2
'         ;
'         ;1
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;ｶﾅ全
'         ;選択した範囲を、カナだけ全角にする。
'         ;zen_kana
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;数全
'         ;選択した範囲を、数字だけ全角にする。
'         ;zen_num
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;英全
'         ;選択した範囲を、英字だけ全角にする。
'         ;zen_alpha
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;半角へ
'         ;選択した範囲を、すべて半角にする。
'         ;han_all
'         ;
'         ;1
'         ;2
'         ;
'         ;1
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;ｶﾅ半
'         ;選択した範囲を、カナだけ半角にする。
'         ;han_kana
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;数半
'         ;選択した範囲を、数字だけ半角にする。
'         ;han_num
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;英半
'         ;選択した範囲を、英字だけ半角にする。
'         ;han_alpha
'         ;
'         ;1
'         ;2
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;数式
'         ;セル数式の扱い。
'         ;formula
'         ;
'         ;3
'         ;
'         ;
'         ;1
'         ;
'         ;式文字
'         ;式無視
'         ;値固定
'         ;

'cells-numberformat
'  address;A1:M23
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;m/d/yyyy h:mm
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General

'cells-name
'       ;=全角半角変換!R15C1
'       ;全角半角変換!_ButtonCaption
'       ;=全角半角変換!R3C2
'       ;全角半角変換!_Comment
'       ;=全角半角変換!R6C2
'       ;全角半角変換!_Contributor
'       ;=全角半角変換!R4C2
'       ;全角半角変換!_Copyright
'       ;=全角半角変換!R5C2
'       ;全角半角変換!_License
'       ;=全角半角変換!R2C2
'       ;全角半角変換!_LocalComment
'       ;=全角半角変換!R1C2
'       ;全角半角変換!_PublicName
'       ;=全角半角変換!R7C2
'       ;全角半角変換!_Since
'       ;=全角半角変換!R10C2
'       ;全角半角変換!_Tag
'       ;=全角半角変換!R9C2
'       ;全角半角変換!_Timestamp
'       ;=全角半角変換!R8C2
'       ;全角半角変換!_Url

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
    Dim Pan As Variant
    Dim i As Long
    
    On Error GoTo DONE
    Pan = Empty
    i = 9
    
    Do Until IsEmpty(ItemAButtonData(Data, i))
        Pan = Array(ItemAButtonData(Data, i), Pan)
        i = i + 1
    Loop
    
DONE:
    ButtonItems = Pan
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
    Dim Pan As Variant
    Dim HasItem As Boolean
    
    Pan = ButtonItems(Data)
    HasItem = False
    
    Do Until IsEmpty(Pan)
        ButtonA.AddItem Pan(0), 1
        Pan = Pan(1)
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

Public Function StackToArray(Pan As Variant) As Variant
    Dim out() As Variant
    Dim x As Variant
    Dim i As Long
    Dim Counter As Long
    
    x = Empty
    Counter = 0
    Do Until IsEmpty(Pan)
        x = Array(Pan(0), x)
        Pan = Pan(1)
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
    VBProjectName = "F3ConvertCharZenkaku"
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
        MyBarMain Me
    End If
End Sub

Public Sub OnButtonToggle()
    Helper.OnButtonToggle
End Sub

' followings need to be public, because they are called from outside by the Helper
' we also can copy the Helper.BarMain code here, and let the followings be private.

Private Sub MyBarMain(Optional oWho As Object = Nothing)
    Dim oAC As Object   ' this is the button itself pressed
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    
    ' no needs to switch by menu name, except one
    If oAC.Tag <> "formula" Then
        ConvertCharZenkaku.DoConvert MakeConvertData(oAC)
    End If
    Set oAC = Nothing
End Sub

Private Function MakeConvertData(oAC As Object) As Variant
    Dim ZenAll As Variant
    ZenAll = Split(oAC.Tag, "_")
    MakeConvertData = Array(Selection, ZenAll(0), ZenAll(1), Helper.GetButton("formula").ListIndex)
End Function


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
    Dim out As Range
    Dim StartAt As Range
    
    Set StartAt = Me.Range(Address)
    If IsEmpty(StartAt.Offset(1, 0).Value) Then
        Set out = StartAt
    Else
        Set out = Me.Range(StartAt, StartAt.End(xlDown))
    End If
    
    Set GetButtonRow = out.EntireRow
End Function

Private Function GetButtonCol(Optional Address As String = "_ButtonCaption") As Range
    Dim StartAt As Range
    Set StartAt = Me.Range(Address)
    Set GetButtonCol = Me.Range(StartAt, StartAt.SpecialCells(xlCellTypeLastCell)).EntireColumn
End Function

Private Function ConvertRangeToArray(Ra As Range) As Variant
    Dim out() As Variant
    Dim i As Long
    
    ReDim out(0 To Ra.Rows.Count - 1)
    For i = 0 To UBound(out)
        out(i) = Ra.Rows(i + 1).Value
    Next
    
    ConvertRangeToArray = out
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

'module
'   name;ConvertCharZenkaku
'{{{
Option Explicit

'data structure
Private Sub ExtractConvertData(Data As Variant, _
        ByRef Target As Object, _
        ByRef ZenHan As Variant, _
        ByRef CharGroup As Variant, _
        ByRef FormulaListIndex As Long)
    Set Target = Data(0)
    ZenHan = Data(1)
    CharGroup = Data(2)
    FormulaListIndex = Data(3)
End Sub

Private Function ExtractFormulaListIndex(Data As Variant) As Long
    ExtractFormulaListIndex = Data(3)
End Function

Private Function MakeAreaData(Target As Range, ConvertParams As Variant, UseFormula As Boolean) As Variant
    MakeAreaData = Array(Target, ConvertParams, UseFormula)
End Function

Private Sub ExtractAreaData(Data As Variant, ByRef Target As Range, ByRef ConvertParams As Variant, ByRef UseFormula As Boolean)
    Set Target = Data(0)
    ConvertParams = Data(1)
    UseFormula = Data(2)
End Sub

Private Function MakeConvertParams(ConvertTo As VbStrConv, RegConv As RegExp) As Variant
    MakeConvertParams = Array(ConvertTo, RegConv)
End Function

Private Sub ExtractConvertParams(Params As Variant, ByRef ConvertTo As VbStrConv, ByRef RegConv As RegExp)
    ConvertTo = Params(0)
    Set RegConv = Params(1)
End Sub

Public Sub DoConvert(Data As Variant)
    Dim TargetAreas As Object
    Dim CellsCount As Long
    Dim DisableUpdate As Boolean
    Dim KeepCalculation As Long
    Dim AreaData As Variant
    
    CellsCount = SmarterTargetRange(Data, TargetAreas)
    If CellsCount = 0 Then Exit Sub
    
    DisableUpdate = (CellsCount > 16)
    
    If DisableUpdate Then
        Application.ScreenUpdating = False
        KeepCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
    End If
    
    For Each AreaData In TargetAreas
        'Debug.Print AreaData(0).Address, AreaData(2)
        AreaConvert AreaData
    Next
    Set TargetAreas = Nothing
    
    If DisableUpdate Then
        Application.Calculation = KeepCalculation
        Application.ScreenUpdating = True
    End If
End Sub

Private Function SmarterTargetRange(Data As Variant, ByRef AreaCollection As Object) As Long
    Dim SmallerTarget As Range
    Dim AnArea As Range
    Dim out As Long
    Dim Target As Object
    Dim ZenHan As Variant
    Dim CharGroup As Variant
    Dim FormulaListIndex As Long
    Dim ConvertParams As Variant
    
    Set AreaCollection = New Collection    ' as a blank array object
    ExtractConvertData Data, Target, ZenHan, CharGroup, FormulaListIndex
    
    If TypeName(Target) <> "Range" Then
        out = 0
    Else
        Set SmallerTarget = Application.Intersect(Target, Target.Worksheet.UsedRange)
        If SmallerTarget Is Nothing Then
            out = 0
        Else
            out = SmallerTarget.Cells.Count
            ConvertParams = MakeConvertParams(GetConvertDirection(ZenHan), GetConvertRegExp(ZenHan, CharGroup))
            For Each AnArea In SmallerTarget.Areas
                SmarterAppend AreaCollection, AnArea, FormulaListIndex, ConvertParams
            Next
        End If
    End If
    
    SmarterTargetRange = out
End Function

Private Sub SmarterAppend(ByRef AreaCollection As Object, Target As Range, FormulaListIndex As Long, ConvertParams As Variant)
    Dim FormulaTarget As Range
    Dim ValueTarget As Range
    Dim HasFormulaAndValue As Boolean
    Dim HasNoFormula As Boolean
    
    ' FormulaListIndex: 1=formula string, 2=no touch formula, 3=value string
    
    HasFormulaAndValue = IsNull(Target.HasFormula)
    If HasFormulaAndValue Then
        HasNoFormula = False
    Else
        HasNoFormula = Not Target.HasFormula
        If HasNoFormula Then
            ' everything is value, equivalent to the mode 3
            FormulaListIndex = 3
        Else
            ' everything is formula, so nothing to do
            If FormulaListIndex = 2 Then Exit Sub
        End If
    End If
    
    ' value to string
    If FormulaListIndex = 3 Then
        ' convert everything as value
        SmarterAppendDivided AreaCollection, Target, ConvertParams, False
    Else
        ' choose value cells, and convert
        If SpecialCellsValue(Target, ValueTarget) Then
            SmarterAppendDivided AreaCollection, ValueTarget, ConvertParams, False
        End If
    End If
    
    ' formula to string
    If (FormulaListIndex = 1) And Not HasNoFormula Then
        If HasFormulaAndValue Then
            ' choose formula cells, and convert
            If SpecialCellsFormula(Target, FormulaTarget) Then
                SmarterAppendDivided AreaCollection, FormulaTarget, ConvertParams, True
            End If
        Else
            ' everything is formula
            SmarterAppendDivided AreaCollection, Target, ConvertParams, True
        End If
    End If
End Sub

Private Sub SmarterAppendDivided(ByRef AreaCollection As Object, _
        Target As Range, ConvertParams As Variant, UseFormula As Boolean)
    
    Const MaximumCellsAtOnce = 8000
    
    Dim AnArea As Range
    Dim Divide As Long
    
    Dim ColAddress As String
    Dim RowAddress As String
    Dim C1 As String
    Dim C2 As String
    Dim R1 As Long
    Dim R2 As Long
    Dim R As Long
    Dim Rm As Long
    Dim StepR As Long
    Dim Cs As Variant
    Dim Rs As Variant
    Dim NewAddress As String
    
    For Each AnArea In Target.Areas
        Divide = Int(AnArea.Cells.Count / MaximumCellsAtOnce) + 1
        If Divide = 1 Then
            AreaCollection.Add MakeAreaData(AnArea, ConvertParams, UseFormula)
        Else
            ColAddress = AnArea.EntireColumn.Address(False, False, xlA1, False)
            RowAddress = AnArea.EntireRow.Address(False, False, xlA1, False)
            Cs = Split(ColAddress, ":")
            Rs = Split(RowAddress, ":")
            C1 = Cs(0)
            C2 = Cs(1)
            R1 = Rs(0)
            R2 = Rs(1)
            
            StepR = (R2 - R1 + 1) / Divide
            If StepR < 1 Then StepR = 1
            
            For R = R1 To R2 Step StepR
                Rm = R + StepR - 1
                If Rm > R2 Then Rm = R2
                NewAddress = C1 & R & ":" & C2 & Rm
                'Debug.Print NewAddress
                AreaCollection.Add MakeAreaData(AnArea.Worksheet.Range(NewAddress), ConvertParams, UseFormula)
            Next
        End If
    Next
End Sub

' temporary switch cells style into string only
' this will avoid verbose formulas error, on updating cell values
Private Sub SetStringStyle(Target As Range)
    Target.NumberFormat = "@"
End Sub

Private Sub SetGeneralStyle(Target As Range)
    Target.NumberFormat = "General"
End Sub

Private Sub AreaConvert(Data As Variant)
    Dim Target As Range
    Dim ConvertParams As Variant
    Dim UseFormula As Boolean
    
    ExtractAreaData Data, Target, ConvertParams, UseFormula
    SetStringStyle Target
    CellsConvert Target, ConvertParams, UseFormula
    SetGeneralStyle Target
End Sub

Private Sub CellsConvert(Target As Range, ConvertParams As Variant, UseFormula As Boolean)
    Dim OldMatrix As Variant
    Dim NewMatrix() As Variant
    Dim Cols As Long
    Dim Rows As Long
    Dim R As Long
    Dim c As Long
    Dim TargetAddress As String
    
    On Error GoTo ProtectedCellOrSo
    TargetAddress = Target.Address(False, False, xlA1, False)
    
    If Target.Cells.Count = 1 Then
        Target.Value = TextConvertKeepEmpty(FormulaOrValue(Target, UseFormula), ConvertParams)
    Else
        OldMatrix = FormulaOrValue(Target, UseFormula)
        Rows = UBound(OldMatrix, 1)
        Cols = UBound(OldMatrix, 2)
        ReDim NewMatrix(1 To Rows, 1 To Cols)
        For R = 1 To Rows
            For c = 1 To Cols
                NewMatrix(R, c) = TextConvertKeepEmpty(OldMatrix(R, c), ConvertParams)
            Next
        Next
        Target.Value = NewMatrix
    End If
    
    Exit Sub
    
ProtectedCellOrSo:
    MsgBox TargetAddress & "の読み書きに失敗しました。", vbExclamation Or vbOKOnly, Err.Number & " " & Err.Description
End Sub

Private Function SpecialCellsFormula(Target As Range, ByRef ReturnSpecial As Range) As Boolean
    On Error GoTo NoCellsFound
    
    If Target.Cells.Count = 1 Then
        ' SpecialCells doesn't work for a single cell
        If Target.HasFormula Then
            Set ReturnSpecial = Target
            SpecialCellsFormula = True
        Else
            Set ReturnSpecial = Nothing
            SpecialCellsFormula = False
        End If
    Else
        Set ReturnSpecial = Target.SpecialCells(xlCellTypeFormulas)
        SpecialCellsFormula = True
    End If
    
    Exit Function
    
NoCellsFound:
    Set ReturnSpecial = Nothing
    SpecialCellsFormula = False
End Function

Private Function SpecialCellsValue(Target As Range, ByRef ReturnSpecial As Range) As Boolean
    On Error GoTo NoCellsFound
    
    If Target.Cells.Count = 1 Then
        ' SpecialCells doesn't work for a single cell
        If Target.HasFormula Then
            Set ReturnSpecial = Nothing
            SpecialCellsValue = False
        Else
            Set ReturnSpecial = Target
            SpecialCellsValue = True
        End If
    Else
        Set ReturnSpecial = Target.SpecialCells(xlCellTypeConstants)
        SpecialCellsValue = True
    End If
    
    Exit Function
    
NoCellsFound:
    Set ReturnSpecial = Nothing
    SpecialCellsValue = False
End Function

Private Function FormulaOrValue(Target As Range, UseFormula As Boolean) As Variant
    If UseFormula Then
        FormulaOrValue = Target.Formula
    Else
        FormulaOrValue = Target.Value
    End If
End Function

Private Function TextConvertKeepEmpty(TextEmptyError As Variant, ConvertParams As Variant) As String
    Dim Text As String
    
    Text = CStr(TextEmptyError)
    If Text = "" Then
        TextConvertKeepEmpty = vbNullString
    Else
        TextConvertKeepEmpty = TextConvert(Text, ConvertParams)
    End If
End Function

Private Function TextConvert(ByVal Text As String, ConvertParams As Variant) As String
    Dim out As String
    
    out = ""
    Do Until Text = ""
        out = out & TextConvertFirstMatch(Text, ConvertParams)
    Loop
    
    TextConvert = out
End Function

Private Function TextConvertFirstMatch(ByRef Text As String, ConvertParams As Variant) As String
    ' we use this for all conversion, because the StrConv loses some non-local characters
    Dim ConvertTo As VbStrConv
    Dim RegConv As RegExp
    Dim Matched As MatchCollection
    Dim M As Match
    Dim out As String
    
    ExtractConvertParams ConvertParams, ConvertTo, RegConv
    Set Matched = RegConv.Execute(Text)
    If Matched.Count = 0 Then
        out = Text
        Text = ""
    Else
        Set M = Matched(0)
        out = Left(Text, M.FirstIndex) & StrConvJa(Mid(Text, M.FirstIndex + 1, M.Length), ConvertTo)
        Text = Mid(Text, M.FirstIndex + M.Length + 1)
        Set M = Nothing
    End If
    
    Set Matched = Nothing
    TextConvertFirstMatch = out
End Function

Private Function GetConvertDirection(ZenHan As Variant) As Long
    If ZenHan = "zen" Then
        GetConvertDirection = vbWide
    Else    ' han
        GetConvertDirection = vbNarrow
    End If
End Function

Private Function GetConvertRegExp(ZenHan As Variant, CharGroup As Variant) As RegExp
    Dim R As RegExp
    
    Set R = New RegExp
    R.Global = False
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = GetConvertRegPattern(ZenHan, CharGroup)
    
    Set GetConvertRegExp = R
End Function

Private Function GetConvertRegPattern(ZenHan As Variant, CharGroup As Variant) As String
    Dim Pattern As String
    Dim Pattern2 As String
    
    Pattern2 = ""
    Select Case CharGroup
    Case "kana"
        If ZenHan = "zen" Then
            Pattern = "｡-ﾟ"
        Else
            Pattern = "ァ-ヶ"
            Pattern2 = "ー"
        End If
    Case "num"
        If ZenHan = "zen" Then
            Pattern = "0-9"
            Pattern2 = "\x2d"   ' -
        Else
            Pattern = "０-９"
            Pattern2 = "－"
        End If
    Case "alpha"
        If ZenHan = "zen" Then
            Pattern = "!-/:-~"
        Else
            Pattern = "！-／：-～￥"
        End If
    Case Else   ' all
        If ZenHan = "zen" Then
            Pattern = " !-~｡-ﾟ"
        Else
            Pattern = "　！-～￥ァ-ヶ"
            Pattern2 = "ー"
        End If
    End Select
    
    Pattern = "[" & Pattern & "][" & Pattern & Pattern2 & "]*"
    GetConvertRegPattern = Pattern
End Function

Private Function StrConvJa(ByVal Text As String, Conversion As VbStrConv) As String
    ' overrides system locale on a non-japanese os
    ' adds japanese specific conversions against the i18n strconv
    Const LocaleIdJa = 1041
    
    Select Case Conversion
    Case vbWide
        ' StrConv doesn't convert into Zenkaku-Yen
        Text = Replace(Text, "\", "￥")
    Case vbNarrow
        ' StrConv supports this one way conversion
        'Text = Replace(Text, "￥", "\")
    End Select
    
    Text = StrConv(Text, Conversion, LocaleIdJa)
    
    StrConvJa = Text
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