

# Introduction #

  * verify input for excel against a reference sheet
  * [verify\_sheet\_input](verify_sheet_input.md) with a marker to tell differences by characters of a long text

## 概要 ##
  * エクセルで別のシートを参照して入力のベリファイを行う
  * [verify\_sheet\_input](verify_sheet_input.md) に、 `文字単位のマーカー` を追加したもの

# Details #

  * this is a system for verifying data entry, a method of double entry until match.
    * you can do the 1st entry on an arbitrary sheet. it is called as a reference sheet and a one to distribute, maybe.
    * you must do the 2nd entry on the specific sheet in this system, because it has the vba code.
  * it does live verifications on the 2nd entry.
    * the cell color of the 2nd sheet changes to alarm when you enter a different data against the 1st one.
    * strings after the first defference are colored for a mark.
    * the alarm disappears when you correct the failure.
    * will verify not only for entering, but also skipping by tab keys.
  * a toolbar or ribbon appears after the book opened.
    * to specify a reference sheet.
    * marking by characters can be disabled by toolbar button, to get speed.

## 説明 ##
  * 同一のデータをシートに２回入力することで、入力検証を行う。
    * １次入力は提出用（参照用）シートに行う。これは任意のブック、シートで構わない。
    * ２次入力は、検証専用シートに行う。検証用シートがマクロを持っている。
  * 入力検証は、検証用への入力の都度、自動で行われる。
    * １次入力と違う文字が入力されたら、検証用シートのセルが警告色になる。
    * `文字入力の場合、最初に異なる文字以降が別色になり識別できる。`
    * 正しく入力しなおせば色は元に戻る。
    * エンターだけでなく、タブキーなどによる移動でも検証する。
  * 検証用ブックを開くと、操作用のボタンがツールバー（リボン）に追加される。
    * 参照用シートの設定など
    * `文字単位のマーカー機能は、速度重視の場合、ツールバーのスイッチでオフにできる。`

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=verify_sheet_input_char)

# How to use #

  1. prepare to verify
    1. a toolbar appears when you open the book. (buttons in the Addin Ribbon for Excel 2007 or later)
    1. activate the 1st entry sheet done, and press the `参照シート設定` button.
    1. this activates the `Veify` sheet and sets the name of the 1st entry sheet as a reference on the button face.
  1. verifying
    1. input data in the `Verify` sheet for the 2nd entry.
    1. the system verify when changing cell values, skipping cell by tab key or mouse, deleting cells and copying.
    1. the cell in the `Verify` sheet turns yellow on validation failure, and turns white when mistakes are corrected.
    1. strings after the first different character are turn red, and turn black when mistakes are corrected.
  1. what the 参照シート設定 button does
    1. for sheets except `Verify`: make the sheet the 1st entry sheet, and jump to the `Verify` sheet.
    1. for `Verify` sheet: reset current settings.
    1. it shows as 参照シート無し when no 1st entry sheets are available.
    1. no verifications run for no reference sheets.
  1. what the 差分 button does
    1. toggles between on and off.
    1. the marker function works when on.
  1. the character marker doesn't seem working (not turning to red) when
    1. input data is shorter than the reference, and matched to the length.
    1. input data has an extra space at the end.
    1. the input data or the reference does not have strings.

## 使い方 ##
  1. 検証準備
    1. ブックを開くと、ツールバーが出る。（エクセル2007以降ではリボンのアドイン）
    1. １次入力したシートを開いて、 `参照シート設定` ボタンを押す。
    1. `Verify` シートが自動でアクティブになり、ボタンに参照先シート名が表示される。
  1. 検証
    1. `Verify` シートに２次入力を行う。
    1. シートへの入力、タブやマウス等によるスキップ、削除、コピーに反応し、検証を実施。
    1. １次と２次のセル内容が異なる場合、 `Verify` のセルが黄色くなり、正しく入力しなおせば白くなる。
    1. さらに、文字単位で、最初に異なる位置以降が赤字になり、、正しく入力しなおせば黒くなる。
  1. 参照シート設定ボタンの動作
    * `Verify` 以外のシート：　そのシートを１次入力シートに設定し、 `Verify` にジャンプする。
    * `Verify` シート：　現在の設定を解除する。
    * 登録された１次シートが利用できない場合、参照シート無しと表示される。
    * 参照シートが設定されている時だけ、検証が働く。
  1. 差分ボタンの動作
    * オン、オフを切り替える動作をする。
    * オンの時には、文字単位のマーカー機能が働く。
  1. 文字単位のマーカーが働かない（文字が赤くならない）ケース
    * 検証側の文字が参照側より短く、入力した範囲では合致している。
    * 検証側の最後に余分なスペースを入れた。
    * 検証側と参照側の両方に文字が入力されていない。

# Code #

```
'workbook
'  name;verify_sheet_input_char.xls

'require

'worksheet
'  name;Verify

'worksheet
'  name;about/SSF

'cells-formula
'  address;A1:B16
'         ;名称
'         ;verify_sheet_input_char
'         ;コメント
'         ;エクセルで別のシートを参照して入力のベリファイを行う
'         ;comment
'         ;verify input for excel against a reference sheet
'         ;著作権
'         ;="Copyright (C) " &R[3]C & "-" & YEAR(R[5]C) & " " & R[2]C
'         ;ライセンス
'         ;自律, 自由, 公正, http://cowares.nobody.jp
'         ;作者
'         ;Tomizono - kobobau.com
'         ;初版
'         ;2011
'         ;配布元
'         ;http://code.google.com/p/cowares-excel-hello/wiki/verify_sheet_input_char
'         ;更新
'         ;40602.6032638889
'         ;keyword
'         ;excel,validation
'         ;
'         ;
'         ;
'         ;
'         ;ボタンの表示
'         ;ボタンの機能
'         ;
'         ;
'         ;参照シート設定
'         ;アクティブシートを参照用に設定します
'         ;差分
'         ;文字単位のマーカーの有効／無効を切り替えます。

'cells-numberformat
'  address;A1:B16
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
'         ;General
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

'cells-name
'  ;=about!R15C1
'  ;about!_ButtonCaption
'  ;=about!R3C2
'  ;about!_Comment
'  ;=about!R6C2
'  ;about!_Contributor
'  ;=about!R4C2
'  ;about!_Copyright
'  ;=about!R5C2
'  ;about!_License
'  ;=about!R2C2
'  ;about!_LocalComment
'  ;=about!R1C2
'  ;about!_PublicName
'  ;=about!R7C2
'  ;about!_Since
'  ;=about!R10C2
'  ;about!_Tag
'  ;=about!R9C2
'  ;about!_Timestamp
'  ;=about!R8C2
'  ;about!_Url

'code
'  name;SSF
'{{{
Option Explicit
 
' シートの定義情報から、マクロ用の簡易ツールバーを生成する。
 
Private oApplication As Application
Private oThisWorkbook As Workbook
 
Private ButtonCaption As Variant
Private MyBar As Office.CommandBar
 
Friend Property Let BarButtons(ButtonArray As Variant)
    ButtonCaption = ButtonArray
End Property
 
Friend Property Get BarButtons() As Variant
    BarButtons = ButtonCaption
End Property
 
' ツールバーから直接呼ばれるメイン関数。
Friend Sub BarMain()
    Dim oAC As Object   ' 押されたボタンをもらう。
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    ' ユーザー定義のメインプロシジャにボタンを引き渡す。
    Main oAC
    Set oAC = Nothing
End Sub
 
' ボタンだけの簡易ツールバーを生成する。
' シートに定義されたボタン名称の数だけボタンを追加する。
Friend Function BarInitialize() As CommandBar
    Dim BAUDISTRIB As String
    Dim ButtonA As CommandBarButton
    Dim strBarName As String
    Dim strButtonName As String
    Dim strButtonDesc As String
    Dim i As Long, j As Long
    Dim strC() As Variant
    Dim Ra As Range
    
    On Error Resume Next
    
    Set oApplication = Application
    Set oThisWorkbook = ThisWorkbook
    ' 名前の衝突回避用に Url 情報を使う。
    BAUDISTRIB = " - " & Me.Parent.Range("_Url").Value
    BAUDISTRIB = " - " & Me.Range("_Url").Value
    ' シート上のボタン定義
    Set Ra = Me.Parent.Range("_ButtonCaption").CurrentRegion.Columns(1)
    Set Ra = Me.Range("_ButtonCaption").CurrentRegion.Columns(1)
    j = Ra.Cells.Count
    ReDim strC(0 To j - 1)
    For i = 0 To j - 1
        strC(i) = Array(Ra.Range("A1").Offset(i, 0).Text, Ra.Range("A1").Offset(i, 1).Text)
    Next
    ButtonCaption = strC
    Set Ra = Nothing
    strBarName = oThisWorkbook.Name & BAUDISTRIB
    Set MyBar = oApplication.CommandBars.Add(Name:=strBarName, Temporary:=True)
    For i = LBound(ButtonCaption) To UBound(ButtonCaption)
        strButtonName = CStr(ButtonCaption(i)(0))
        strButtonDesc = CStr(ButtonCaption(i)(1))
        Set ButtonA = MyBar.Controls.Add(Type:=1, Temporary:=True)
        With ButtonA
            .Style = msoButtonCaption
            .OnAction = Me.CodeName & ".BarMain"
            .Caption = strButtonName
            .Tag = strButtonName
            .TooltipText = strButtonDesc
            .BeginGroup = True
        End With
        Set ButtonA = Nothing
    Next
    MyBar.Visible = True
    MyBar.Position = msoBarTop
    Set BarInitialize = MyBar
End Function
 
' ツールバーを除去する。
Friend Sub BarTerminate()
    On Error Resume Next
    MyBar.Delete
    Set oApplication = Nothing
    Set oThisWorkbook = Nothing
End Sub
 
'}}}

'module
'  name;Module1
'{{{
Option Explicit
 
Public Sub Main(oAC As Object)
    Select Case oAC.Index
    Case 1
        Menu_ReferenceSheet oAC
    Case 2
        Menu_CharacterMarker oAC
    End Select
End Sub
 
Private Sub Menu_ReferenceSheet(oAC As Object)
    Dim Ws As Worksheet
    
    Set Ws = ActiveSheet
    Verify.SetReferenceSheet Ws, oAC
End Sub

Private Sub Menu_CharacterMarker(oAC As Object)
    Verify.SetDisableCharacterMarker oAC
End Sub
'}}}

'code
'  name;Verify
'{{{
Option Explicit

Private PrevCell As Range
Private RefSheet As Worksheet
Private RefButton As Office.CommandBarButton
Private EnableCharacterMarker As Boolean

' take a reference sheet for 1st input
Friend Sub SetReferenceSheet(Ws As Worksheet, Button As Object)
    If Ws Is Me Then
        If RefSheet Is Nothing Then Exit Sub
        If MayResetReference Then Exit Sub
        Set RefSheet = Nothing
    Else
        Set RefSheet = Ws
    End If
    
    Set RefButton = Button
    Reset Button.Parent
End Sub

' enable/disable character marker function
Friend Sub SetDisableCharacterMarker(Button As Object)
    EnableCharacterMarker = Not EnableCharacterMarker
    If EnableCharacterMarker Then
        Button.State = msoButtonDown
    Else
        Button.State = msoButtonUp
    End If
End Sub

' initialize
Friend Sub Reset(Bar As Object)
    Me.Activate
    Range("A1").Select
    Set PrevCell = ActiveCell
    SetupButton
    EnableCharacterMarker = False   ' opposite of the desired state
    SetDisableCharacterMarker Bar.Controls(2)
End Sub

' verify on input event
Private Sub Worksheet_Change(ByVal Target As Range)
    DoVerify Target
    Set PrevCell = Target
End Sub

' verify on cursor movement inter cells
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    DoVerify PrevCell
    Set PrevCell = Target.Cells(1)
End Sub

' core logic to verify
Private Sub DoVerify(Target As Range)
    On Error GoTo TrapAll
    
    Dim Ra As Range
    Dim SmallerTarget As Range
    
    If RefSheet Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    
    If Target.Cells.Count = 1 Then
        VerifySingleCell Target
    Else
        Set SmallerTarget = Intersect(Target, Union(Me.UsedRange, _
            Me.Range(RefSheet.UsedRange.Address(True, True, xlA1, False))))
        For Each Ra In SmallerTarget
            VerifySingleCell Ra
        Next
    End If
    Exit Sub
    
TrapAll:
    Debug.Print Err.Number, Err.Description
    Err.Clear
    SetupButton
End Sub

' verify a single cell
Private Sub VerifySingleCell(Target As Range)
    Dim RefTarget As Range
    Set RefTarget = Refered(Target)
    If IsSame(Target, RefTarget) Then
        RemoveAlarm Target, RefTarget
    Else
        ShowAlarm Target, RefTarget
    End If
    Set RefTarget = Nothing
End Sub

' get the reference cell against the 2nd
Private Function Refered(x As Range) As Range
    On Error GoTo LostSheet
    Set Refered = RefSheet.Range(x.Address(True, True, xlA1, False))
    Exit Function
    
LostSheet:
    Set Refered = x
    Set RefSheet = Nothing
    SetupButton
End Function

' activate alarm
Private Sub ShowAlarm(x As Range, y As Range)
    Const AlarmColor = 6
    If EnableCharacterMarker Then ShowMarker x, y
    If x.Interior.ColorIndex = AlarmColor Then Exit Sub
    x.Interior.ColorIndex = AlarmColor
End Sub

' deactivate alarm
Private Sub RemoveAlarm(x As Range, y As Range)
    If EnableCharacterMarker Then RemoveMarker x, y
    If x.Interior.ColorIndex = xlColorIndexNone Then Exit Sub
    x.Interior.ColorIndex = xlColorIndexNone
End Sub

' mark characters
Private Sub ShowMarker(x As Range, y As Range)
    Const AlarmColor = 3
    Const NormalColor = xlAutomatic
    Dim xText As String
    Dim yText As String
    Dim xLen As Long
    Dim yLen As Long
    Dim i As Long
    Dim iEnd As Long
    
    If x.HasFormula Then Exit Sub
    If TypeName(x.Value) <> "String" Then Exit Sub
    If TypeName(y.Value) <> "String" Then Exit Sub
    
    xText = x.Value
    yText = y.Value
    xLen = Len(xText)
    yLen = Len(yText)
    If xLen < yLen Then
        iEnd = xLen
    Else
        iEnd = yLen
    End If
    
    For i = 1 To iEnd
        If Not IsSameText(Mid(xText, i, 1), Mid(yText, i, 1)) Then Exit For
    Next
    
    ' mark x after i
    x.Characters.Font.ColorIndex = NormalColor
    If xLen < i Then Exit Sub
    x.Characters(i).Font.ColorIndex = AlarmColor
End Sub

' remove marks
Private Sub RemoveMarker(x As Range, y As Range)
    Const NormalColor = xlAutomatic
    If x.HasFormula Then Exit Sub
    If TypeName(x.Value) <> "String" Then Exit Sub
    x.Characters.Font.ColorIndex = NormalColor
End Sub

' update toolbar button for reference sheet
Private Sub SetupButton()
    On Error GoTo ButtonFailure
    
    If IsAlive(RefSheet) Then
        With RefButton
            .Caption = "[" & RefSheet.Parent.Name & "]" & RefSheet.Name
            .State = msoButtonDown
        End With
    Else
        With RefButton
            .Caption = "参照シート無し"
            .State = msoButtonUp
        End With
    End If
    Exit Sub
    
ButtonFailure:
    Debug.Print Err.Number, Err.Description
End Sub

' tell my reference sheet is void or not
Private Function IsAlive(Ws As Worksheet) As Boolean
    If Ws Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim i As Long
    i = Ws.Index
    If Err.Number = 0 Then IsAlive = True
End Function

' confirm to remove reference
Private Function MayResetReference() As Boolean
    MayResetReference = _
        (MsgBox("参照シートの設定を解除しますか？", vbOKCancel, _
            "検証シートがアクティブなときに参照ボタンを押すと、設定を解除します。") _
        = vbCancel)
End Function

' validate a cell
Private Function IsSame(x As Range, y As Range) As Boolean
    IsSame = IsSameText(x.Value, y.Value)
End Function

' valid condition
' --- change this to get a custom result on verification rules.
Private Function IsSameText(x As String, y As String) As Boolean
    IsSameText = (x = y)
End Function
'}}}

'code
'  name;ThisWorkbook
'{{{
Option Explicit

' ツールバー初期化と終了
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SSF.BarTerminate
End Sub
 
Private Sub Workbook_Open()
    SSF.BarInitialize
    ThisWorkbook.Saved = True
End Sub
'}}}


```


# Snapshots #

  * Excel 2000
> > ![http://lh5.googleusercontent.com/-HDASK5t9IlY/TWtVwxlG6SI/AAAAAAAAADY/mefO-HtpA3g/s1600/shot1.png](http://lh5.googleusercontent.com/-HDASK5t9IlY/TWtVwxlG6SI/AAAAAAAAADY/mefO-HtpA3g/s1600/shot1.png)
  * Excel 2007
> > ![http://lh4.googleusercontent.com/-5DjjseLBINU/TWtVxgB-8cI/AAAAAAAAADc/x9nlHi1Njag/s1600/shot2.png](http://lh4.googleusercontent.com/-5DjjseLBINU/TWtVxgB-8cI/AAAAAAAAADc/x9nlHi1Njag/s1600/shot2.png)