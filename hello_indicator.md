# Introduction #

  * how to use an active indicator informing the progress in vba

## 概要 ##
  * VBAで進行状況をライブ表示する

# Details #

  * vba cannot create another thread to work as a dynamic indicator by the design.
  * here are 2 ideas to make a pseudo live and interructive indicator those show the progress and give a stop function.

## 説明 ##
  * VBA では仕様により、動的な情報表示に使えるような別スレッドを作ることができない。
  * ここでは、２つのサンプルを示す。いずれも、プログレスを動的に示しす機能と、停止ボタンのような割り込み機能を、擬似的に提供する。

# How to use #

  1. run a macro `test_ToolBarIndicator`
    1. a tool bar appears to show the progress, and disappears on done.
    1. press a `stop` button to stop the macro interructively.
  1. run a macro `test_ShapeIndicator`
    1. a rectangle shape appears to show the progress, and disappears on done.
    1. press the rectangle to stop the macro interructively.

## 使い方 ##
  1. マクロ `test_ToolBarIndicator` を実行する
    1. ツールバーが現れて、進行状況を表示し、終了すると消える。
    1. マクロを途中で中断するときは、 `stop` ボタンを押す。
  1. マクロ `test_ShapeIndicator` を実行する
    1. 四角形が現れて、進行状況を表示し、終了すると消える。
    1. マクロを途中で中断するときは、その四角形を押す。

# Code #

```
'workbook
'  name;hello_indicator.xls


'module
'  name;Module1
'{{{
Option Explicit

Private TestDone As Boolean

Public Sub test_ToolBarIndicator()
    Dim St As Office.CommandBar
    Dim i As Long
    Const iStart = 0
    Const iEnd = 100
    
    Set St = CreateToolBar
    
    TestDone = False
    St.Controls(1).Enabled = False
    St.Controls(2).Caption = iStart
    St.Controls(2).State = msoButtonDown
    St.Controls(3).Caption = iEnd
    St.Controls(3).State = msoButtonDown
    
    For i = iStart To iEnd
        HeavyTask
        
        St.Controls(2).Caption = i
        St.Controls(3).Caption = iEnd - i
        
        DoEvents
        If TestDone Then Exit For
    Next
    
    RemoveToolBar St
    Set St = Nothing
End Sub

Public Sub test_ShapeIndicator()
    Dim St As Shape
    Dim Ws As Worksheet
    Dim i As Long
    Const iStart = 0
    Const iEnd = 100
    
    Set Ws = ActiveSheet
    Set St = Ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 280, 100)
    St.OnAction = "test_Stop"
    St.Fill.ForeColor.SchemeColor = 13
    St.TextFrame.Characters.Text = ""
    St.TextFrame.Characters.Font.ColorIndex = 3
    St.Shadow.Type = msoShadow6
    
    TestDone = False
    
    For i = iStart To iEnd
        HeavyTask
        
        St.TextFrame.Characters.Text = "実行中 " & i & vbLf & "残り  " & iEnd - i & vbLf & "Click to Stop"
        
        DoEvents
        If TestDone Then Exit For
    Next
    
    St.Delete
    Set St = Nothing
    Set Ws = Nothing
End Sub

Public Sub test_Stop()
    TestDone = True
End Sub

Private Sub HeavyTask()
    Application.Wait Now() + Rnd(2) / 24 / 60 / 60
End Sub

Private Function CreateToolBar() As Office.CommandBar
    Dim ButtonA As CommandBarButton
    Dim ButtonCaption As Variant
    Dim i As Long
    Dim MyBar As Office.CommandBar
    
    ButtonCaption = Array("disable", "inc", "dec", "stop")
    Set MyBar = Application.CommandBars.Add(Name:="test_hello_indicator", Temporary:=True)
    
    For i = LBound(ButtonCaption) To UBound(ButtonCaption)
        Set ButtonA = MyBar.Controls.Add(Type:=1, Temporary:=True)
        With ButtonA
            .Style = msoButtonCaption
            .OnAction = "test_Stop"
            .Caption = ButtonCaption(i)
            .BeginGroup = True
        End With
        Set ButtonA = Nothing
    Next
    MyBar.Visible = True
    MyBar.Position = msoBarTop
    Set CreateToolBar = MyBar
End Function

Private Sub RemoveToolBar(Bar As Office.CommandBar)
     Bar.Delete
End Sub
'}}}



```