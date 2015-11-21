# Introduction #

  * how to use the timer to make a background macro in vba

## 概要 ##
  * VBAのタイマーを使って、マクロをバックグラウンドで動かす

# Details #

  * we use the built-in timer in the excel application.
  * the task runs when the Sheet1 is active and the breaker switch is ON
  * the task stops when the Sheet2 is active, so you can do any other things without disturbed by the macro
  * when the Sheet1 is activated again, the task continues automatically
  * note this is not a real background task, just doing a time sharing, because the excel VBA has only a single thread by design

## 説明 ##
  * エクセル Application の内蔵タイマーを使う
  * Sheet1 がアクティブで、スイッチが ON のときにタスクが動く
  * Sheet2 がアクティブになるとタスクは止まるので、マクロに邪魔されず他の作業ができる
  * Sheet1 がもう一度アクティブになると、タスクは勝手に再開する
  * これは本物のバックグラウンドではない。時間を切り分けているだけだ。エクセル VBA は仕様上、単一スレッドでしか動かないから

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. activate Sheet1, double click the A1 cell, this is a breaker switch.
  1. activate Sheet2, back to the Sheet1 after a few moment, watch the work
  1. the task is simple, to write down in the column C

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. Sheet1 をアクティブにして、 A1 セルをダブルクリックする。これがスイッチ
  1. Sheet2 をアクティブにして、しばらくして Sheet1 に戻り、作業の様子を見る
  1. タスクは単純で、 C 列を下に記入していくだけのもの

# Snapshots #

![http://3.bp.blogspot.com/_EUW0nrj9XlM/TRljNQR9-nI/AAAAAAAAAAs/GbDnYc4-sBg/s1600/shot1.png](http://3.bp.blogspot.com/_EUW0nrj9XlM/TRljNQR9-nI/AAAAAAAAAAs/GbDnYc4-sBg/s1600/shot1.png)
![http://2.bp.blogspot.com/_EUW0nrj9XlM/TRljNq6om7I/AAAAAAAAAAw/QQNZ3vis9S4/s1600/shot2.png](http://2.bp.blogspot.com/_EUW0nrj9XlM/TRljNq6om7I/AAAAAAAAAAw/QQNZ3vis9S4/s1600/shot2.png)

# Code #

```
'workbook
'  name;hello_timer.xls

'worksheet
'  name;Sheet2

'worksheet
'  name;Sheet1

'cells-formula
'  address;A1:A1
'         ;ON

'code
'  name;Sheet1
'{{{
Option Explicit

Const RepeatAfter = "0:00:02"
Const MaxDelay = "0:00:10"

Private Fire As Boolean
Private Alive As Boolean
Private Submitted As Variant

Friend Sub Continue()
    Dim TheTime As Date
    Dim TheDelay As Date
    Dim TheProc As String
    
    If Not (Alive And Fire) Then Exit Sub
    Task
    
    TheTime = Now() + TimeValue(RepeatAfter)
    TheDelay = TheTime + TimeValue(MaxDelay)
    TheProc = Me.CodeName & ".Continue"
    Submitted = Array(TheTime, TheProc, TheDelay)
    Application.OnTime TheTime, TheProc, TheDelay
End Sub

Private Sub RemoveTimer()
    If Not IsArray(Submitted) Then Exit Sub
    On Error Resume Next
    Application.OnTime Submitted(0), Submitted(1), Submitted(2), False
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Address(False, False, xlA1, False) <> "A1" Then Exit Sub
    
    Alive = Not Alive
    Target.Interior.ColorIndex = IIf(Alive, 3, 1)
    If Alive Then
        Fire = True
        Continue
    Else
        RemoveTimer
    End If
    
    Cancel = True
End Sub

Private Sub Worksheet_Activate()
    Fire = True
    Continue
End Sub

Private Sub Worksheet_Deactivate()
    Fire = False
    RemoveTimer
End Sub

' This is a small fragment of to do
Private Sub Task()
    Static RememberRow As Long
    
    RememberRow = RememberRow + 1
    Me.Cells(RememberRow, 3).Value = Now()
End Sub
'}}}

```
