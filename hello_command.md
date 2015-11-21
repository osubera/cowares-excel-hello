# Introduction #

  * how to execute a windows shell command in vba

## 概要 ##
  * VBAでウィンドウズのコマンドやプログラムを実行する

# Details #

  * we can friendly use a commandl line programs by WshShell object.
  * the last example shows that any other script languages are now ready to use to improve our vba macro.

## 説明 ##
  * WshShell オブジェクトを使うと、コマンドプロンプトで使うプログラムを容易に扱える。
  * 最後の例で示した方法を使えば、他のどんなスクリプト言語でも VBA マクロを強化する目的で組み込める。

# How to use #

  1. run a macro `test_*`
  1. we have 5 tests here.
  1. you need to close the notepad by yourself to continue the macro, this is an example to wait the application quit.
  1. the last one uses an external `SORT` command as a string sort function in vba.

## 使い方 ##
  1. マクロ `test_*` を実行する。
  1. ５つの例が載っている。
  1. 最初のはアプリケーションの終了を待つサンプルなので、メモ帳を手で終了させないと先に進まない。
  1. 最後のは、外部の `SORT` コマンドを、 VBA での文字列ソート関数のように使っている。

# Code #

```
'workbook
'  name;hello_command.xls

'require
'  ;{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B} 1 0 Windows Script Host Object Model

'worksheet
'  name;Sheet1


'module
'  name;Module1
'{{{
Option Explicit

' talk with command shell

Sub test_ReturnCode()
    Dim wsh As IWshRuntimeLibrary.WshShell
    Dim ReturnCode As Long
    
    Set wsh = New WshShell
    
    ReturnCode = wsh.Run("notepad.exe", WshNormalFocus, True)
    Debug.Assert ReturnCode = 0
    
    ReturnCode = wsh.Run("cmd.exe /C dir", WshNormalNoFocus, True)
    Debug.Assert ReturnCode = 0
    
    ReturnCode = wsh.Run("cmd.exe /C dir WE_DO_NOT_HAVE_SUCH_FILES.", WshHide, True)
    Debug.Assert ReturnCode = 1
    
    Set wsh = Nothing
End Sub

Sub test_SendKey()
    Dim wsh As IWshRuntimeLibrary.WshShell
    Dim con As IWshRuntimeLibrary.WshExec
    Dim i As Long
    
    Set wsh = New WshShell
    Set con = wsh.Exec("notepad.exe")
    
    For i = 1 To 5
        Debug.Print con.Status, con.ProcessId
        wsh.SendKeys "Hello Notepad.{ENTER}"
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Next
    
    ' send key is danger,
    ' because if the Excel is active when sending the following keys,
    ' the Excel itself quits unexpectedly.
    'wsh.SendKeys "%{F4}n"
    con.Terminate
    
    Do While con.Status = WshRunning
        DoEvents
    Loop
    Debug.Assert con.Status = WshFinished
    Debug.Assert con.ExitCode = 0
    
    Set con = Nothing
    Set wsh = Nothing
End Sub

Sub test_StdOut()
    Dim wsh As IWshRuntimeLibrary.WshShell
    Dim con As IWshRuntimeLibrary.WshExec
    
    Set wsh = New WshShell
    Set con = wsh.Exec("cmd.exe /C ver")
    
    Do While con.Status = WshRunning
        Debug.Print con.Status, con.ProcessId
        DoEvents
    Loop
    Debug.Print con.Status, con.ProcessId
    
    If Not con.StdErr.AtEndOfStream Then
        Debug.Print con.StdErr.ReadAll
    Else
        Debug.Print "No Errors"
    End If
    
    Debug.Assert con.ExitCode = 0
    
    Do Until con.StdOut.AtEndOfStream
        Debug.Print con.StdOut.ReadLine
    Loop
    
    Set con = Nothing
    Set wsh = Nothing
End Sub

Sub test_StdErr()
    Dim wsh As IWshRuntimeLibrary.WshShell
    Dim con As IWshRuntimeLibrary.WshExec
    
    Set wsh = New WshShell
    Set con = wsh.Exec("cmd.exe /C ls")
    
    Do While con.Status = WshRunning
        Debug.Print con.Status, con.ProcessId
        DoEvents
    Loop
    Debug.Print con.Status, con.ProcessId
    
    If Not con.StdErr.AtEndOfStream Then
        Debug.Print con.StdErr.ReadAll
    Else
        Debug.Print "No Errors"
    End If
    
    Debug.Assert con.ExitCode = 1
    
    Do Until con.StdOut.AtEndOfStream
        Debug.Print con.StdOut.ReadLine
    Loop
    
    Set con = Nothing
    Set wsh = Nothing
End Sub

Sub test_StdIn_StdOut()
    Dim wsh As IWshRuntimeLibrary.WshShell
    Dim con As IWshRuntimeLibrary.WshExec
    Dim v As Variant
    Dim w As Variant
    
    Set wsh = New WshShell
    Set con = wsh.Exec("cmd.exe /C sort")
    
    w = Array("Hello", "Console", "World", "!", "hello", "console", "world", ".")
    For Each v In w
        con.StdIn.WriteLine v
    Next
    con.StdIn.Close
    
    Do While con.Status = WshRunning
        Debug.Print con.Status, con.ProcessId
        DoEvents
    Loop
    Debug.Print con.Status, con.ProcessId
    
    If Not con.StdErr.AtEndOfStream Then
        Debug.Print con.StdErr.ReadAll
    Else
        Debug.Print "No Errors"
    End If
    
    Debug.Assert con.ExitCode = 0
    
    Do Until con.StdOut.AtEndOfStream
        Debug.Print con.StdOut.ReadLine
    Loop
    
    Set con = Nothing
    Set wsh = Nothing
End Sub
'}}}


```

# Results #

```
test_StdOut()

 0             740 
 0             740 
 1             740 
No Errors

Microsoft Windows 2000 [Version 5.00.2195]


test_StdErr()

 0             828 
 0             828 
 1             828 
'ls' は、内部コマンドまたは外部コマンド、
操作可能なプログラムまたはバッチ ファイルとして認識されていません。


test_StdIn_StdOut()

 0             828 
 0             828 
 0             828 
 1             828 
No Errors
!
.
console
Console
hello
Hello
World
world
```