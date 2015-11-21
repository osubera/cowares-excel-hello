# Introduction #

  * how to use an windows clipboard in vba

## 概要 ##
  * VBAでウィンドウズのクリップボードを使う

# Details #

  * use Microsoft Forms 2.0 Object Library

## 説明 ##
  * Microsoft Forms 2.0 Object Library を使う

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. sub test1() gets text from the clipboard.
  1. sub test2() puts text to the clipboard.

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. sub test1() でテキストをクリップボードから取得する。
  1. sub test1() でテキストをクリップボードに転送する。

# Code #

```
'workbook
'  name;hello_clipboard.xls

'require
'  ;{0D452EE1-E08F-101A-852E-02608C4D0BB4} 2 0 Microsoft Forms 2.0 Object Library

'module
'  name;Module1
'{{{
Option Explicit

' クリップボードからテキストを取得する。
Private Function CopyFromClipboard() As String
    Const CFText As Long = 1
    Dim Text As String
    Dim Clip As MSForms.DataObject
    Set Clip = New MSForms.DataObject
    Clip.GetFromClipboard
    If Clip.GetFormat(CFText) Then
        Text = Clip.GetText()
    Else
        Text = ""
    End If
    CopyFromClipboard = Text
End Function
 
' クリップボードにテキストを格納する。
Private Sub CopyToClipboard(Text As String)
    Dim Clip As MSForms.DataObject
    Set Clip = New MSForms.DataObject
    Clip.SetText Text
    Clip.PutInClipboard
End Sub

Sub test1()
    Debug.Print CopyFromClipboard
End Sub

Sub test2()
    CopyToClipboard "Hello Excel"
End Sub
'}}}

```
