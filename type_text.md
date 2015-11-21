

# Introduction #

  * work with verious character encodings in vbscript

## 概要 ##
  * VBScript で、様々な文字コードに対応する

# Details #

![http://4.bp.blogspot.com/-T4UVGaHdipY/TVjYMPBaI5I/AAAAAAAAADI/p-PwKZnF9tc/s1600/shot1.png](http://4.bp.blogspot.com/-T4UVGaHdipY/TVjYMPBaI5I/AAAAAAAAADI/p-PwKZnF9tc/s1600/shot1.png)
![http://2.bp.blogspot.com/-dBM2PNA6PgQ/TVjYMhARxCI/AAAAAAAAADM/ME3TKWn9zw8/s1600/shot2.png](http://2.bp.blogspot.com/-dBM2PNA6PgQ/TVjYMhARxCI/AAAAAAAAADM/ME3TKWn9zw8/s1600/shot2.png)

  * a command to convert an encoding of text files, for Japanese or other non-US languages.
  * converts an euc-jp text file into a unicode text.
  * converts a unicode text into an euc-jp text file.
  * save a utf-8 text file without BOM.
  * no preconditions are required for almost Windows, because it's all written by `VBScript` .
  * and free to customize.

## 説明 ##
  * コマンドラインで、日本語などの文字エンコードを変換する。
  * euc-jp などでエンコードしたテキストファイルを、ウィンドウズの Unicode に変換して読む。
  * ウィンドウズの Unicode テキストファイルを、 euc-jp などエンコードを指定して保存する。
  * BOM 無しの utf-8 ファイルを保存する。
  * `VBScript` だけで書かれているので、ほとんどのウィンドウズ環境でそのまま動く。
  * 好き勝手にカスタマイズできる。

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=type_text)

# How to use #

  1. open a Windows Command Prompt. (cmd.exe)
  1. `CScript //NoLogo type.vbs /e:iso-2022-jp jis.txt`
  1. `CScript //NoLogo text.vbs /e:euc-jp < sjis.txt euc.txt`

## 使い方 ##
  1. コマンドプロンプトを開く。
  1. `CScript //NoLogo type.vbs /e:iso-2022-jp jis.txt`
  1. `CScript //NoLogo text.vbs /e:euc-jp < sjis.txt euc.txt`


# Code #

### type.vbs ###

```
' type
' print encoded files to stdout.
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo type.vbs /e:Charset FILE [MORE FILES]

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Named, Args.Unnamed
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Opts, Files)
    Dim File, Charset
    Dim StdOut, ts
    Dim R
    Const adTypeText = 2
    
    Set StdOut = WScript.StdOut
    Set ts = CreateObject("ADODB.Stream")
    Charset = Opts("e")
    If Charset = "" Then Charset = "utf-8"
    Set R = RegLineFeed
    
    For Each File in Files
        ts.Open
        ts.Type = adTypeText
        ts.Charset = Charset
        ts.LoadFromFile File
        StdOut.Write R.Replace(ts.ReadText, vbCrLf)
        ts.Close
    Next
    
    Set R = Nothing
    Set ts = Nothing
    Set StdOut = Nothing
End Sub

Function RegLineFeed()
    Dim R
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = True
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "\r?\n"
    
    Set RegLineFeed= R
End Function
```

### text.vbs ###

```
' text
' save as an encoded file from stdin.
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo text.vbs /e:Charset FILE

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Named, Args.Unnamed(0)
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Opts, File)
    Dim Charset
    Dim StdIn
    
    Set StdIn = WScript.StdIn
    Charset = Opts("e")
    If Charset = "" Then Charset = "utf-8"
    
    If LCase(Charset) = "utf-8" Then
        SaveTextNoBom StdIn, File, Charset
    Else
        SaveText StdIn, File, Charset
    End If
    
    StdIn.Close
    Set StdIn = Nothing
End Sub

Private Sub SaveText(inSt, File, Charset)
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    Dim outSt
    
    Set outSt = CreateObject("ADODB.Stream")
    outSt.Open
    outSt.Type = adTypeText
    outSt.Charset = Charset
    outSt.WriteText inSt.ReadAll
    outSt.SaveToFile File, adSaveCreateOverWrite
    outSt.Close
    Set outSt = Nothing
End Sub

Private Sub SaveTextNoBom(inSt, File, Charset)
    ' expect Charset="utf-8"
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    Dim buff, outSt
    
    Set outSt = CreateObject("ADODB.Stream")
    Set buff = CreateObject("ADODB.Stream")
    buff.Open
    buff.Type = adTypeText
    buff.Charset = Charset
    buff.WriteText inSt.ReadAll
    
    ' skip 3 bytes BOM
    buff.Position = 0
    buff.Type = adTypeBinary
    buff.Position = 3
    
    ' save as binary
    outSt.Open
    outSt.Type = adTypeBinary
    outSt.Write buff.Read
    buff.Close
    outSt.SaveToFile File, adSaveCreateOverWrite
    
    outSt.Close
    Set buff = Nothing
    Set outSt = Nothing
End Sub
```

### xtype.bat ###

```
@echo off
cscript //nologo C:\tmp\type.vbs %*
```

### xtext.bat ###

```
@echo off
cscript //nologo C:\tmp\text.vbs %*
```

### textc.vbs ###

```
' textc
' convert an encoded file to another encoding
' Copyright (C) 2011 Tomizono - movvba.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo textc.vbs /in:CharsetFrom /out:CharsetTo /n:lf FILE_FROM FILE_TO

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Named, Args.Unnamed(0), Args.Unnamed(1)
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Opts, FileFrom, FileTo)
    Dim CharsetIn, CharsetOut, LineFeed
    Dim StreamIn
    
    CharsetIn = Opts("in")
    If CharsetIn = "" Then CharsetIn = "utf-8"
    CharsetOut = Opts("out")
    If CharsetOut = "" Then CharsetOut = "utf-8"
    LineFeed = Opts("n")
    
    Set StreamIn = OpenText(FileFrom, CharsetIn)
    
    If LCase(CharsetOut) = "utf-8" Then
        SaveTextNoBom StreamIn, FileTo, CharsetOut, LineFeed
    Else
        SaveText StreamIn, FileTo, CharsetOut, LineFeed
    End If
    
    StreamIn.Close
    Set StreamIn = Nothing
End Sub

Private Function OpenText(File, Charset)
    Const adTypeText = 2
    Dim ts
    
    Set ts = CreateObject("ADODB.Stream")
    ts.Open
    ts.Type = adTypeText
    ts.Charset = Charset
    ts.LoadFromFile File
    Set OpenText = ts
End Function

Private Sub SaveText(inSt, File, Charset, LineFeed)
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    Dim outSt
    
    Set outSt = CreateObject("ADODB.Stream")
    outSt.Open
    outSt.Type = adTypeText
    outSt.Charset = Charset
    outSt.WriteText ConvertLineFeed(inSt.ReadText, LineFeed)
    outSt.SaveToFile File, adSaveCreateOverWrite
    outSt.Close
    Set outSt = Nothing
End Sub

Private Sub SaveTextNoBom(inSt, File, Charset, LineFeed)
    ' expect Charset="utf-8"
    Const adTypeBinary = 1
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    Dim buff, outSt
    
    Set outSt = CreateObject("ADODB.Stream")
    Set buff = CreateObject("ADODB.Stream")
    buff.Open
    buff.Type = adTypeText
    buff.Charset = Charset
    buff.WriteText ConvertLineFeed(inSt.ReadText, LineFeed)
    
    ' skip 3 bytes BOM
    buff.Position = 0
    buff.Type = adTypeBinary
    buff.Position = 3
    
    ' save as binary
    outSt.Open
    outSt.Type = adTypeBinary
    outSt.Write buff.Read
    buff.Close
    outSt.SaveToFile File, adSaveCreateOverWrite
    
    outSt.Close
    Set buff = Nothing
    Set outSt = Nothing
End Sub

Private Function ConvertLineFeed(Text, NewLineFeed)
  If NewLineFeed = "" Then
    ConvertLineFeed = Text
  Else
    Dim LineFeed
    Select Case LCase(NewLineFeed)
    Case "br"
      LineFeed = "<br/>"
    Case "c"
      LineFeed = "\n"
    Case "lf"
      LineFeed = vbLf
    Case "cr"
      LineFeed = vbCr
    Case "crlf"
      LineFeed = vbCrLf
    Case Else
      LineFeed = vbCrLf
    End Select
    ConvertLineFeed = RegLineFeed.Replace(Text, LineFeed)
  End If
End Function

Function RegLineFeed()
    Dim R
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = True
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "\r?\n"
    
    Set RegLineFeed= R
End Function
```