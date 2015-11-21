

# Introduction #

  * work with binary dump in vbscript

## 概要 ##
  * VBScript で、バイナリダンプに関わる

# Details #

![http://3.bp.blogspot.com/-CB73KAs_ZbE/TVjWZSKbwHI/AAAAAAAAAC8/vItF6N31JZM/s1600/shot1.png](http://3.bp.blogspot.com/-CB73KAs_ZbE/TVjWZSKbwHI/AAAAAAAAAC8/vItF6N31JZM/s1600/shot1.png)
![http://1.bp.blogspot.com/-CF39aumQUxk/TVjWZraarmI/AAAAAAAAADA/4pzGgolYJqo/s1600/shot2.png](http://1.bp.blogspot.com/-CF39aumQUxk/TVjWZraarmI/AAAAAAAAADA/4pzGgolYJqo/s1600/shot2.png)
![http://4.bp.blogspot.com/-QnRd6Eh07QQ/TVjWaPd-BsI/AAAAAAAAADE/IDS_Rhegs-w/s1600/shot3.png](http://4.bp.blogspot.com/-QnRd6Eh07QQ/TVjWaPd-BsI/AAAAAAAAADE/IDS_Rhegs-w/s1600/shot3.png)

  * a command to exchange between a binary file and a dump text.
  * converts a binary file into a hexagonal dump or a bit dump text.
  * converts a hexagonal dump text into a binary file.
  * 1 bit binary dump text will be some artistic works, maybe.
  * no preconditions are required for almost Windows, because it's all written by `VBScript` .
  * and free to customize.

## 説明 ##
  * コマンドラインで、バイナリダンプとダンプテキストからのバイナリ生成を行う。
  * バイナリファイルを１６進や２進にダンプし、テキストに変換する。
  * １６進ダンプテキストからバイナリファイルを生成する。
  * ２進ダンプを使ってアスキーアート芸術の創作活動をする。
  * `VBScript` だけで書かれているので、ほとんどのウィンドウズ環境でそのまま動く。
  * 好き勝手にカスタマイズできる。

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=type_bin)

# How to use #

  1. open a Windows Command Prompt. (cmd.exe)
  1. `CScript //NoLogo typebin.vbs /c:444 shop.png > shop.txt`
  1. `CScript //NoLogo dechex.vbs < shop.txt another.png`

## 使い方 ##
  1. コマンドプロンプトを開く。
  1. `CScript //NoLogo typebin.vbs /c:444 shop.png > shop.txt`
  1. `CScript //NoLogo dechex.vbs < shop.txt another.png`


# Code #

### type\_bin.vbs ###

```
' typebin
' print binary files to stdout.
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo typebin.vbs /o FILE [MORE FILES]

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Named, Args.Unnamed
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Opts, Files)
    Dim File
    Dim vChars, vWrap, PrintAs, vSkip
    Dim bLowbitFirst, bC0, bC1
    Dim StdOut, ts
    Const adTypeBinary = 1
    Const adTypeText = 2
    
    'Set StdOut = WScript.StdOut
    Set StdOut = New StringStream
    Set ts = CreateObject("ADODB.Stream")
    
    vChars = 128
    If Opts("c") <> "" Then vChars = CLng(Opts("c"))
    vWrap = 16
    If Opts("w") <> "" Then vWrap = CLng(Opts("w"))
    vSkip = 0
    If Opts("s") <> "" Then vSkip = CLng(Opts("s"))
    bLowbitFirst = True
    If Opts("bit") <> "" Then bLowbitFirst = (UCase(Left(Opts("bit"), 1)) = "L")
    bC0 = "0"
    If Opts("bit0") <> "" Then bC0 = Opts("bit0")
    bC1 = "1"
    If Opts("bit1") <> "" Then bC1 = Opts("bit1")
    
    PrintAs = "x"
    For Each x In Array("d", "o", "b")
        If Opts.Exists(x) Then PrintAs = x
    Next
    
    For Each File in Files
        If Files.Count > 1 Then StdOut.Write File & vbCrLf
        ts.Open
        ts.Type = adTypeBinary
        ts.LoadFromFile File
        ts.Position = vSkip
        
        Select Case PrintAs
        Case "x"
            EncHexDelimited ts, StdOut, " ", vChars, vWrap
        Case "d"
            EncDecDelimited ts, StdOut, " ", vChars, vWrap
        Case "o"
            EncOctDelimited ts, StdOut, " ", vChars, vWrap
        Case "b"
            EncBitDelimited ts, StdOut, "", vChars, vWrap, bLowbitFirst, bC0, bC1
        End Select
        ts.Close
        StdOut.Write vbCrLf
    Next
    
    Set ts = Nothing
    WScript.Echo StdOut.Text
    Set StdOut = Nothing
End Sub

Private Sub EncHexDelimited(inSt, outSt, Delimiter, Limit, Wrap)
    Dim Counter, bs
    Counter = 0
    Do Until inSt.EOS
        If Limit = Counter Then Exit Do
        If Counter Mod Wrap = 0 And Counter > 0 Then outSt.Write vbCrLf
        bs = inSt.Read(1)
        outSt.Write Right("00" & LCase(Hex(AscB(bs))), 2) & Delimiter
        Counter = Counter + 1
    Loop
End Sub

Private Sub EncDecDelimited(inSt, outSt, Delimiter, Limit, Wrap)
    Dim Counter, bs
    Counter = 0
    Do Until inSt.EOS
        If Limit = Counter Then Exit Do
        If Counter Mod Wrap = 0 And Counter > 0 Then outSt.Write vbCrLf
        bs = inSt.Read(1)
        outSt.Write Right("   " & AscB(bs), 3) & Delimiter
        Counter = Counter + 1
    Loop
End Sub

Private Sub EncOctDelimited(inSt, outSt, Delimiter, Limit, Wrap)
    Dim Counter, bs, hb, lb, b
    Counter = 0
    Do Until inSt.EOS
        If Limit = Counter Then Exit Do
        If Counter Mod Wrap = 0 And Counter > 0 Then outSt.Write vbCrLf
        bs = inSt.Read(1)
        b = AscB(bs)
        lb = b Mod 16
        hb = Int(b / 16)
        outSt.Write Right("  " & lb, 2) & Delimiter & Right("  " & hb, 2) & Delimiter
        Counter = Counter + 1
    Loop
End Sub

Private Sub EncBitDelimited(inSt, outSt, Delimiter, Limit, Wrap, LowbitFirst, C0, C1)
    Dim Counter, bs, x, j
    Dim CStrB
    Dim Bits(7)
    CStrB = Array(C0, C1)
    
    Counter = 0
    Do Until inSt.EOS
        If Limit = Counter Then Exit Do
        If Counter Mod Wrap = 0 And Counter > 0 Then outSt.Write vbCrLf
        
        bs = inSt.Read(1)
        x = AscB(bs)
        For j = 0 To 7
            Bits(j) = x Mod 2
            x = Int(x / 2)
        Next
        If LowbitFirst Then
            For j = 0 To 7
                outSt.Write CStrB(Bits(j)) & Delimiter
            Next
        Else
            For j = 7 To 0 Step -1
                outSt.Write CStrB(Bits(j)) & Delimiter
            Next
        End If
        
        Counter = Counter + 1
    Loop
End Sub

Class StringStream
    Public Text
    
    Public Sub Write(Data)
        Text = Text & Data
    End Sub
End Class
```

### dechex.vbs ###

```
' dechex
' save a binary file from a hex dump text
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo dechex.vbs < [TEXT_FILE] BINARY_FILE

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Unnamed(0)
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(File)
    Dim StdIn, ts
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    
    Set StdIn = WScript.StdIn
    Set ts = CreateObject("ADODB.Stream")
    
    ts.Open
    ts.Type = adTypeBinary
    DecHexDelimited StdIn, ts
    ts.SaveToFile File, adSaveCreateOverWrite
    ts.Close
    
    StdIn.Close
    Set ts = Nothing
    Set StdIn = Nothing
End Sub

Private Sub DecHexDelimited(inSt, outSt)
    Dim buff
    Const adTypeBinary = 1
    Const adTypeText = 2
    Dim R, M
    Dim ByteData, bs
    
    bs = Replace(inSt.ReadAll, vbLf, " ")
    ' work arround a bug that vb regexp fails in multi-lilne searching
    Set buff = CreateObject("ADODB.Stream")
    buff.Open
    ' binary writing fails in scripting by design,
    ' so we use this unicode buff stream as a proxy
    Set R = RegNextDelimiter
    
    Do Until bs = ""
        Set M = R.Execute(bs)
        If M.Count = 0 Then Exit Do
        ByteData = M(0).SubMatches(0)
        bs = M(0).SubMatches(1)
        buff.Type = adTypeText
        buff.Charset = "unicode"
        'buff.WriteText  ChrB(CByte("&H" & ByteData)) & ChrB(0)
        buff.WriteText  ChrW(CByte("&H" & ByteData))
        buff.Position = 0
        buff.Type = adTypeBinary
        buff.Position = 2   ' skip bom
        outSt.Write buff.Read(1)
        buff.Position = 0
    Loop
    
    buff.Close
    Set R = Nothing
    Set buff = Nothing
End Sub

Private Function RegNextDelimiter()
    Dim R
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = False
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = "([0-9A-Za-z]+)[^0-9A-Za-z]*(.*)"
    
    Set RegNextDelimiter = R
End Function
```

### ge-jitu.bat ###

```
@echo off
cscript //nologo C:\tmp\typebin.vbs /c:16 /b /w:4 /bit0:╋ /bit1:┓ %*
rem cscript //nologo C:\tmp\typebin.vbs /c:12 /b /w:4 /bit0:┛ /bit1:┓ %*
cscript //nologo C:\tmp\typebin.vbs /c:16 /b /w:4 /bit0:┻ /bit1:┫ %*
cscript //nologo C:\tmp\typebin.vbs /c:8 /b /w:4 /bit0:● /bit1:○ %*
cscript //nologo C:\tmp\typebin.vbs /c:8 /b /w:4 /bit0:ね /bit1:れ %*
rem cscript //nologo C:\tmp\typebin.vbs /c:8 /b /w:4 /bit0:さ /bit1:ち %*
rem cscript //nologo C:\tmp\typebin.vbs /c:8 /b /w:4 /bit0:し /bit1:じ %*
rem cscript //nologo C:\tmp\typebin.vbs /c:4 /b /w:4 /bit0:Ｒ /bit1:Я %*
cscript //nologo C:\tmp\typebin.vbs /c:16 /b /w:8 /bit0:" " /bit1:* %*
cscript //nologo C:\tmp\typebin.vbs /c:16 /b /w:8 /bit0:">" /bit1:"<" %*
```