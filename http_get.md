

# Introduction #

  * download web pages in command prompt

## 概要 ##
  * コマンドラインで、ウェブページをダウンロードする

# Details #

![http://3.bp.blogspot.com/_EUW0nrj9XlM/TU3_c7YopvI/AAAAAAAAACk/K7ZgW1HjSJQ/s1600/shot1.png](http://3.bp.blogspot.com/_EUW0nrj9XlM/TU3_c7YopvI/AAAAAAAAACk/K7ZgW1HjSJQ/s1600/shot1.png)

  * do a pinpoint downloading to a web page of file. and do a bulk downloading to a list of urls.
  * no preconditions are required for almost Windows, because it's all written by `VBScript` .
  * and free to customize.
  * to download a single file, do `CScript httpget.vbs TARGET_URL SAVE_FILE_NAME`
  * to download multiple files, do `CScript httpgets.vbs < FILE_NAME_OF_URL_AND_SAVE_LIST`
  * use `httpgetw.vbs` for GUI.
  * see [hello\_http\_get](hello_http_get.md) and [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md)

## 説明 ##
  * 個別ページやファイルをピンポイント指定でダウンロードできる他、URLの一覧をテキストファイルで作っておき、一括して取得することもできる。
  * `VBScript` だけで書かれているので、ほとんどのウィンドウズ環境でそのまま動く。
  * 好き勝手にカスタマイズできる。
  * 単一ファイルの取得は、 `CScript httpget.vbs ダウンロードするURL名 保存するファイル名`
  * 複数ファイルの取得は、 `CScript httpgets.vbs < ダウンロードするURLとファイル名のリスト`
  * `httpgetw.vbs` は GUI 用
  * [hello\_http\_get](hello_http_get.md) と [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) を参考

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=http_get)

# How to use #

  1. open a Windows Command Prompt. (cmd.exe)
  1. enter `CScript httpget.vbs http://cowares.nobody.jp/favicon.ico favicon.ico`
  1. you've got a favicon.ico saved in the current directory.

## 使い方 ##
  1. コマンドプロンプトを開く。
  1. `CScript httpget.vbs http://cowares.nobody.jp/favicon.ico favicon.ico`
  1. カレントディレクトリに favicon.ico が保存されるはず。


# Code #

### httpget.vbs ###

```
' httpget
' download a file from web
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript httpget.vbs http://cowares.nobody.jp/favicon.ico C:\tmp\favicon.ico

On Error Resume Next
Set Args = WScript.Arguments
Main Args(0), Args(1)
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Url, FileName)
    Dim tp
    Set tp = CreateObject("MSXML2.XMLHTTP")
    'Set tp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    tp.Open "GET", Url, False
    tp.send
    WScript.Echo tp.Status & " " & tp.statusText & " " & Url
    SaveBinaryFile tp.responseBody, FileName
    
    Set tp = Nothing
End Sub

Function SaveBinaryFile(Data, FileName)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    Dim Stream
    
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = adTypeBinary
    Stream.Write Data
    Stream.SaveToFile FileName, adSaveCreateOverWrite
    Stream.Close
    Set Stream = Nothing
End Function
```

### httpgets.vbs ###

```
' httpgets
' download files from web
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript httpgets.vbs < urls.txt

' urls.txt contains list of url and filename pairs
' see examples in urls.txt

On Error Resume Next
Main
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main()
    Dim Text, UrlFileName
    Dim StdIn
    Set StdIn = WScript.StdIn
    
    Do Until StdIn.AtEndOfStream
        Text = Trim(StdIn.ReadLine)
        If Left(Text,4) = "http" Then
            UrlFileName = Split(Text, " ", 2)
            If Ubound(UrlFileName) = 1 Then
                HttpGet UrlFileName(0), UrlFileName(1)
            End If
        End If
    Loop
    
    Set StdIn = Nothing
End Sub

Sub HttpGet(Url, FileName)
    Dim tp
    Set tp = CreateObject("MSXML2.XMLHTTP")
    'Set tp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    tp.Open "GET", Url, False
    tp.send
    WScript.Echo tp.Status & " " & tp.statusText & " " & Url
    SaveBinaryFile tp.responseBody, FileName
    
    Set tp = Nothing
End Sub

Function SaveBinaryFile(Data, FileName)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    Dim Stream
    
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = adTypeBinary
    Stream.Write Data
    Stream.SaveToFile FileName, adSaveCreateOverWrite
    Stream.Close
    Set Stream = Nothing
End Function
```

### httpgetw.vbs ###

```
' httpgetw
' download files from web
' does same as httpgets, though it accepts files instead of stdin
' also it will generate incremental file names if omitted
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript httpgetw.vbs urls.txt [more_urls.txt]

' urls.txt contains list of url and filename pairs
' file paths are considered to be relative to each list file.
' see examples in urls.txt

On Error Resume Next
Set StdErr = new StringStream
Set Args = WScript.Arguments
Main Args
If Err.Number <> 0 Then WScript.Echo Err.Description
If StdErr.Text <> "" Then WScript.Echo StdErr.Text
WScript.Quit(Err.Number)

Sub Main(ListFiles)
    Dim Text, UrlFileName, SaveFile, FileCounter, i, x
    Dim ListFile, fs, Shell
    Dim StdIn
    Const TristateFalse = 0
    Const ForReading = 1
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Shell = CreateObject("WScript.Shell")
    
    For Each ListFile In ListFiles
        If fs.FileExists(ListFIle) Then
            StdErr.WriteLine ListFile
            Shell.CurrentDirectory = fs.GetFile(ListFIle).ParentFolder
            
            FileCounter = 0
            For Each File In fs.GetFolder(Shell.CurrentDirectory).Files
                x = Left(File.Name, InStr(File.Name, "."))
                If IsNumeric(x) Then
                    i = CLng(x)
                    If FileCounter < i Then FileCounter = i
                End If
            Next
            
            Set StdIn = fs.OpenTextFile(ListFile, ForReading, False, TristateFalse)
            
            Do Until StdIn.AtEndOfStream
                Text = Trim(StdIn.ReadLine)
                If Left(Text,4) = "http" Then
                    UrlFileName = Split(Text, " ", 2)
                    If Ubound(UrlFileName) = 0 Then
                        FileCounter = FileCounter + 1
                        SaveFile = Cstr(FileCounter) & ".bin"
                    ElseIf Ubound(UrlFileName) = 1 Then
                        SaveFile = UrlFileName(1)
                    Else
                        SaveFile = ""
                    End If
                    If SaveFile <> "" Then
                        HttpGet UrlFileName(0), SaveFile
                    End If
                End If
            Loop
            
            StdIn.Close
            Set StdIn = Nothing
        Else
            StdErr.WriteLine "Warning: skipping unavailable file: " & ListFile
        End If
    Next
    
    Set Shell = Nothing
    Set fs = Nothing
End Sub

Sub HttpGet(Url, ByVal FileName)
    Dim tp
    Set tp = CreateObject("MSXML2.XMLHTTP")
    'Set tp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    On Error Resume Next
    tp.Open "GET", Url, False
    tp.send
    If Err.Number <> 0 Then
        StdErr.WriteLine Err.Number & " " & Err.Description & " " & Url
        Err.Clear
    ElseIf tp.Status <> 200 Then
        StdErr.WriteLine tp.Status & " " & tp.statusText & " " & Url
    Else
        If FileName = "" Then
            FileName = Right(Url, Len(Url) - InStrRev(Url, "/"))
        ElseIf Right(FileName, 1) = "." Then
            FileName = FileName & Replace(Split(tp.getResponseHeader("content-type"), ";")(0), "/", ".")
        End If
        SaveBinaryFile tp.responseBody, FileName
    End If
    
    Set tp = Nothing
End Sub

Function SaveBinaryFile(Data, FileName)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    Dim Stream
    
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = adTypeBinary
    Stream.Write Data
    Stream.SaveToFile FileName, adSaveCreateOverWrite
    Stream.Close
    Set Stream = Nothing
End Function

Class StringStream
    Public Text
    
    Public Sub WriteLine(Data)
        Text = Text & Data & vbCrLf
    End Sub
End Class
```

### example.bat ###

```
@ECHO OFF

REM httpget example

SET MSG=Done
CScript httpget.vbs http://cowares.nobody.jp/favicon.ico C:\tmp\favicon.ico
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

SET MSG=Done
CScript httpget.vbs http://cowares.nobody.jp/favicon.ic C:\tmp\favicon.err
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

REM httpgets example

SET MSG=Done
CScript httpgets.vbs < urls.txt
IF NOT ERRORLEVEL 0 SET MSG=Error %ERRORLEVEL%
ECHO %MSG%

```

### urls.txt ###

```
http://www.post.japanpost.jp/zipcode/dl/oogaki/zip/37kagawa.zip C:\tmp\oogaki37kagawa.zip
http://www.post.japanpost.jp/zipcode/dl/kogaki/zip/37kagawa.zip C:\tmp\kogaki37kagawa.zip
http://www.post.japanpost.jp/zipcode/dl/roman/37kagawa_rome.zip C:\tmp\roman37kagawa.zip
```
