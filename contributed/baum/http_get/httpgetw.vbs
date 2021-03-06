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
