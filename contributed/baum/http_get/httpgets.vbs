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
    'Set tp = CreateObject("MSXML2.XMLHTTP")
    Set tp = CreateObject("MSXML2.XMLHTTP.6.0")
    
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
