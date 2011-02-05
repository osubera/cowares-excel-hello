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
