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
