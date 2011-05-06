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
