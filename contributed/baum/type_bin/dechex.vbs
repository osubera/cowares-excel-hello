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
