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
