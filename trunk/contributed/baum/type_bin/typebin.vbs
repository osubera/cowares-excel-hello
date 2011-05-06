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
