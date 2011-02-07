' head
' print the first 10 lines of each file to standard output.
' read standard input when without file or the file name is `-`
' precede a file name to each file when more than one files are specified.
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

' usage> CScript //Nologo head.vbs [/n:LINE NUMBER] [/c:CHARACTER NUMBER] [FILE]...

On Error Resume Next
Set Args = WScript.Arguments
Main Args.Unnamed, Args.Named
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

Sub Main(Files, Opts)
    Dim vFiles, FilesCount
    Dim vLines
    
    FilesCount = Files.Count
    If FilesCount = 0 Then
        vFiles = Array("-")
        FilesCount = 1
    Else
        Set vFiles = Files
    End If
    
    vLines = 10
    If Opts("n") <> "" Then
        vLines = CLng(Opts("n"))
    End If
    
    vChars = 0
    If Opts("c") <> "" Then
        vChars = CLng(Opts("c"))
    End If
    
    Head vFiles, FilesCount, vLines, vChars
End Sub

Sub Head(Files, FilesCount, vLines, vChars)
    Dim File, WithHeader, CountDown
    Dim StdOut, WithConsole
    Dim fs, ts
    Const TristateFalse = 0
    Const ForReading = 1
    
    WithConsole = HasConsole()
    If WithConsole Then
        Set StdOut = WScript.StdOut
    Else
        Set StdOut = New StringStream
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    WithHeader = (FilesCount > 1)
    
    For Each File in Files
        If WithHeader Then StdOut.WriteLine File
        
        If File = "-" Then
            Set ts = WScript.StdIn
        Else
            Set ts = fs.OpenTextFile(File, ForReading, False, TristateFalse)
        End If
        
        If vChars > 0 Then
            HeadCharAFile ts, StdOut, vChars
        Else
            HeadLineAFile ts, StdOut, vLines
        End If
        
        ts.Close
        Set ts = Nothing
    Next
    
    If Not WithConsole Then WScript.Echo StdOut.Text
    
    Set fs = Nothing
    Set StdOut = Nothing
End Sub

Sub HeadLineAFile(inSt, outSt, ByVal Counter)
    Do Until inSt.AtEndOfStream
        Counter = Counter - 1
        If Counter < 0 Then Exit Do
        outSt.WriteLine inSt.ReadLine
    Loop
End Sub

Sub HeadCharAFile(inSt, outSt, ByVal Counter)
    If inSt.AtEndOfStream Then Exit Sub
    outSt.Write inSt.Read(Counter) & vbCrLf
End Sub

Function HasConsole()
    HasConsole = (UCase(Left(Right(WScript.FullName,11),1)) = "C")
End Function

Class StringStream
    Public Text
    
    Public Sub WriteLine(Data)
        Text = Text & Data & vbCrLf
    End Sub
    
    Public Sub Write(Data)
        Text = Text & Data
    End Sub
End Class
