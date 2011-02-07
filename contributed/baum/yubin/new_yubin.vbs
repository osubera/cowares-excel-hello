' new_yubin
' initialize mdb file from zip numbers csv.
' csv url: http://www.post.japanpost.jp/zipcode/download.html
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

Const adOpenForwardOnly = 0
Const adLockOptimistic = 3
Const TristateFalse = 0
Const ForReading = 1
Const CsvNA = 0
Const CsvOogaki = 1
Const CsvJigyo = 2
Const ColumnsOogaki = 15
Const ColumnsJigyo = 13

On Error Resume Next
Set Args = WScript.Arguments
Set StdErr = New StringStream
Main Args
If Err.Number <> 0 Then WScript.Echo Err.Description
If StdErr.Text <> "" Then WScript.Echo StdErr.Text
WScript.Quit(Err.Number)

Private Function GetMdbName()
    GetMdbName = Replace(WScript.ScriptFullName, WScript.ScriptName, "yubin.vbs") & ".mdb"
    'GetMdbName = "C:\tmp\yubin.vbs.mdb"
End Function

Private Function GetConnectionString()
    GetConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
        GetMdbName & _
        ";User ID=Admin;Password=;"
End Function

Sub Main(Files)
    Dim Con, File
    
    Set Con = CreateObject("ADODB.Connection")
    Con.Open GetConnectionString()
    
    For Each File In Files
        UpdateFromCsv Con, File
    Next
    
    Con.Close
    Set Con = Nothing
End Sub

Sub UpdateFromCsv(Con, FileName)
    Select Case WhichCsv(FileName)
    Case CsvOogaki
        StdErr.Print "市町村郵便番号の更新: " & FileName
        AppendCsvOogaki Con, FileName
    Case CsvJigyo
        StdErr.Print "大口事業所の更新: " & FileName
        AppendCsvJigyo Con, FileName
    Case Else
        StdErr.Print "読めないデータ: " & FileName
    End Select
End Sub

Sub AppendCsvJigyo(Con, FileName)
    AppendZipCode Con, FileName, "zipcode_j", RegParseCsvJigyo, ColumnsJigyo
    Con.Execute "make_jigyo"
End Sub

Sub AppendCsvOogaki(Con, FileName)
    AppendZipCode Con, FileName, "zipcode_k", RegParseCsvOogaki, ColumnsOogaki
    AppendKenShiCho Con
End Sub

Sub AppendKenShiCho(Con)
    Con.Execute "make_ken"
    Con.Execute "make_shi_yubin"
    Con.Execute "make_cho"
End Sub

Sub AppendZipCode(Con, FileName, TableName, R, ColumnsNumber)
    Dim dbs, fs, ts, rRes, RawText, i
    
    Set dbs = CreateObject("ADODB.Recordset")
    Set fs = CreateObject("Scripting.FileSystemObject")
    dbs.Open TableName, Con, adOpenForwardOnly, adLockOptimistic
    Set ts = fs.OpenTextFile(FileName, ForReading, False, TristateFalse)
    
    Con.BeginTrans
    Do Until ts.AtEndOfStream
        RawText = ts.ReadLine
        Set rRes = R.Execute(RawText)
        If rRes.Count = 0 Then
            StdErr.Print "Parse Error: " & RawText
        Else
            If rRes(0).SubMatches.Count <> ColumnsNumber Then
                StdErr.Print "Parse Error: " & RawText
            Else
                dbs.AddNew
                For i = 1 To ColumnsNumber
                    dbs.Fields(i) = rRes(0).SubMatches(i - 1)
                Next
                dbs.Update
            End If
        End If
    Loop
    Con.CommitTrans
    
    ts.Close
    dbs.Close
    Set ts = Nothing
    Set fs = Nothing
    Set dbs = Nothing
End Sub

Private Function RegParseCsvOogaki()
    Dim R, i
    Const PNum = "([^,""]*)\s*"
    Const PStr = """([^,""]*)\s*"""
    Const PC = ","
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = True
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = PNum
    For i = 1 To 8
        R.Pattern = R.Pattern & PC & PStr
    Next
    For i = 1 To 6
        R.Pattern = R.Pattern & PC & PNum
    Next
    
    Set RegParseCsvOogaki = R
End Function

Private Function RegParseCsvJigyo()
    Dim R, i
    Const PNum = "([^,""]*)\s*"
    Const PStr = """([^,""]*)\s*"""
    Const PC = ","
    
    Set R = CreateObject("VBScript.RegExp")
    R.Global = True
    R.IgnoreCase = False
    R.MultiLine = False
    R.Pattern = PNum
    For i = 1 To 9
        R.Pattern = R.Pattern & PC & PStr
    Next
    For i = 1 To 3
        R.Pattern = R.Pattern & PC & PNum
    Next
    
    Set RegParseCsvJigyo = R
End Function

Function WhichCsv(FileName)
    Dim out, fs, ts, rRes, RawText
    
    out = CsvNA
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.OpenTextFile(FileName, ForReading, False, TristateFalse)
    
    If Not ts.AtEndOfStream Then
        RawText = ts.ReadLine
        Set rRes = RegParseCsvOogaki.Execute(RawText)
        If rRes.Count > 0 Then
            out = CsvOogaki
        Else
            Set rRes = RegParseCsvJigyo.Execute(RawText)
            If rRes.Count > 0 Then out = CsvJigyo
        End If
        Set rRes = Nothing
    End If
    
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
    
    WhichCsv = out
End Function

Class StringStream
    Public Text
    
    Public Sub Print(Data)
        Text = Text & Data & vbCrLf
    End Sub
End Class
