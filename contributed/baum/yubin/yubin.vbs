' yubin
' search zip numbers of Japan.
' Copyright (C) 2011 Tomizono - kobobau.com
' Fortitudinous, Free, Fair, http://cowares.nobody.jp

Const PopX = 500    ' x location by screen twips
Const PopY = 1000   ' y location by screen twips

Const HowNA = 0
Const HowFromYubin = 1
Const HowFromCho = 2
Const HowFromJigyo = 3
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Dim Bag     ' global resources

On Error Resume Next
Set Bag = New GlobalResources
Set Args = WScript.Arguments
Main Args
If Err.Number <> 0 Then WScript.Echo Err.Description
WScript.Quit(Err.Number)

'=== main flows begin ===

Sub Main(Args)
    If WScript.Interactive Then
        If Args.Count > 0 Then
            SearchCommand Args
        Else
            SearchPrompt HasConsole()
        End If
    Else
        SearchBatch
    End If
End Sub

Sub SearchPrompt(WithConsole)
    Dim AText, out
    
    Do
        If WithConsole Then
            WScript.Echo out
            AText = WScript.StdIn.ReadLine
        Else
            AText = InputBox(out, Bag.Title, AText, PopX, PopY)
        End If
        
        If AText = "" Then Exit Sub
        
        Bag.Joken.SetText AText
        out = DoSearch()
    Loop
End Sub

Sub SearchCommand(Args)
    Bag.Joken.SetArgs Args
    WScript.Echo DoSearch()
End Sub

Sub SearchBatch()
    Dim AText, out
    Dim inS, outS
    
    Set inS = WScript.StdIn
    Set outS = WScript.StdOut
    
    Do Until inS.AtEndOfStream
        AText = inS.ReadLine
        Bag.Joken.SetText AText
        out = DoSearch()
        outS.WriteLine "<" & AText
        outS.WriteLine out
    Loop
    
    outS.Close
    inS.Close
    Set outS = Nothing
    Set inS =  Nothing
End Sub

'=== main flows end ===
'=== core seacher begin ===

Function DoSearch()
    Dim out
    
    Select Case Bag.Joken.How
    Case HowFromYubin
        out = DoYubinSearch()
    Case HowFromCho
        out = DoChoSearch()
    Case HowFromJigyo
        If Bag.Joken.Cho = "" Then
            out = DoChoSearch()
        Else
            out = DoJigyoSearch()
        End If
    End Select
    'out = Bag.Joken.Dump
    
    DoSearch = out
End Function

Function DoChoSearch()
    Dim out
    
    If Bag.Joken.Ken = "åß" Or Bag.Joken.Ken = "ÇØÇÒ" Then
        out = ListKen()
    ElseIf Bag.Joken.Shi = "" Then
        out = ListShi()
    ElseIf Bag.Joken.Cho = "" Then
        out = FromShi()
    Else
        out = FromCho()
    End If
    
    DoChoSearch = out
End Function

Function DoJigyoSearch()
    DoJigyoSearch = FromJigyo()
End Function

Function DoYubinSearch()
    DoYubinSearch = FromYubin()
End Function

'=== core seacher end ===
'=== SQL common functions begin ===

Function GetId(Con, Sql)
    Dim dbs, out
    
    Set dbs = CreateObject("ADODB.Recordset")
    dbs.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
    
    If dbs.EOF Then
        out = 0
    Else
        out = dbs.Fields(0)
    End If
    
    dbs.Close
    Set dbs = Nothing
    
    GetId = out
End Function

Function ListRecords(Con, Sql)
    Dim dbs, out
    
    Set dbs = CreateObject("ADODB.Recordset")
    dbs.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
    
    Do Until dbs.EOF
        out = out & dbs.Fields(0) & vbCrLf
        dbs.MoveNext
    Loop
    
    dbs.Close
    Set dbs = Nothing
    
    ListRecords = out
End Function

Function ListRecordsTab(Con, Sql, Cols)
    Dim dbs, out, i
    
    Set dbs = CreateObject("ADODB.Recordset")
    dbs.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
    
    Do Until dbs.EOF
        i = i + 1
        out = out & dbs.Fields(0) & vbTab
        If i Mod Cols = 0 Then out = out & vbCrLf
        dbs.MoveNext
    Loop
    
    dbs.Close
    Set dbs = Nothing
    
    ListRecordsTab = out
End Function

Function ListRecordsLength(Con, Sql, Length)
    Dim dbs, out, i
    
    Set dbs = CreateObject("ADODB.Recordset")
    dbs.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
    
    Do Until dbs.EOF
        i = i + Len(dbs.Fields(0)) + 1
        If i > Length Then
            out = out & vbCrLf
            i = Len(dbs.Fields(0)) + 1
        End If
        out = out & dbs.Fields(0) & " "
        dbs.MoveNext
    Loop
    
    dbs.Close
    Set dbs = Nothing
    
    ListRecordsLength = out
End Function

'=== SQL common functions end ===
'=== SQL functions begin ===

Function ListKen()
    ListKen = ListRecordsTab(Bag.Connect, _
        "SELECT ken_name FROM ken ORDER BY code", _
        4)
End Function

Function ListShi()
    Dim KenCode
    
    KenCode = GetKen()
    ListShi = ListRecordsLength(Bag.Connect, _
        "SELECT shi_name FROM shi WHERE code BETWEEN " & _
        KenCode & "000 AND " & KenCode & "999 ORDER BY code", _
        21)
End Function

Function FromYubin()
    FromYubin = ListRecords(Bag.Connect, _
        "SELECT yubin & ' ' & ken_name & ' ' & shi_name & ' ' & cho_name AS out" & _
        "  FROM cho INNER JOIN (shi INNER JOIN ken ON shi.ken_code = ken.code) ON cho.code = shi.code" & _
        " WHERE yubin BETWEEN '" & Bag.Joken.YubinB & "' AND '" & Bag.Joken.YubinE & "' ORDER BY yubin")
End Function

Function FromShi()
    Dim Code
    Code = GetShi()
    FromShi = ListRecords(Bag.Connect, _
        "SELECT yubin_misc & ' ' & ken_name & ' ' & shi_name & ' Å¶'  AS out" & _
        "  FROM shi INNER JOIN ken ON shi.ken_code = ken.code" & _
        " WHERE shi.code = " & Code & " AND yubin_misc IS NOT NULL")
End Function

Function FromCho()
    Dim Code
    Code = GetShi()
    
    FromCho = ListRecords(Bag.Connect, _
        "SELECT yubin & ' ' & ken_name & ' ' & shi_name & ' ' & cho_name AS out" & _
        "  FROM cho INNER JOIN (shi INNER JOIN ken ON shi.ken_code = ken.code) ON cho.code = shi.code" & _
        " WHERE cho.code = " & Code & " AND cho_name LIKE '" & Bag.Joken.Cho & "%' ORDER BY yubin")
    If FromCho <> "" Then Exit Function
    
    FromCho = ListRecords(Bag.Connect, _
        "SELECT yubin & ' ' & ken_name & ' ' & shi_name & ' ' & cho_name AS out" & _
        "  FROM cho INNER JOIN (shi INNER JOIN ken ON shi.ken_code = ken.code) ON cho.code = shi.code" & _
        " WHERE cho.code = " & Code & " AND cho_hira LIKE '" & ToOogaki(Bag.Joken.Cho) & "%' ORDER BY yubin")
    If FromCho <> "" Then Exit Function
    
    FromCho = FromShi()
End Function

Function FromJigyo()
    Dim Code
    Code = GetShi()
    
    FromJigyo = ListRecords(Bag.Connect, _
        "SELECT yubin & ' ' & ken_name & ' ' & shi_name & ' ' & cho_name AS out" & _
        "  FROM jigyo AS cho INNER JOIN (shi INNER JOIN ken ON shi.ken_code = ken.code) ON cho.code = shi.code" & _
        " WHERE cho.code = " & Code & " AND cho_name LIKE '" & Bag.Joken.Cho & "%' ORDER BY yubin")
    If FromJigyo <> "" Then Exit Function
    
    FromJigyo = ListRecords(Bag.Connect, _
        "SELECT yubin & ' ' & ken_name & ' ' & shi_name & ' ' & cho_name AS out" & _
        "  FROM jigyo AS cho INNER JOIN (shi INNER JOIN ken ON shi.ken_code = ken.code) ON cho.code = shi.code" & _
        " WHERE cho.code = " & Code & " AND cho_hira LIKE '" & ToOogaki(Bag.Joken.Cho) & "%' ORDER BY yubin")
End Function

Function GetShi()
    Dim out, Ken
    
    Ken = GetKen()
    out = GetId(Bag.Connect, "SELECT code FROM shi WHERE ken_code = " & Ken & " AND shi_name = '" & Bag.Joken.Shi & "'")
    If out = 0 Then
        out = GetId(Bag.Connect, "SELECT code FROM shi WHERE ken_code = " & Ken & " AND shi_hira = '" & ToOogaki(Bag.Joken.Shi) & "'")
    End If
    
    GetShi = out
End Function

Function GetKen()
    Dim out
    
    out = GetId(Bag.Connect, "SELECT code FROM ken WHERE ken_name = '" & Bag.Joken.Ken & "'")
    If out = 0 Then
        out = GetId(Bag.Connect, "SELECT code FROM ken WHERE ken_hira = '" & ToOogaki(Bag.Joken.Ken) & "'")
    End If
    
    GetKen = out
End Function

'=== SQL functions end ===
'=== utility functions begin ===

Function ToOogaki(Text)
    Dim x, out
    
    out = Text
    For Each x In Array( _
            Array("Çü", "Ç†"), _
            Array("Ç°", "Ç¢"), _
            Array("Ç£", "Ç§"), _
            Array("Ç•", "Ç¶"), _
            Array("Çß", "Ç®"), _
            Array("Ç¡", "Ç¬"), _
            Array("Ç·", "Ç‚"), _
            Array("Ç„", "Ç‰"), _
            Array("ÇÂ", "ÇÊ"), _
            Array("ÇÏ", "ÇÌ") _
        )
        out = Replace(out, x(0), x(1))
    Next
    
    ToOogaki = out
End Function

Function HasConsole()
    HasConsole = (UCase(Left(Right(WScript.FullName,11),1)) = "C")
End Function

'=== utility functions end ===
'=== classes begin ===
' GlobalResources, SearchConditions

Class GlobalResources
    Public Title, KenKotei, ShiKotei
    Public Joken, Connect
    Public Shell
    
    Private Sub Class_Initialize
        Set Shell = CreateObject("WScript.Shell")
        Set Joken = New SearchConditions
        
        KenKotei = Shell.Environment("USER")("KENKOTEI")
        If KenKotei <> ""  Then ShiKotei = Shell.Environment("USER")("SHIKOTEI")
        Title = KenKotei & ShiKotei & " Åú óXï÷î‘çÜåüçı - yubin"
        
        Set Connect = CreateObject("ADODB.Connection")
        Connect.Open GetConnectionString()
    End Sub
    
    Private Sub Class_Terminate
        Connect.Close
        Set Connect = Nothing
        Set Joken = Nothing
        Set Shell = Nothing
    End Sub
    
    Private Function GetMdbName()
        GetMdbName = WScript.ScriptFullName & ".mdb"
        'GetMdbName = "C:\tmp\db1.mdb"
    End Function

    Private Function GetConnectionString()
        GetConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            GetMdbName & _
            ";User ID=Admin;Password=;"
    End Function
End Class

Class SearchConditions
    Public Ken, Shi, Cho
    Public Yubin, YubinB, YubinE
    Public How
    
    Public Sub SetText(Text)
        SetConditions Split(Replace(Text, "Å@", " "))
    End Sub
    
    Public Sub SetArgs(Args)
        SetConditions Args
    End Sub
    
    Public Function Dump()
        Dump = Join(Array(How, Yubin, YubinB, YubinE, Ken, Shi, Cho), vbCrLf)
    End Function
    
    Private Sub SetConditions(p)
        Dim x, i
        Dim q(2)
        
        How = HowNA
        
        If Bag.KenKotei <> "" Then
            q(0) = Bag.KenKotei
            i = 1
        End If
        If Bag.ShiKotei <> "" Then
            q(1) = Bag.ShiKotei
            i = 2
        End If
        
        For Each x In p
            x = Trim(x)
            If x <> "" Then
                Select Case How
                Case HowNA
                    ' The First Item
                    If Left(x, 1) = "Åê" Or Left(x, 1) = "$" Then
                        How = HowFromJigyo
                    Else
                        If IsNumeric(Left(x, 1)) Then
                            How = HowFromYubin
                            q(0) = ""
                            ScoopNumber x, q(0)
                        Else
                            How = HowFromCho
                            q(i) = x
                            i = i + 1
                        End If
                    End If
                Case HowFromYubin
                    ' The 2nd and after
                    ScoopNumber x, q(0)
                Case Else
                    ' The 2nd and after
                    If i > 2 Then
                        q(2) = q(2) & x
                    Else
                        q(i) = x
                    End If
                    i = i + 1
                End Select
            End If
        Next
        
        If How = HowFromYubin Then
            Yubin = Left(q(0), 7)
            i = Len(Yubin)
            If i = 7 Then
                YubinB = Yubin
                YubinE = Yubin
            Else
                YubinB = Yubin & String(7 - i, "0")
                YubinE = Yubin & String(7 - i, "9")
            End If
        Else
            Ken = SafeString(q(0))
            Shi = SafeString(q(1))
            Cho = SafeString(q(2))
        End If
    End Sub
    
    Private Sub ScoopNumber(Text, out)
        Dim i, x
        
        For i = 1 to Len(Text)
            x = Mid(Text, i, 1)
            If IsNumeric(x) Then out = out & CStr(CLng(x))
        Next
    End Sub
    
    Private Function SafeString(Text)
        SafeString = Replace(Replace(Text, """", " "), "'", " ")
    End Function
End Class

'=== classes end ===
