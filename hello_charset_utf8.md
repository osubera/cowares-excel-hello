# Introduction #

  * how to convert utf-8 and unicode at a low level in vba
  * while using other built-in functions as described at [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) is smart, codes here still are useful to understand what is the unicode in vba, and what is utf-8 encoding.
  * also provides binary dump tools.

## 概要 ##
  * VBAで utf-8 と unicode の低レベルな変換をする
  * [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) で使ったような組み込み関数を使うのが楽だけど、このコードは、 VBA でのユニコードの扱いや utf-8 エンコードが何なのか知るのに役立つ。
  * また、バイト列を書き出すツールも一緒に入れている。

# Details / 説明 #

| BomUnicode | constant array of unicode bom, FF FE. <br /> ユニコードBOMの定数配列 |
|:-----------|:-----------------------------------------------------------|
| BomUTF8    | utf-8 encoded unicode bom.  <br /> UTF-8 に変換したユニコードBOM     |
| EncUTF8    | convert unicode words (integers) into utf-8 bytes. <br /> ユニコードワード（整数）列を utf-8 バイト列に |
| DecUTF8    | convert utf-8 bytes into unicode integers. <br /> utf-8 バイト列をユニコード整数列に |
| BytesToString | convert bytes array into a text String. <br /> バイト列を String 型テキストに |
| StringToBytes | convert a text String into bytes array. <br /> String 型テキストをバイト列に |
| WordsToString | convert words array into a text String. <br /> ワード列を String 型テキストに |
| StringToWords | convert a text String into words array. <br /> String 型テキストをワード列に |
| LongsToString | convert ANSI long integers array into a text String. <br /> ANSI (Shift\_JIS) ロング列を String 型テキストに |
| StringToLongs | convert a text String into ANSI long integers array. <br /> String 型テキストを ANSI ロング列に |
| BytesToWords | convert bytes array into words array. <br /> バイト列をワード列に    |
| WordsToBytes | convert words array into bytes array. <br /> ワード列をバイト列に    |
| EncHexDelimited | convert numbers array into a text String as hexadecimal representation. <br /> 数値列を１６進表記のテキストに |
| DecHexDelimited | convert a hexadecimal text String into bytes array. <br /> １６進表記のテキストをバイト列に |
| DecHexDelimitedL | convert a hexadecimal text String into long integers array. <br /> １６進表記のテキストをロング列に |
| FixLineFeed | replace every linefeed character into the specified one. <br /> 改行記号を指定のものに揃える |

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. test1() does demonstrations and round trip tests.

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. test1() がデモと相互変換のテストをする。

# Code #

```

'workbook
'  name;hello_charset_utf8.xls

'module
'  name;HelloCharsetUtf8
'{{{
Option Explicit

' walking through the Unicode and UTF-8 in VBA

Const LineFeed = vbCrLf
Const Delimiter = " "

Sub test1()
    ' check round trip conversions
    Dim Text As String
    
    ' see utf8 bom conversions
    Debug.Print EncHexDelimited(BomUnicode)         ' FF FE
    Debug.Print EncHexDelimited(BomUTF8)            ' EF BB BF
    Debug.Print EncHexDelimited(DecUTF8(BomUTF8))   ' EF BB BF
    Debug.Print IIf(EncHexDelimited(BomUnicode) = EncHexDelimited(WordsToBytes(DecUTF8(BomUTF8))), "OK", "NG")
    
    ' see Unicode and ANSI, differences in AscB(), AscW() and Asc() function
    Text = "Free 自由"
    Debug.Print EncHexDelimited(StringToBytes(Text))    ' Unicode as byte array
    Debug.Print EncHexDelimited(StringToWords(Text))    ' Unicode as integer array
    Debug.Print EncHexDelimited(StringToLongs(Text))    ' ANSI (Shift_JIS) as long array
    Debug.Print IIf(Text = BytesToString(StringToBytes(Text)), "OK", "NG")
    Debug.Print IIf(Text = WordsToString(StringToWords(Text)), "OK", "NG")
    Debug.Print IIf(Text = LongsToString(StringToLongs(Text)), "OK", "NG")
    
    ' see utf8 conversions
    Debug.Print EncHexDelimited(EncUTF8(StringToWords(Text)))           ' UTF-8 encode
    Debug.Print EncHexDelimited(DecUTF8(EncUTF8(StringToWords(Text))))  ' UTF-8 decode
    Debug.Print WordsToString(DecUTF8(EncUTF8(StringToWords(Text))))    ' to Unicode string
    Debug.Print IIf(Text = WordsToString(DecUTF8(EncUTF8(StringToWords(Text)))), "OK", "NG")
    
    ' more
    Text = "4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30"
    Debug.Print EncHexDelimited(DecHexDelimited(Text))
    Debug.Print IIf(Text & Delimiter = EncHexDelimited(DecHexDelimited(Text)), "OK", "NG")
    Debug.Print EncHexDelimited(BytesToWords(DecHexDelimited(Text)))
    Debug.Print IIf(Text & Delimiter = EncHexDelimited(WordsToBytes(BytesToWords(DecHexDelimited(Text)))), "OK", "NG")
    
    Text = "AB" & vbCrLf & "cd" & vbCr & "EF" & vbLf & vbCr & "ghij" & vbLf & vbLf & "KLM"
    Debug.Print FixLineFeed(Text, "##")
    Debug.Print FixLineFeed(Text)
    Debug.Print EncHexDelimited(StringToBytes(Text))
    Debug.Print EncHexDelimited(StringToBytes(FixLineFeed(Text)))
End Sub


' UTF-8エンコード

Function BomUnicode() As Byte()
    Dim out(0 To 1) As Byte
    out(0) = &HFF
    out(1) = &HFE
    BomUnicode = out
End Function

Function BomUTF8() As Byte()
    BomUTF8 = EncUTF8(BytesToWords(BomUnicode))
End Function

Function EncUTF8(Data As Variant, Optional ByVal Length As Long = -1) As Byte()
    Dim out() As Byte
    Dim pan As Variant
    Dim i As Long
    Dim LengthUTF8 As Long
    Dim UpperByte As Long
    Dim LowerByte As Long
    Dim MiddleByte As Long
    
    pan = Empty
    LengthUTF8 = 0
    If Length = -1 Then Length = UBound(Data) + 1
    
    For i = 0 To Length - 1
        Select Case Data(i)
        Case 0 To 127       ' goes 1 byte for 7bit ascii
            ' aaaabbb0 00000000 → aaaabbb0
            ' SingleByte: begins 0, w.band(7)
            pan = Array(CByte(Data(i) And &H7F), pan)
            LengthUTF8 = LengthUTF8 + 1
        Case 128 To 2047    ' goes 2 bytes for many european languages
            ' aaaabbbb ccc00000 → aaaabb01 bbccc011 → bbccc011 aaaabb01
            ' UpperByte: begins 110, w.shift(6).band(5)
            ' LowerByte: begins 10,  w.band(6)
            ' the upper goes 1st, to help decoding
            UpperByte = CByte((Int(Data(i) / &H40) And &H1F) Or &HC0)
            LowerByte = CByte((Data(i) And &H3F) Or &H80)
            pan = Array(LowerByte, Array(UpperByte, pan))
            LengthUTF8 = LengthUTF8 + 2
        Case Else           ' goes 3 bytes for far east languages
            ' aaaabbbb ccccdddd → aaaabb01 bbcccc01 dddd0111 → dddd0111 bbcccc01 aaaabb01
            ' UpperByte:  begins 1110, w.shift(12).band(4)
            ' MiddleByte: begins 10,   w.shift(6).band(6)
            ' LowerByte:  begins 10,   w.band(6)
            UpperByte = CByte((Int(Data(i) / &H1000) And &HF) Or &HE0)
            MiddleByte = CByte((Int(Data(i) / &H40) And &H3F) Or &H80)
            LowerByte = CByte((Data(i) And &H3F) Or &H80)
            pan = Array(LowerByte, Array(MiddleByte, Array(UpperByte, pan)))
            LengthUTF8 = LengthUTF8 + 3
        End Select
    Next
    
    If LengthUTF8 > 0 Then
        ReDim out(0 To LengthUTF8 - 1)
        For i = LengthUTF8 - 1 To 0 Step -1
            out(i) = pan(0)
            pan = pan(1)
        Next
    End If
    
    EncUTF8 = out
End Function

Function DecUTF8(Data As Variant, Optional ByVal Length As Long = -1) As Integer()
    Dim out() As Integer
    Dim pan As Variant
    Dim i As Long
    Dim LengthUnicode As Long
    Dim UpperByte As Integer
    Dim LowerByte As Integer
    Dim UnicodeWord As Integer
    Dim Continued As Long
    
    pan = Empty
    LengthUnicode = 0
    If Length = -1 Then Length = UBound(Data) + 1
    
    For i = 0 To Length - 1
        Select Case Data(i)
        Case 0 To &H7F      ' 0xxx-,  goes 1 byte for 7bit ascii
            ' aaaabbb0 00000000 ← aaaabbb0
            UpperByte = &H0
            LowerByte = CInt(Data(i))
            Continued = 0
        Case &H80 To &HBF   ' 10xx-,  middle and lower bits of non-ascii languages
            If Continued = 1 Then   ' lower
                LowerByte = LowerByte + CInt(Data(i) And &H7F)
            Else                    ' middle
                UpperByte = UpperByte + CInt(Int((Data(i) And &H7F) / &H4))
                LowerByte = LowerByte + CInt(Data(i) And &H3) * &H40
            End If
            Continued = Continued - 1
        Case &HC0 To &HDF   ' 110x-,  goes the upper of 2 bytes for many european languages
            ' aaaabbbb ccc00000 ← aaaabb01 bbccc011 ← bbccc011 aaaabb01
            UpperByte = CInt(Int((Data(i) And &HBF) / &H4))
            LowerByte = CInt(Data(i) And &H4) * &H40
            Continued = 1
        Case Else
        'Case &HE0 To &HEF   ' 1110-,  goes the upper of 3 bytes for far east languages
        'Case &HF0 To &HFF   ' 1111-,  unknown, includes bom, do same as far east
            ' aaaabbbb ccccdddd ← aaaabb01 bbcccc01 dddd0111 ← dddd0111 bbcccc01 aaaabb01
            UpperByte = CInt((Data(i) And &HF) * &H10)
            LowerByte = &H0
            Continued = 2
        End Select
        
        If Continued = 0 Then
            If UpperByte < &H80 Then
                UnicodeWord = LowerByte + UpperByte * CInt(&H100)
            Else    ' minus bit
                UnicodeWord = LowerByte + (UpperByte - &H100) * CInt(&H100)
            End If
            pan = Array(UnicodeWord, pan)
            LengthUnicode = LengthUnicode + 1
        End If
    Next
    
    If LengthUnicode > 0 Then
        ReDim out(0 To LengthUnicode - 1)
        For i = LengthUnicode - 1 To 0 Step -1
            out(i) = pan(0)
            pan = pan(1)
        Next
    End If
    
    DecUTF8 = out
End Function

' ユニコード文字とバイト配列間の変換

Function BytesToString(Data As Variant, Optional ByVal Length As Long = -1) As String
    Dim out As String
    Dim i As Long
    
    out = ""
    If Length = -1 Then Length = UBound(Data) + 1
    For i = 0 To Length - 1
        out = out & ChrB(Data(i))
    Next
    
    BytesToString = out
End Function

Function StringToBytes(Text As String) As Byte()
    Dim out() As Byte
    Dim i As Long
    Dim Length As Long
    
    Length = LenB(Text)
    ReDim out(0 To Length - 1)
    For i = 1 To Length
        out(i - 1) = AscB(MidB(Text, i, 1))
    Next
    
    StringToBytes = out
End Function

' ユニコード文字とワード配列間の変換

Function WordsToString(Data As Variant, Optional ByVal Length As Long = -1) As String
    Dim out As String
    Dim i As Long
    
    out = ""
    If Length = -1 Then Length = UBound(Data) + 1
    For i = 0 To Length - 1
        out = out & ChrW(Data(i))
    Next
    
    WordsToString = out
End Function

Function StringToWords(Text As String) As Integer()
    Dim out() As Integer
    Dim i As Long
    Dim Length As Long
    
    Length = Len(Text)
    ReDim out(0 To Length - 1)
    For i = 1 To Length
        out(i - 1) = AscW(Mid(Text, i, 1))
    Next
    
    StringToWords = out
End Function

' ユニコード文字とロング配列間の変換

Function LongsToString(Data As Variant, Optional ByVal Length As Long = -1) As String
    Dim out As String
    Dim i As Long
    
    out = ""
    If Length = -1 Then Length = UBound(Data) + 1
    For i = 0 To Length - 1
        out = out & Chr(Data(i))
    Next
    
    LongsToString = out
End Function

Function StringToLongs(Text As String) As Long()
    Dim out() As Long
    Dim i As Long
    Dim Length As Long
    
    Length = Len(Text)
    ReDim out(0 To Length - 1)
    For i = 1 To Length
        out(i - 1) = Asc(Mid(Text, i, 1))
    Next
    
    StringToLongs = out
End Function

' ワードまとめと分離

Function BytesToWords(Data As Variant, Optional ByVal Length As Long = -1) As Integer()
    Dim WordLength As Long
    Dim i As Long
    Dim out() As Integer
    
    If Length = -1 Then Length = UBound(Data) + 1
    WordLength = Int(Length / 2) + Length Mod 2
    ' this will lose an information of the original length, odd
    If WordLength > 0 Then
        ReDim out(0 To WordLength - 1)
        For i = 0 To Length - 1 Step 2
            out(i / 2) = Data(i)
        Next
        For i = 1 To Length - 1 Step 2
            If Data(i) < &H80 Then
                out((i - 1) / 2) = out((i - 1) / 2) + Data(i) * CLng(&H100)
            Else    ' minus bit
                out((i - 1) / 2) = out((i - 1) / 2) + (Data(i) - &H100) * CLng(&H100)
            End If
        Next
    End If
    
    BytesToWords = out
End Function

Function WordsToBytes(Data As Variant, Optional ByVal Length As Long = -1) As Byte()
    Dim ByteLength As Long
    Dim i As Long
    Dim out() As Byte
    
    If Length = -1 Then Length = UBound(Data) + 1
    ByteLength = Length * 2
    If ByteLength > 0 Then
        ReDim out(0 To ByteLength - 1)
        For i = 0 To Length - 1
            out(i * 2) = CByte(Data(i) And &HFF)
            If Data(i) >= 0 Then
                out(i * 2 + 1) = CByte(Int(Data(i) / &H100))
            Else
                out(i * 2 + 1) = CByte(Int(Data(i) / &H100) + &H100)
            End If
        Next
    End If
    
    WordsToBytes = out
End Function

' バイト配列と16進ダンプ間の変換

Function EncHexDelimited(Data As Variant, Optional ByVal Length As Long = -1) As String
    Dim out As String
    Dim i As Long
    
    out = ""
    If Length = -1 Then Length = UBound(Data) + 1
    For i = 0 To Length - 1
        out = out & Hex(Data(i)) & Delimiter
    Next
    
    EncHexDelimited = out
End Function

Function DecHexDelimited(Text As String) As Byte()
    Dim Line As Variant
    Dim ByteData As Variant
    Dim out() As Byte
    Dim pan As Variant
    Dim Counter As Long
    Dim i As Long
    
    pan = Empty
    Counter = 0
    
    For Each Line In Split(Text, LineFeed)
        For Each ByteData In Split(Line, Delimiter)
            If ByteData <> "" Then
                pan = Array(CByte("&H" & ByteData), pan)
                Counter = Counter + 1
            End If
        Next
    Next
    
    If Counter > 0 Then
        ReDim out(0 To Counter - 1)
        For i = Counter - 1 To 0 Step -1
            out(i) = pan(0)
            pan = pan(1)
        Next
    End If
    
    DecHexDelimited = out
End Function

Function DecHexDelimitedL(Text As String) As Long()
    Dim Line As Variant
    Dim LongData As Variant
    Dim out() As Long
    Dim pan As Variant
    Dim Counter As Long
    Dim i As Long
    
    pan = Empty
    Counter = 0
    
    For Each Line In Split(Text, LineFeed)
        For Each LongData In Split(Line, Delimiter)
            If LongData <> "" Then
                pan = Array(CLng("&H" & LongData), pan)
                Counter = Counter + 1
            End If
        Next
    Next
    
    If Counter > 0 Then
        ReDim out(0 To Counter - 1)
        For i = Counter - 1 To 0 Step -1
            out(i) = pan(0)
            pan = pan(1)
        Next
    End If
    
    DecHexDelimitedL = out
End Function

' 改行統一 (vbCrLF, vbLF, vbCr を指定のものに統一する)

Function FixLineFeed(Text As String, Optional NewLineFeed As String = vbCrLf) As String
    Dim pan As Collection
    Dim out As String
    Dim Line As Variant
    Dim MoreLine As Variant
    Dim MoreEnd As Long
    Dim i As Long
    
    Set pan = New Collection
    For Each Line In Split(Text, vbLf)
        If Line = "" Then
            pan.Add ""
        Else
            MoreLine = Split(Line, vbCr)
            MoreEnd = UBound(MoreLine)
            If MoreLine(MoreEnd) = "" Then MoreEnd = MoreEnd - 1
            For i = 0 To MoreEnd
                pan.Add MoreLine(i)
            Next
        End If
    Next
    Do While pan.Count > 0
        out = out & pan(1) & NewLineFeed
        pan.Remove 1
    Loop
    
    If out = "" Then
        FixLineFeed = ""
    Else
        FixLineFeed = Left(out, Len(out) - Len(NewLineFeed))
    End If
End Function

'}}}

```

### Result ###

```
FF FE 
EF BB BF 
FEFF 
OK
46 0 72 0 65 0 65 0 20 0 EA 81 31 75 
46 72 65 65 20 81EA 7531 
46 72 65 65 20 FFFF8EA9 FFFF9752 
OK
OK
OK
46 72 65 65 20 E8 87 AA E7 94 B1 
46 72 65 65 20 81EA 7531 
Free 自由
OK
4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30 
OK
304A 5E30 308A 306A 3055 3044 307E 305B 
OK
AB##cd##EF####ghij####KLM
AB
cd
EF

ghij

KLM
41 0 42 0 D 0 A 0 63 0 64 0 D 0 45 0 46 0 A 0 D 0 67 0 68 0 69 0 6A 0 A 0 A 0 4B 0 4C 0 4D 0 
41 0 42 0 D 0 A 0 63 0 64 0 D 0 A 0 45 0 46 0 D 0 A 0 D 0 A 0 67 0 68 0 69 0 6A 0 D 0 A 0 D 0 A 0 4B 0 4C 0 4D 0 
```
