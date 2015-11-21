# Introduction #

  * how to convert euc-jp and iso-2022-jp to shift\_jis at a low level in vba
  * today we have other built-in functions as described at [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) , no needs to these old fashoned code.
  * i just post it here, to remember the rule of basic conversion.

## 概要 ##
  * VBAで euc-jp と iso-2022-jp を shift\_jis に低レベルな変換をする
  * 今では [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) で使ったような組み込み関数を使うので、こんな時代物を使うことはない。
  * 変換規則の記録として、ここに投稿しておく。

# Details #

  * 2 functions for euc, basic one is to convert its 1 character (2 byte) value into a 2 byte long value, and another is to convert byte arrays into a unicode string.
  * 2 functions also for jis, like above.
  * the purpose is to use it as a VBA string, so our goal is to get a unicode string
  * Chr() function returns a unicode string for a given long value of an ANSI (Shift\_JIS) code. so converting a sjis value into a unicode string is easy.
  * this converter is not completed, just what as-is working.
  * test module uses the result of [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) , we define constants for each character set as a byte array.

## 説明 ##
  * euc の１文字（２バイト）を sjis の２バイトに変換する基本関数と、それをつかってバイト列をユニコード文字列へと変換する関数がある。
  * jis も同様に２種類の関数を持つ。
  * VBA の文字列として使うために、ユニコード文字列が最終形になる。
  * Chr() 関数に ANSI コード（シフトJIS）をロング値で渡せば、ユニコード文字を返すので、 sjis コードのロング値からユニコード文字への変換は、この関数を使うだけだ。
  * 変換関数は完成されたものでなく、動いていたものを持ってきただけだ。
  * テストモジュールでは、 [hello\_charset\_adodb\_stream](hello_charset_adodb_stream.md) の結果を利用して、それぞれの文字コードに対応したバイト配列を定数で定義して利用する。

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. converter functions are located at ToShiftJis module.
  1. test1() in test module is an executable test.

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. ToShiftJis 標準モジュールに変換関数がある
  1. test 標準モジュールの test1() が実行可能なテストコード

# Code #

```
'workbook
'  name;hello_charset_euc_jis_sjis.xls

'require

'worksheet
'  name;Sheet1


'module
'  name;ToShiftJis
'{{{
Option Explicit

Public Function Ajis2sjis(ByVal aH As Long, ByVal aL As Long) As Long
' JISコードの第一、第二バイトをShift JISの2バイトに変換する。
    Const L256 As Long = 256
    aH = aH + &HA1
    aL = aL + &H7E
    If aL >= L256 Then
        aL = aL - L256
        aH = aH + 1
    End If
    If (aH Mod 2) = 0 Then
        If aL < &HDE Then
            aL = aL - 1
        End If
        aL = aL - &H5E
    End If
    aH = Int(aH / 2) Xor &HE0
    Ajis2sjis = aL + aH * L256
End Function

Public Function Aeuc2sjis(ByVal aH As Long, ByVal aL As Long) As Long
' eucコードの第一、第二バイトをShift JISの2バイトに変換する。
    Aeuc2sjis = Ajis2sjis(aH And &H7F, aL And &H7F)
End Function

Public Function Euc2Sjis(ByRef binArray As Variant, ByRef ArrayLength As Long, _
                    ByRef LfCount As Long) As String
' Binary 配列でもらったeuc文字列を、Shift JIS文字列で返す。
' データ末尾に第一バイトの端数があった場合、ArrayLength=1を返し、
' 同時にArray先頭に第一バイトをセットする。
' 端数が無ければ、ArrayLength=0を返し、Arrayには何もしない。
' コード0は、長さ0の文字列として返す。(EOF用)
' 改行は 0a に統一。その際Arrayデータを変更することもある。
' 処理した改行数を lfCount に返す。
    Dim i As Long
    Dim strB As String
    Dim Kanji As Boolean
    Dim Cr As Boolean
    Kanji = False
    Cr = False
    strB = ""
    LfCount = 0
    For i = 0 To ArrayLength - 1
        If Cr Then
            strB = strB & Chr(&HA)
            LfCount = LfCount + 1
            If binArray(i) = &HA Then
                binArray(i) = 0
            End If
            Cr = False
        End If
        If binArray(i) = &HA Then
            LfCount = LfCount + 1
        End If
        If Kanji Then
            strB = strB & Chr(Aeuc2sjis(binArray(i - 1), binArray(i)))
            Kanji = False
        Else
            If binArray(i) >= &H80 Then
                Kanji = True
            ElseIf binArray(i) = &HD Then
                Cr = True
            ElseIf binArray(i) > 0 Then
                strB = strB & Chr(binArray(i))
            End If
        End If
    Next
    If Kanji Or Cr Then
        binArray(0) = binArray(ArrayLength - 1)
        ArrayLength = 1
    Else
        ArrayLength = 0
    End If
    Euc2Sjis = strB
End Function

Public Function Jis2Sjis(ByRef binArray As Variant, ByRef ArrayLength As Long, _
                    ByRef LfCount As Long) As String
' Binary 配列でもらったJIS文字列を、Shift JIS文字列で返す。
' データ末尾に第一バイトの端数があった場合、ArrayLength=1を返し、
' 同時にArray先頭に第一バイトをセットする。
' Esc情報はstaticで内部に持ってしまう。
' 端数が無ければ、ArrayLength=0を返し、Arrayには何もしない。
' コード0は、長さ0の文字列として返す。(EOF用)
' 改行は 0a に統一。その際Arrayデータを変更することもある。
' 処理した改行数を lfCount に返す。
    Dim i As Long
    Dim strB As String
    Dim Kanji As Boolean
    Dim Cr As Boolean
    Static strEsc As String
    Static iEsc  As Long
    Kanji = False
    Cr = False
    strB = ""
    LfCount = 0
    For i = 0 To ArrayLength - 1
        ' 改行関連
        If Cr Then
            strB = strB & Chr(&HA)
            LfCount = LfCount + 1
            If binArray(i) = &HA Then
                binArray(i) = 0
            End If
            Cr = False
        End If
        If binArray(i) = &HA Then
            LfCount = LfCount + 1
        End If
        ' JISモード切替
        Select Case binArray(i)
        Case 0      ' nop
        Case 27     ' Esc
            iEsc = 101
        Case Else
            Select Case iEsc
            Case 1      'Esc(J ローマ字?
            Case 2      'Esc$@ JIS1978
            Case 3      'Esc$B JIS1983
                If Kanji Then
                    strB = strB & Chr(Ajis2sjis(binArray(i - 1), binArray(i)))
                End If
                Kanji = Not Kanji
            Case 101    'after Esc
                strEsc = Chr(binArray(i))
                iEsc = 102
                Kanji = False
            Case 102    'after Esc+
                strEsc = strEsc & Chr(binArray(i))
                Select Case strEsc
                Case "(B"   ' Ascii
                    'strB = strB & vbLf & "!!!debug!!!" & vbLf & "JIS Mode 1" & vbLf
                    iEsc = 0
                Case "(J"   ' JISローマ字
                    strB = strB & vbLf & "!!!debug!!!" & vbLf & "JIS Mode 2" & vbLf
                    LfCount = LfCount + 3   ' for debug
                    iEsc = 1
                Case "$@"   ' JIS1978
                    strB = strB & vbLf & "!!!debug!!!" & vbLf & "JIS Mode 3" & vbLf
                    LfCount = LfCount + 3   ' for debug
                    iEsc = 2
                Case "$B"   ' JIS1983
                    'strB = strB & vbLf & "!!!debug!!!" & vbLf & "JIS Mode 4" & vbLf
                    iEsc = 3
                Case Else
                    strB = strB & vbLf & "!!!debug!!!" & vbLf & strEsc & " is unknown JIS Mode" & vbLf
                    LfCount = LfCount + 3   ' for debug
                    iEsc = 0
                End Select
            Case Else   'Esc(B Ascii
                Select Case binArray(i)
                Case &HD
                    Cr = True
                Case Else
                    strB = strB & Chr(binArray(i))
                End Select
            End Select
        End Select
    Next
    ' 繰越
    If Kanji Or Cr Then
        binArray(0) = binArray(ArrayLength - 1)
        ArrayLength = 1
    Else
        ArrayLength = 0
    End If
    Jis2Sjis = strB
End Function


'}}}

'module
'  name;test
'{{{
Option Explicit

' お帰りなさいませ
Const TEXT_unicode = "4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30"
Const TEXT_sjis = "82 A8 8B 41 82 E8 82 C8 82 B3 82 A2 82 DC 82 B9"
Const TEXT_euc = "A4 AA B5 A2 A4 EA A4 CA A4 B5 A4 A4 A4 DE A4 BB"
Const TEXT_jis = "1B 24 42 24 2A 35 22 24 6A 24 4A 24 35 24 24 24 5E 24 3B 1B 28 42"

Function TextToByteArray(Text As String) As Byte()
    Dim x As Variant
    Dim i As Long
    Dim out() As Byte
    
    x = Split(Text, " ")
    ReDim out(0 To UBound(x))
    For i = 0 To UBound(x)
        out(i) = CByte("&H" & x(i))
    Next
    
    TextToByteArray = out
End Function

Function DoConvert(ByRef Enc, ByRef ByteBuffer() As Byte, ByRef ByteLength As Long) As String
    Dim out As String
    Dim LfCount As Long
    Dim i As Long
    Const L256 As Long = 256
    
    out = ""
    Select Case Enc
    Case 0  ' sjis
        ' cheat in this test, to know every character in the text is 2 byte pair.
        ' and the function Chr() return a Unicode Character for a given Shift_JIS long code.
        For i = 0 To ByteLength / 2 - 1
            out = out & Chr(L256 * ByteBuffer(i * 2) + ByteBuffer(i * 2 + 1))
        Next
    Case 1  ' euc
        out = Euc2Sjis(ByteBuffer, ByteLength, LfCount)
        ByteLength = LenB(out)
    Case 2  ' jis
        out = Jis2Sjis(ByteBuffer, ByteLength, LfCount)
        ByteLength = LenB(out)
    End Select
    
    For i = 0 To ByteLength - 1
        ByteBuffer(i) = AscB(MidB(out, i + 1, 1))
    Next
    
    DoConvert = out
End Function

Sub test1()
    Dim ByteBuffer(0 To 63) As Byte
    Dim ByteLength As Long
    Dim x As Variant
    Dim i As Long
    Dim out As String
    Dim Enc As Long
    Dim Charsets As Variant
    
    Charsets = Array(TEXT_sjis, TEXT_euc, TEXT_jis)
    For Enc = 0 To UBound(Charsets)
    
        x = TextToByteArray(CStr(Charsets(Enc)))
        ByteLength = UBound(x) + 1
        For i = 0 To ByteLength - 1
            ByteBuffer(i) = x(i)
        Next
        
        Debug.Print DoConvert(Enc, ByteBuffer, ByteLength)
        
        out = ""
        For i = 0 To ByteLength - 1
            out = out & Hex(ByteBuffer(i)) & " "
        Next
        Debug.Print out
        Debug.Print IIf(Trim(out) = TEXT_unicode, "OK", "NG")
        
    Next
End Sub
'}}}
```

### Result ###

```
お帰りなさいませ
4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30 
OK
お帰りなさいませ
4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30 
OK
お帰りなさいませ
4A 30 30 5E 8A 30 6A 30 55 30 44 30 7E 30 5B 30 
OK
```
