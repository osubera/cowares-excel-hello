# Introduction #

  * implements NumberString worksheet function in vba

## 概要 ##
  * VBAで NumberString ワークシート関数を実装する

# Details #

  * the `NumberString` worksheet function is undocumented and available only at cells.
  * it is said this function is to keep compatibilities with a Japanese Lotus-123.
  * the function convert a number into a Japanese business style Kanji string.

## 説明 ##
  * `NumberString` ワークシート関数は、ヘルプに載ってないし、セルでしか使えない。
  * この関数は、日本語版ロータス123との互換性を保つためにあるらしい。
  * この関数は、数字を日本の商慣習に合わせた漢数字に変換する。

# Code #

```
'workbook
'  name;number_string.xls

'require

'worksheet
'  name;Sheet1

'cells-formula
'  address;A1:M55
'         ;
'         ;Buitin Worksheet Function NumberString
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;VBA User Function NumberStringVBA
'         ;
'         ;
'         ;
'         ;
'         ;
'         ;0
'         ;1
'         ;2
'         ;3
'         ;4
'         ;
'         ;
'         ;0
'         ;1
'         ;2
'         ;3
'         ;4
'         ;-10
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-10
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-9
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-9
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-8
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-8
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-7
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-7
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-6
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-6
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-5
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-5
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-4
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-4
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-3
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-3
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-2
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-2
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;-1
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;-1
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;0
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;0
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;2
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;2
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;3
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;3
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;4
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;4
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;5
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;5
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;6
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;6
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;7
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;7
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;8
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;8
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;9
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;9
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;11
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;11
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;12
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;12
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;13
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;13
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;14
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;14
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;15
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;15
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;16
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;16
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;17
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;17
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;18
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;18
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;19
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;19
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;20
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;20
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;21
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;21
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;200
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;200
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;2000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;2000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;20000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;20000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;10000000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;10000000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;100000000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;100000000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1000000000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1000000000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;1234567000000000000
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;=NUMBERSTRING(RC1,R2C)
'         ;
'         ;1234567000000000000
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)
'         ;=NUMBERSTRINGvba(RC1,R2C)


'module
'  name;NumberString
'{{{
Option Explicit

Public Function NumberStringVBA(Number As Variant, Optional Style As Long = 1) As String
    Dim x As Variant
    Dim Keta As Long
    Dim Atai As Long
    Dim AtaiMan As Long
    Dim AtaiChou As Variant
    Dim Sign As String
    Dim out As String
    
    x = CDec(Val(Number))
    Sign = IIf(x < 0, "△", "")
    x = Int(Abs(x))
    Keta = 0
    AtaiChou = 0
    out = ""
    
    Do Until x = 0
        Atai = DecMod(x, 10)
        AtaiMan = IIf(Keta Mod 4 = 0, DecMod(x, 10000), 0)
        If Keta = 12 And x >= 10000 Then AtaiChou = x
        out = HabuiteYoshi(KanjiNumber(Atai, Style), KetaNumber(Keta, Style), Atai, AtaiMan, AtaiChou, Keta, Style) & out
        Keta = Keta + 1
        x = Int(x / 10)
    Loop
    If out = "" Then out = KanjiNumber(0, Style)
    
    NumberStringVBA = Sign & out
End Function

Public Function KanjiNumber(Number As Long, Optional Style As Long = 1) As String
    Const Style1 = "〇一二三四五六七八九"
    Const Style2 = "〇壱弐参四伍六七八九"
    Select Case Style
    Case 1, 3
        KanjiNumber = Mid(Style1, Number + 1, 1)
    Case 2
        KanjiNumber = Mid(Style2, Number + 1, 1)
    End Select
End Function

Public Function KetaNumber(Number As Long, Optional Style As Long = 1) As String
    Const Style1 = " 十百千万十百千億十百千兆十百千"
    Const Style2 = " 拾百阡萬拾百阡億拾百阡兆拾百阡"
    Select Case Style
    Case 1
        KetaNumber = Trim(Mid(Style1, Number + 1, 1))
    Case 2
        KetaNumber = Trim(Mid(Style2, Number + 1, 1))
    End Select
End Function

Public Function HabuiteYoshi(AtaiMoji As String, KetaMoji As String, _
                Atai As Long, AtaiMan As Long, AtaiChou As Variant, _
                Keta As Long, Style As Long) As String
    
    If AtaiChou = 0 Then
        HabuiteYoshi = AtaiMoji & KetaMoji
        Select Case Atai
        Case 0
            If Style <> 3 Then
                If AtaiMan = 0 Then
                    HabuiteYoshi = ""
                Else
                    HabuiteYoshi = KetaMoji
                End If
            End If
        Case 1
            If Style = 1 Then
                If Keta Mod 4 <> 0 Then
                    HabuiteYoshi = KetaMoji
                End If
            End If
        End Select
    ElseIf Keta = 12 Then
        HabuiteYoshi = AtaiMoji & KetaMoji
    Else    ' Keta > 12
        HabuiteYoshi = AtaiMoji
    End If
End Function

Public Function DecMod(x As Variant, y As Variant) As Variant
    ' VB buitin MOD will overflow when x > 2.147e9 for y=10
    DecMod = x - y * Int(x / y)
End Function
'}}}

```