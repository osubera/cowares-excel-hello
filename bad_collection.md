# Introduction #

  * bad vba: collection object is not for thousands of data

## 概要 ##
  * 失敗マクロ: VBAの Collection オブジェクトは大量のデータに向いてない

# Details #

  * i expected the Collection Object had a data structure of a flexible list, when i wrote the page [hello\_array\_collection](hello_array_collection.md) . i was wrong.
  * the speed of `Add` and `Remove` depends the number of data in the Collection.
  * Removing item located near the last of large amount of data is especially slow.
  * this means the Collection does something like `ReDim`, and the performance is worse.
  * i found this when i tried to implement the quick sort for Collection Object on [hello\_sort\_collection](hello_sort_collection.md) .
  * still the Collection is useful for easy operations.

## 説明 ##
  * [hello\_array\_collection](hello_array_collection.md) のページを書いたとき、 Collection オブジェクトは、柔軟なリストのデータ構造だと期待していた。これは間違い。
  * `Add` と `Remove` の速度は、 Collection のデータ数に依存する。
  * データの数が多い Collection の最後のあたりで削除をすると極端に遅い。
  * どうやら Collection は `ReDim` と似たことをやっていて、しかもさらに遅いようだ。
  * [hello\_sort\_collection](hello_sort_collection.md) で Collection 用のクイックソートを実装してみたときに気がついた。
  * 操作が簡単という点では Collection は便利なのだが。

# Bad Code #

```
'workbook
'  name;bad_collection.xls


'module
'  name;Module1
'{{{
Option Explicit

#Const P = "A"

Sub measure_Collection()
    Dim x As Collection
    Dim BeginAt As Single
    Dim EndAt As Single
    Dim i As Long
    Dim j As Long
    Dim Repeat As Long
    Dim Length As Long
    
    Set x = New Collection
    
    For j = 1 To 4
        Length = 10 ^ j
        Do While x.Count < Length
            x.Add Space(8192)
        Loop
        
        Repeat = 1000
        BeginAt = Timer()
        For i = 1 To Repeat
            #If P = "A" Then
                x.Add Space(8192), Before:=1
                x.Remove 2
            #ElseIf P = "B" Then
                x.Add Space(8192)
                x.Remove 2
            #ElseIf P = "C" Then
                x.Add Space(8192), Before:=1
                x.Remove Length - 1
            #ElseIf P = "D" Then
                x.Add Space(8192)
                x.Remove Length - 1
            #Else
                x.Count
            #End If
        Next
        EndAt = Timer()
        Debug.Print Length, EndAt - BeginAt
    Next
    
    Set x = Nothing
End Sub
'}}}



```

### Result ###

```
A
 10            8.007813E-02 
 100           6.835938E-02 
 1000          0.0703125 
 10000         0.0703125 
B
 10            0.1582031 
 100           0.2089844 
 1000          0.2011719 
 10000         0.1992188 
C
 10            0.1601563 
 100           0.2109375 
 1000          0.3105469 
 10000         3.746094 
D
 10            8.007813E-02 
 100           8.984375E-02 
 1000          0.1601563 
 10000         3.796875 
E
 10            0 
 100           0 
 1000          0 
 10000         0 
```