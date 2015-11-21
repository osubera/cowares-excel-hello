# Introduction #

  * how to use flexible arrays (collection) in vba

## 概要 ##
  * VBAで伸縮可能な配列（コレクション）を使う

# Details #

  * important: see also [bad\_collection](bad_collection.md) .
  * ReDim is a standard solution to handle this, but too much use of ReDim costs.
  * the Collection object can do this.

## 説明 ##
  * [bad\_collection](bad_collection.md) に重要な追加情報あり。
  * 通常は ReDim で良いが、頻繁に変えるには ReDim は重い。
  * Collection object を使うとよい。

# Code #

```
    Dim x As Variant
    Dim a As Collection
    Set a = New Collection
    
    a.Add "1st"
    a.Add "2nd"
    
    Debug.Print a(1)
    a.Add "3rd", Before:=1
    Debug.Print a(1), a(3)
    a.Add "4th", After:=2
    Debug.Print a(3)
    
    Debug.Print "loop by index"
    For x = 1 To a.Count
        Debug.Print a(x)
    Next
    
    a.Remove 1
    
    Debug.Print "loop by iterator"
    For Each x In a
        Debug.Print x
    Next
    Set a = Nothing
```

results

```
1st
3rd           2nd
4th
loop by index
3rd
1st
4th
2nd
loop by iterator
1st
4th
2nd
```