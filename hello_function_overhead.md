

# Introduction #

  * how much overhead is there in calling a function in vba

## 概要 ##
  * VBAで関数呼び出しのオーバーヘッド負荷はどれぐらいあるか

# Details #

  * i know there is an overhead in calling a function, but i don't know the amount of loss. let me measure it.
  * the results strongly encourages to choose calling a function when it helps reading codes.
  * there exists a loss truely, but the amount is tiny.

## 説明 ##
  * 関数呼び出しでオーバーヘッドがあるのは知っていても、実際どれぐらいなのかは知らない。測ってみよう。
  * この結果なら、コードが読みやすくなる理由で関数を使うのに躊躇する必要はなさそうだ。
  * 確かにロスはあるが、非常に小さい。

# Code #

```
'module
'  name;Module1
'{{{
Option Explicit

Function Cons(x As Long, y As Long) As Variant
    Cons = Array(x, y)
End Function

Function Cons2(x As Long, y As Long) As Long()
    Dim out(0 To 1) As Long
    out(0) = x
    out(1) = y
    Cons2 = out
End Function

Function Car(x As Variant) As Long
    Car = x(0)
End Function

Function Cdr(x As Variant) As Long
    Cdr = x(1)
End Function

Sub test_measure_write()
    Const Repeat = 100000
    Dim i As Long
    Dim z As Variant
    Dim BeginAt As Single
    Dim EndAt As Single
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = Array(Rnd * 100, Rnd * 100)
    Next
    EndAt = Timer()
    Debug.Print "Direct Write", EndAt - BeginAt
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = Cons(Rnd * 100, Rnd * 100)
    Next
    EndAt = Timer()
    Debug.Print "Variant Cons", EndAt - BeginAt
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = Cons2(Rnd * 100, Rnd * 100)
    Next
    EndAt = Timer()
    Debug.Print "Long() Cons", EndAt - BeginAt
End Sub

Sub test_measure_read()
    Const Repeat = 1000000
    Dim i As Long
    Dim x As Variant
    Dim z As Variant
    Dim BeginAt As Single
    Dim EndAt As Single
    
    x = Array(123, 456)
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = x(0)
    Next
    EndAt = Timer()
    Debug.Print "Direct Select", EndAt - BeginAt
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = Car(x)
    Next
    EndAt = Timer()
    Debug.Print "Car", EndAt - BeginAt
    
    BeginAt = Timer()
    For i = 1 To Repeat
        z = Cdr(x)
    Next
    EndAt = Timer()
    Debug.Print "Cdr", EndAt - BeginAt
End Sub
'}}}


```

# Results #


```
Direct Write   0.890625 
Variant Cons   1.15625 
Long() Cons    1.0625 
Direct Select  0.6171875 
Car            1.5 
Cdr            1.476563
```