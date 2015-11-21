# Introduction #

  * bad vba: variant denies to accept object sometimes

## 概要 ##
  * 失敗マクロ: VBAの Variant 型は、たまに Object 型を受け付けない

# Details #

  * the Variant is a useful data type that accept any other types include an Object type.
  * but the conversion raises a runtime error 438 sometimes.
  * we can handle this by defining a variant variable explicitly and setting it to refer the object.

## 説明 ##
  * バリアント型は他のすべての型を受け入れる便利な型だ。オブジェクト型も受け入れる。
  * ところが、実行時にエラー 438 を出して、変換に失敗するときがある。
  * そんな時は、バリアント型の変数を１個明確に定義して、そこに Set 文で参照の代入をしてやると解決する。

# Bad Code #

```
'workbook
'  name;bad_objects_array_variant.xls


'class
'  name;Class1
'{{{
Option Explicit

Public Sub OK()
    Give Me
End Sub

Public Sub NG()
    On Error Resume Next
    Gives Me
    Debug.Print Err.Number, Err.Description
End Sub

Public Sub YES()
    Dim x As Variant
    Set x = Me
    Gives x
End Sub
'}}}

'module
'  name;Module1
'{{{
Option Explicit

Public Sub Give(Who As Variant)
    Debug.Print TypeName(Who)
End Sub

Public Sub Gives(ParamArray Whos() As Variant)
    Dim Who As Variant
    For Each Who In Whos
        Give Who
    Next
End Sub

Sub test()
    Dim x As Class1
    Set x = New Class1
    x.OK
    x.NG
    x.YES
    Set x = Nothing
End Sub
'}}}


```

### Result ###

```
Class1
 438          オブジェクトは、このプロパティまたはメソッドをサポートしていません。
Class1
```