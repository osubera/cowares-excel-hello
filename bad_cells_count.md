# Introduction #

  * bad vba: Cells.Count fails on excel 2007 or later

## 概要 ##
  * 失敗マクロ: VBAの Cells.Count はエクセル2007以降でエラーになる

# Details #

  * the total cell number increased to  2<sup>34</sup> from legacy 2<sup>24</sup> in excel 2007.
  * `Sheet1.Cells.Count` is not safe now. it raises `Error 6` overflow, because the Count property is designed as a 32bit integer.
  * a new propery `Cells.CountLarge` is added. and it returns a 64bit signed integer.
  * because there're no such variable types declared in the VBA, this property is not useful to calculate. it may raise an error on implicit data type conversion.
  * using a Variant variable with explicit `CDec` conversion is a simple solution.
  * converting into a Double is an another way, when calculations of approximate values are important than precise values.
  * `Cells.CountLarge` uses late-binding. it raises a run time error, not at compile time in excel 2000. so we can switch dynamically by the `Application.Version`.
  * when the required value is not a cell counts number itself but a true-or-false value against a threshold value, it is better to try and catch the error to judge by a simple code.
  * when running an excel 2007 in a legacy compatible mode, there're no errors because the total amount of cells is also shrinked as legacy.

## 説明 ##
  * エクセル2007では、セルの総数が、従来の 2<sup>24</sup> 個から 2<sup>34</sup> 個に増えた。
  * `Sheet1.Cells.Count` などを安易に使うと、 Count プロパティは 32bit integer 型なので、 'Error 6' オーバーフローエラーになる。
  * 64bit signed integer を返す `Cells.CountLarge` というプロパティが新しく設けられた。
  * VBA では、この型をサポートしないため、そのまま計算式に使うと、暗黙の型変換でエラーが起こる。
  * Variant 型変数に `CDec` 変換して代入すると、とりあえず使える。
  * 正確な値より、概算でいいから計算を重視するなら、 Double 型を使うのも手。
  * `Cells.CountLarge` は、エクセル2000ではコンパイルエラーでなく実行時エラーになるので、 'Application.Version' で判定し動的に切り替えることもできる。
  * セル数を知りたい目的が何らかの真偽判定のときは、それを直接行う関数を独立させ、エラー処理で切り抜ける方が、丈夫なコードになる。
  * 互換モードでは、セルの数も従来と同じなので、エラーが出ない。

# Bad Code #

```
'module
'  name;BadCellsCount
'{{{
Option Explicit

Sub test_all()
    ThisSimpleCodeFailsOnExcel2007
    ThisCodeIsForExcel2007AndAbove
    OtherWays
End Sub

Sub ThisSimpleCodeFailsOnExcel2007()
    On Error Resume Next
    Debug.Print ActiveSheet.Cells.Count
    If Err.Number = 0 Then Exit Sub
    
    ' expect Error 6 for Excel 2007 and later
    Debug.Print Err.Number, Err.Description
End Sub

Sub ThisCodeIsForExcel2007AndAbove()
    On Error Resume Next
    Debug.Print ActiveSheet.Cells.CountLarge
    If Err.Number = 0 Then Exit Sub
    
    ' expect Error 438 for Excel 2003 and earlier
    Debug.Print Err.Number, Err.Description
End Sub

Sub OtherWays()
    Debug.Print IsASingleCell(ActiveCell)
    Debug.Print IsASingleCell(ActiveSheet.Cells)

    Debug.Print CellsCountLarge(ActiveCell)
    Debug.Print CellsCountLarge(ActiveSheet.Cells)
End Sub

Private Function IsASingleCell(Target As Range) As Boolean
    On Error GoTo MayFailOnExcel2007
    
    IsASingleCell = (Target.Cells.Count = 1)
    Exit Function
    
MayFailOnExcel2007:
    If Err.Number = 6 Then
        ' overflowed, means very large, larger than 1, maybe
        IsASingleCell = False
        Exit Function
    Else
        Err.Raise Err.Number
    End If
End Function

Private Function CellsCountLarge(Target As Range) As Variant
    If Application.Version >= 12 Then
        CellsCountLarge = CDec(Target.Cells.CountLarge)
    Else
        CellsCountLarge = Target.Cells.Count
    End If
End Function
'}}}

```

### Result ###

```
Excel 2000 Results

 16777216 
 438          オブジェクトは、このプロパティまたはメソッドをサポートしていません。
True
False
 1 
 16777216 

Excel 2007 Results
(legacy compatible mode)

 16777216 
16777216
True
False
 1 
 16777216 

Excel 2007 Results
(native mode)

 6            オーバーフローしました。
17179869184
True
False
 1 
 17179869184 

```