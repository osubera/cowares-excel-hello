# Introduction #

  * bad vba: `Cells.SpecialCells` ignores a sigle cell searching

## 概要 ##
  * 失敗マクロ: VBAの `Cells.SpecialCells` はセル１個だけを検索できない

# Details #

  * `SpecialCells` is a useful method to switch procedures by cells formulas and values.
  * when a single cell is given to this method, it ignores the given range. for example, when `Sheet1.Range("C2")` is given, it searches in whole `Sheet1.Cells`.
  * this is not a bug, but a design. commonly in every version of excel.
  * to go with this busybody, and dare we want to search in a very single cell itself, there's a tricky solution. a `Sheet1.Range("C2,C2")` has an appearance of 2 areas and 2 total cells, and it works as a sigle cell searching.
  * anyway we have to care this, particularly when using a cell range generated dynamically. or we fail.

## 説明 ##
  * `SpecialCells` は、セル数式の状態により異なる処理をするために欠かせない。
  * 検索対象として、１つだけのセル、 `Sheet1.Range("C2")` など、を指定すると、セルからの検索でなく `Sheet1.Cells` 全体を検索対象に切り替える。
  * この動作はバグでなく仕様で、全バージョンで共通に見られる。
  * 自動的なおせっかいを回避して、単一セルからの検索を指定したい場合、 `Sheet1.Range("C2,C2")` のように、同一セルの２重エリアで、あたかもセル数が２のように見せかけるという抜け道が使える。
  * いずれにせよ、動的に作られたセル範囲に対して `SpecialCells` で絞り込む場合、セル１個だけを特別扱いしなければ、間違った動作になる。

# Bad Code #

```
Option Explicit

Sub test_all()
    check1
    check2
    check3
    check4
End Sub

Sub check1()
    With Me.UsedRange
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
End Sub

Sub check2()
    On Error GoTo NoCellsFound
    
    With Me.Range("C2:E4")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    With Me.Range("C5:E6")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    
    Exit Sub
    
NoCellsFound:
    Debug.Print Err.Number, Err.Description, Err.Source
    Resume Next
End Sub

Sub check3()
    On Error GoTo NoCellsFound
    
    With Me.Range("C2:C2")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    With Me.Range("C5:C5")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    
    Exit Sub
    
NoCellsFound:
    Debug.Print Err.Number, Err.Description, Err.Source
    Resume Next
End Sub

Sub check4()
    On Error GoTo NoCellsFound
    
    ' TRICK!
    With Me.Range("C2,C2")
        Debug.Print .Cells.Count, .Areas.Count, .Address
    End With
    
    With Me.Range("C2,C2")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    With Me.Range("C5,C5")
        Debug.Print .SpecialCells(xlCellTypeConstants).Address(False, False, xlA1, False), "Constants"
        Debug.Print .SpecialCells(xlCellTypeFormulas).Address(False, False, xlA1, False), "Formulas"
    End With
    
    Exit Sub
    
NoCellsFound:
    Debug.Print Err.Number, Err.Description, Err.Source
    Resume Next
End Sub

```

### Result ###

```
check1()

C2:E4,B2:B6   Constants
C5:E6         Formulas

check2()

C2:E4         Constants
 1004         該当するセルが見つかりません。 Microsoft Excel
 1004         該当するセルが見つかりません。 Microsoft Excel
C5:E6         Formulas

check3()

C2:E4,B2:B6   Constants
C5:E6         Formulas
C2:E4,B2:B6   Constants
C5:E6         Formulas

check4()

 2             2            $C$2,$C$2
C2            Constants
 1004         該当するセルが見つかりません。 Microsoft Excel
 1004         該当するセルが見つかりません。 Microsoft Excel
C5            Formulas

```