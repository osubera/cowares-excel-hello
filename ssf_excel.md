

# notices for Microsoft Excel #

We mostly target Microsoft Excel Books here, so we have some additional rules to work better with them.

  1. linefeed is `vbCrLf (x0d0a)`
  1. an ssf line begins with `' (x2c)`
  1. a void line contains only white spaces or blank
  1. escaped lines begin with `'{{{` and end with `'}}}`
  1. no localized characters in ssf keys, ascii 7bit alphabets and numbers only

## マイクロソフト エクセル に特化した事項 ##

ここではエクセル VBA の Visual Basic Editor ウィンドウに、 ssf を貼り付ける上で役に立つ表記上のルールをあげておく。

  1. 改行記号は `vbCrLf (x0d0a)`
  1. ssf 有効行は、先頭が `' (x2c)` の行
  1. 無効行は空行
  1. 行エスケープは、 `'{{{` で開始、 `'}}}` で終了。
  1. ssf キーは、欧米版を考慮して半角英数字だけにしておく。

# ssf block definitions for Excel #

### idenfity ssf lines ###

  1. a line begins with `' (x2c)`
  1. remove the 1st character and continuous `spaces(x20)` and `tabs(x09)` from the 2nd
  1. remove the linefeed character at the line end
  1. escaped lines are not effected above all rules
  1. an escaped line block is regarded as a single line of anonymous data

### ssf key line ###

  1. an ssf key line is the 1st line of a ssf block
  1. remove continuous `spaces(x20)` and `tabs(x09)` at the line end
  1. the ssf key is all remained

### ssf value line ###

  1. ssf value lines are the 2nd line and followings of a ssf block
  1. an ssf value line is splited into 2 parts of name and data, by a delimiter
  1. the delimiter is `; (x3b)` , we take the earlier one only
  1. an anonymous data begins with the delimiter
  1. different format of data parts, for the different ssf key

## エクセル向けの ssf ブロック定義 ##

### ssf 行の識別 ###

  1. 行の先頭文字が `' (x2c)` の行
  1. 先頭文字に続く連続した `空白(x20)` と `タブ(x09)` は先頭文字とともに除去
  1. 行末改行文字は除去
  1. 上記すべてに優先して、行エスケープ処理を行う
  1. 行エスケープは、処理上は１行の無名データとして扱う

### ssf キー行 ###

  1. ssf ブロックの先頭行が、 ssf キー行
  1. 末尾の連続した `空白(x20)` と `タブ(x09)` は除外
  1. 残った文字列全体が、このブロックのキーとなる

### ssf 値行 ###

  1. ssf ブロックの２行目以降が、 ssf 値行
  1. ssf 値行は、区切り記号によって、名前部分とデータ部分に分かれる
  1. 区切り記号は `; (x3b)` で、先頭に近い１つだけが有効
  1. 先頭が区切り記号のものを、無名データと呼ぶ
  1. データ部分もさらにキーごとに固有の書式を持つ

# ssf block as an independent flow #

  1. the sequence of ssf blocks is important on its meanings
  1. there're no parents or children relationships between ssf blocks. they're just serialized
  1. always exists a default target for implicit things
  1. consider there're global variables to keep defaults, maybe some block change its value, then the change impacts all blocks after that
> > A ssf block never contains other ssf blocks, nor be included in another ssf block.
> > This feature is very important to make things simple and clear.
> > Though a Worksheet object contains Range objects in the Excel Spreadsheet object model,
> > an ssf block of Cells is not included in an ssf block of Worksheet.
> > Then what tells whose property those Cells are?
> > Think ssf blocks are just flows.
> > They are independent to each other, but depend an environment, such as global variables.
> > There may be a global variable that has a "current worksheet name",
> > an ssf block of Worksheet may change that value,
> > and an ssf block of Cell may read that value to know a target sheet to act on.

## ssf ブロックは独立した流れ ##

  1. ssf ブロックは逐次処理で解釈する。
  1. 並び順は包含性ではなく、処理順を示している。
  1. 処理対象は、常にデフォルトの相手だ。
  1. デフォルトは、グローバルな環境変数などにあると考えてよく、先行する処理がそれを書き換えると、以降の処理にも影響するというルール。
> > Workbook を指定する処理があれば、それより後のブロックでの処理は、その Workbook に対して行う。
> > これは包含関係によるのではなく、そこでグローバルな環境の「対象とするWorkbook」が書き換えられたからと考える。
> > ssf は包含性を明確にする記述方法を持ってないので、グローバルな切り替えという発想で見ないと解釈にゆらぎが生じる。

# history / 改定履歴 #

  * 2011/5/2 many verbs are removed and added at svn r 735




# Declare / 宣言 #

### ssf-begin ###

```
'ssf-begin
';
```

  1. declare a start line of ssf
> > ssf の開始を宣言する。
  1. a blank ssf-line is always required for the 2nd line of this block.
> > ２行目に空の ssf 行を必ず書く。
  1. an ssf parser can use this block to determin both characters of begin and end of an ssf line.
> > ssf パーサーは、このブロックを利用して ssf 行の先頭記号と行末記号を判定することができる。
  1. ssf-begin and ssf-end are optional. an ssf parser can ignore this block to parse everything before the declare.
> > ssf-begin と ssf-end は必須ではない。 ssf パーサーは、これを無視して、宣言前のブロックをパースして使ってもよい。

### ssf-end ###

```
'ssf-end
```

  1. declare an end line of ssf
> > ssf の終了を宣言する。

# Book, Sheet / ブック、シート #

### names / 名前 ###

#### name ####

```
' name:Sheet1
' name:Sheet1/Sheet2
' name:
```

  1. switch a target object, book, sheet or module
> > ブック、シート、モジュールのような対象物を切り替える
  1. can specify the name of it
> > 対象物の名前を指定できる
  1. when the target has 2 different names, use a `/(0x2f)` to separate them. the 2nd line of the example is showing a worksheet with its name is "Sheet1" and codename is "Sheet2"
> > 対象物が２つの異なる名前を持つ時は、 `/(0x2f)` で区切る。２行目の例では、 "Sheet1" という name と "Sheet2" という codename を持つようなワークシートを示す。
  1. the 2nd name for a book is considered to be a vba project name.
> > ブックの場合、２つ目の名前は、 vba プロジェクト名として扱う。
  1. a new one is created if no names or not found,
> > 名前が指定されないときや、指定した名前がないときは新規に作る

### workbook ###

```
'workbook
'      name:Book2
```

  1. supports name
> > name が有効

### worksheet ###

```
'worksheet
'      name:Sheet1
```

  1. supports name
> > name が有効

# Cells / セル #

### names / 名前 ###

#### address ####

```
' address;R5C1:R17C2
```

  1. specify a range of cells
> > 対象セルのアドレスを指定
  1. the default is entire cells
> > 省略時は全体

#### delimiter ####

```
' delimiter;,
'          ;orange,,,apple,,
```

| orange |  |  | apple |  |  |
|:-------|:-|:-|:------|:-|:-|

  1. specify a delimiter to separate each data
> > データを分割する区切り文字を指定する

###### ====== ######

```
' delimiter;,+
'          ;orange,,,apple,,
```

| orange | apple |
|:-------|:------|

  1. continuous delimiters are considered as one and the last delimiter is ignored when a `+(x2b)` exists at the 2nd character
> > `+(x2b)` が２文字目にあれば、連続した区切りは１つとし、末尾の区切りは無視する

###### ====== ######

```
' delimiter; +
'          ;orange       apple
'          ;strawberry   lemon
```

| orange | apple |
|:-------|:------|
| strawberry | lemon |

  1. use this `space+(x202b)` to arrange vertically
> > `空白+(x202b)` で縦に揃える

###### ====== ######

```
' delimiter;★+
'          ;蜜柑★林檎★苺★檸檬
```

| 蜜柑 | 林檎 | 苺 | 檸檬 |
|:---|:---|:--|:---|

  1. far east characters are also available as a delimiter
> > 漢字なども区切り文字として使える

#### sparse ####

```
'    sparse; +
'          ;R9C2     apple or poison
```

|  | **B** |
|:-|:------|
| **9** | apple or poison |

  1. specify a single data with its address
> > 単一のデータをアドレスとともに指定する
  1. this example shows continuous spaces are used as a delimiter
> > 連続する空白を区切り文字として指定

###### ====== ######

```
'    sparse;#+
'          ;R9C2####apple or #poison#
```

|  | **B** |
|:-|:------|
| **9** | apple or #poison# |

  1. only the 1st one of delimiters effects, separates the former address and the latter data
> > 最初の区切りだけが有効で、１つ目がアドレス、２つ目がデータ

###### ====== ######

```
'    sparse;#
'          ;R9C2####apple or #poison#
```

|  | **B** |
|:-|:------|
| **9** | ###apple or #poison# |

  1. this example shows a single # is used as a delimiter
> > １つの # だけを区切りに指定した例

#### skip ####

```
'   skip;3
```

  1. means continuous 3 empty data
> > 空データが３つ続くという意味
  1. reader should not touch the cell but just move the cursor for the specified amount
> > 読み取りの場合、セルには何もせず、カーソルを指示された数だけ進めるべき

#### repeat ####

```
'  repeat;5
'           ;=R[-1]C
```

  1. means continuous 5 same data `=R[-1]C`
> > `=R[-1]C` という同一データが５つ続くという意味

#### fill ####

```
'    fill;
'   value;yyyy/m/d
'        ;B3:B5
'        ;D3:D5
'        ;E8:F9
```

  1. declare the data first, then put to the address list specified in an anonymous line
> > 先にデータを与え、無名データで渡されたアドレス一覧に埋め込む。


#### value ####

  1. declare a fixed value to fill.
> > fill で使う固定値を宣言する。

#### unit ####

```
'  unit;pt
```

  1. declare a scale unit
> > サイズの単位を宣言する。



### cells-formula ###

  1. supports address, delimiter and sparse
> > address, delimiter, sparse が有効

###### ====== ######

```
'cells-formula
' address;R1C1:R3C2
'        ;Oranges
'        ;1600
'        ;Apples
'        ;1400
'        ;Total
'        ;=SUM(R[-2]C:R[-1]C)
```

|  | **A** | **B** |
|:-|:------|:------|
| **1** | Oranges | 1,600 |
| **2** | Apples | 1,400 |
| **3** | Total | 3,000 |

  1. specify formulas of cell by anonymous data
> > 無名データでセル数式を書く
  1. a line goes into a single cell
> > １行が１セルに対応
  1. horizontal scan from left to right, then go to next line
> > 行ごとに順に処理し、横方向に左から右へセルを進む

###### ====== ######

```
'cells-formula
' address;R2C1:R3C2
'        ;license
'{{{
Fortitudinous, Free, Fair
自律,自由,公正
http://cowares.nobody.jp
'}}}
'        ;this project belongs to the public domain.
```

|  | **A** | **B** |
|:-|:------|:------|
| **2** | license | Fortitudinous, Free, Fair<br>自律,自由,公正<br><a href='http://cowares.nobody.jp'>http://cowares.nobody.jp</a> <br>
<tr><td> <b>3</b> </td><td> this project belongs to the public domain. </td><td>       </td></tr></tbody></table>

  1. escaped lines are folded as below, before parsing
> > 行エスケープは解釈の前に次のように折りたたまれる
```
'        ;license
'        ;Foritudinous,F...(CrLf)...(CrLf)http...nobody.jp
'        ;this project belongs to the public domain.
```

### cells-text ###

  1. supports address, delimiter and sparse
> > address, delimiter, sparse が有効

###### ====== ######

```
'cells-text
' address;R1C1:R3C2
'        ;Oranges
'        ;1,600
'        ;Apples
'        ;1,400
'        ;Total
'        ;3,000
```

|  | **A** | **B** |
|:-|:------|:------|
| **1** | Oranges | 1,600 |
| **2** | Apples | 1,400 |
| **3** | Total | 3,000 |

  1. similar as the cell-formula except this shows appearances as a text
> > cell-formula とほぼ同じだが、こちらは外観をテキストで示す
  1. at the B3 cell, while we have a formula SUM, the shown is a calculated and formatted value
> > B3セルには SUM 数式があるが、計算結果を書式つきで出している

### cells-numberformat ###

  1. supports address, delimiter and sparse
> > address, delimiter, sparse が有効

###### ====== ######

```
'cells-numberformat
'  sparse; +
'        ;R9C2      m/d/yyyy h:mm
'        ;R9C3      General
```

  1. shows the format strings of cells
> > セルの書式設定を示す
  1. shows different strings in localized versions
> > 各国版では書式文字が違って見える場合もある

### cells-numberformat-local ###

  1. supports address, delimiter and sparse
> > address, delimiter, sparse が有効

###### ====== ######

```
'cells-numberformat-local
'  sparse; +
'        ;R9C2      yyyy/m/d h:mm
'        ;R9C3      G/標準
```

  1. shows the localized format strings of cells
> > セルの各国対応書式設定を示す

### cells-name ###

  1. supports sparse
> > sparse が有効

###### ====== ######

```
'cells-name
'    sparse; +
'          ;=Sheet1!R6C1:R6C2      Author
'          ;=Sheet1!R3C2           Sheet1!Comment
```

  1. shows the name of cells
> > セルの名前を示す
  1. mark the Names object belongs to a Excel Workbook object, not a Worksheet
> > エクセルではワークシートでなくブックが、名前オブジェクトを所有していることに注意

### cells-color ###

```
'cells-color
'  address;C27:H27
'   repeat;2
'         ;#FFCC99
'   repeat;4
'         ;#CCFFCC
```

### cells-background-color ###

```
'cells-background-color
'  address;A24:M24
'   repeat;13
'         ;#FF6600
```

### cells-font-name ###

```
'cells-font-name
'  address;B24
'         ;Arial
```

### cells-font-size ###

```
'cells-font-size
'  address;B24
'         ;12
```

### cells-font-bold ###

```
'cells-font-bold
'  address;B24
'         ;yes
```

### cells-font-italic ###

```
'cells-font-italic
'  address;B24
'         ;yes
```

### cells-height ###

```
'cells-height
'   unit;pt
'  address;A24
'         ;14.25
```

### cells-width ###

```
'cells-width
'   unit;pt
'  address;B1
'         ;96.75
```

### cells-v-align ###

```
'cells-v-align
'  address;C35:F35
'   repeat;4
'         ;center
```

### cells-h-align ###

```
'cells-h-align
'  address;C35:F35
'   repeat;4
'         ;right
```

### cells-shrink ###

```
'cells-shrink
'  address;B24
'         ;yes
```

### cells-wrap ###

```
'cells-wrap
'  address;B24
'         ;yes
```

### cells-border ###



# VBA Code / VBA コード #

### names / 名前 ###

#### name ####

```
'  name;HelloWorld
```

  1. give the name of the code, which is displayed in a project pane on the Visual Basic Editor
> > Visual Basic ウィンドウのプロジェクトペインに表示されるコード名を示す

### code ###

  1. supports name
> > name が有効
  1. belongs to one of Microsoft Excel Objects, such a worksheet or the ThisWorkbook
> > Microsoft Excel Objects 内にあるワークシートや ThisWorkbook に属する

```
'code
'    name;ThisWorkbook
'{{{
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = IIf(MsgBox("Don't leave me!", vbExclamation Or vbYesNo, "ooooops!") = vbNo, True, False)
End Sub

'}}}
```

### module ###

  1. supports name
> > name が有効
  1. one of modules
> > 標準モジュール

```
'module
'  name;HelloWorld
'{{{
Public Sub HelloWorld()
    Selection.Value = "Hello World"
End Sub
'}}}
```

### class ###

  1. supports name
> > name が有効
  1. one of classes
> > クラスモジュール

```
'class
'    name;Class1
'{{{
Private Sub Class_Initialize()
    Debug.Print "Class1 initialized"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Class1 terminated"
End Sub
'}}}
```

### require ###

```
'require
'  ;{420B2830-E718-11CF-893D-00A0C9054228} 1 0 Microsoft Scripting Runtime
```

  1. references required to add
> > 追加すべき参照設定
  1. guid on the registry, major version, minor version and description
> > レジストリ上のguid, 主バージョン, 副バージョン, 説明文
  1. delimiters are 'space(x020)`
> > '空白(x020)` で区切る

# Book Properties / ブックに属するもの #

### book-identity ###

```
'book-identity
'  title;セル計算値確定
'  description;セル数式の計算値を確定する
```

  1. informations found in a buitlin book properties.
> > 標準のブック情報に入る項目
  1. title and description are listed when the book is used as an addin.
> > title と description は、アドインブックの一覧で利用される。