# Introduction #

  * how to change the name "ThisWorkbook" in vba

## 概要 ##
  * VBAで "ThisWorkbook" という名前を変える

# Details #

  * small tips arround the ThisWorkbook.
  * anyway, Microsoft clearly says that changing the name is NOT recommended even by hands on VBIDE.

## 説明 ##
  * ThisWorkbook 周辺のいろいろ
  * いずれにせよ、マイクロソフトはこれを ThisWorkbook 以外に変更するのはお奨めしないと明言している。たとえ開発ウィンドウで手で変更するとしてもだ。

# Cases #

### Case 1 ###

  * change "ThisWorkbook" to "bok"
> > ThisWorkbook を bok に変える

#### NG ####

```
ThisWorkbook.VBProject.VBComponents("ThisWorkbook").Name = "bok"
```

  * above codes destroy the book.
> > 上記の方法は、ブックを破壊する。
  * after you save it, you can never open it.
> > 保存したら最後、二度と開けない。
  * the above changes only the project pane, no changes at the properties pane.
> > この方法はプロジェクトペインの表示だけ変更して、プロパティまでは変更していない。

#### OK ####

```
ThisWorkbook.VBProject.VBComponents("ThisWorkbook").Properties("_CodeName") = "bok"
```

  * the above is working, and the book is not broken yet.
> > 上記の方法は機能するし、今のところブックも壊れていない
  * anyway i don't want to do it.
> > でも自分で使おうとは思わない。

#### why? ####

  1. there're 2 properties for the CodeName, "CodeName" and "`_`CodeName".
> > CodeName 関連では "CodeName" と "`_`CodeName" と、２つのプロパティがある。
  1. the former is safe, readonly.
> > 前者は読み出し専用で安全なもの。
  1. the latter seems to be defined to do something not recommended, so I did.
> > 後者は推奨できない理由で定義されてるように思えて、試してみた。
  1. at least, the "`_`CodeName" changes both the project and properties pane.
> > とりあえずこれなら、プロジェクトとプロパティの両方を変更してくれる。

### Case 2 ###

  * modify VBA source codes in the ThisWorkbook by VBA
> > ThisWorkbook の VBA コードを、 VBA で書き換える
  * sometime it success, many time it fail, I don't know why yet.
> > 成功するときもあるが、失敗することが多い。理由は不明
  * this causes a no return fall of Excel itself.
> > 失敗するとエクセルごと落ちる
  * it may be,,, i think,
    1. source code change of the ThisWorkbook let Excel to recompile the project.
    1. after that, any changes of source raise an unrecoverable Excel error.
    1. thus changing the code at the last of the VBA task may bring a lucky result.
  * たぶん、こういうことか、、、
    1. ThisWorkbook のコードを書き換えると、エクセルが再コンパイルを開始する。
    1. この後に何か他のソースを変更すると、エクセルが回復不能なエラーを起こす。
    1. よって、こいつを書き換えるのを一連の作業の最後にすると、いい結果を生むのかも。

