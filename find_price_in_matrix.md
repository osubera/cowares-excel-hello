

# Introduction #

  * find a unit price from a flexible price matrix

## 概要 ##
  * エクセルで修正可能なマトリックス単価表から単価を探す

# Requirements #

  * select a price using two axis.
  * the flag (the 1st axis) is a value to select a group of prices.
  * each flag (each group) has its own rules (the 2nd axis) to select a final price.
  * because the rules are not finalized, we need to make an exception. in that case, the price overwritten should be explicitly distiguished by a sign.
  * because the rules are variable, we need a flexible way to apply a new one.

## 要望 ##
  * ２つの軸から単価を決める。
  * フラグ（第１軸）は、おおまかなグループ分けをする。
  * フラグごとに異なるルール（第２軸）があり、最終的に１つの単価が決まる。
  * 未知のルールによる例外措置が必要なので、これにより上書きした場合、目印をつけて分かるようにする。
  * ルールは流動的なので、変更しやすい表現が欲しい。

# Details #

  * both the 1st and the 2nd axis are placed vertically in columns. we select a final unit price from a price table of vertical matrix.
  * to overwrite a price calculated, we use 2 cells for a price, to calculate and to be overwritten.
  * the 1st axis is a value of the flag.
  * there is a rule area in a worksheet. the rules are defined there by putting some values in it. each value specify that the column is required to meet the rule. selects the price when all of them are satisfied.
  * an upper rule has a priority than the lower. maybe, the higher price is to be upper.
  * calculated price is 0 when no rules are met.
  * we have an image of a rule area below.

## 説明 ##
  * 第１軸と第２軸を共にカラムとして縦方向に持つ、縦型のマトリックス単価表を作り、それを参照して最終単価を決める。
  * 上書き可能なように、物件のデータ行には計算による単価欄と上書き入力可能な単価欄を並べて２つ作り、上書き入力をした場合の区別をつけやすくする。
  * フラグの数値をそのまま第１軸の数値として使う。
  * フラグごとに単価を決めるルールエリアをワークシートの下方に設ける。 そこに印を入れる方法でルールを表現する。 計算対象行の対応する欄すべてに記入があれば、その単価を使う。
  * ルールは上優先で解釈するので、単価の高い（優先度の高い）順に並べておく。
  * フラグとルールの合致するものが無ければ、単価は 0 円とする。
  * ルールエリアのイメージ
> > ![http://2.bp.blogspot.com/_EUW0nrj9XlM/TUeYhzpng0I/AAAAAAAAACU/eCSGSM4YS3A/s1600/shot1.png](http://2.bp.blogspot.com/_EUW0nrj9XlM/TUeYhzpng0I/AAAAAAAAACU/eCSGSM4YS3A/s1600/shot1.png)


# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=find_price_in_matrix)

# How to use #

  * needs to enable the macro, because a worksheet function is described by vba macro.
  * `FindPrice` Worksheet Function is available to calculate a unit price.

## 使い方 ##
  * ワークシート関数をマクロで定義しているので、マクロを有効にしないと動かない。
  * `FindPrice` というワークシート関数が追加されているので、それで単価を計算する。

### ワークシート関数 FindPrice  / Worksheet Function ###
  * =FindPrice(ルールエリア, フラグ, 行データ) のように書く。
    * このブックでは B列に次のような式が入れてある。 =findprice(ルールエリア,F3,E3:P3)
    * このブックでは `ルールエリア` の範囲名が定義されている。 =PriceMatrix!A1:N30
  * 第１引数：ルールエリア
    * 第２軸のルールと単価を定義した表の範囲を参照する。
  * 第２引数：フラグ
    * 第１軸の値を指定する。
  * 第３引数：行データ
    * 第２軸の値の入った範囲を指定する。
    * ルールエリアの項目と同じものが同じ順で並んでいなければならない。
  * 戻り値：単価
    * ルールに合致する単価を返す。合致するルールが無い場合、 0 を返す。

### ルールエリア / RuleArea ###
  * `PriceMatrix` ワークシートに、範囲名 `ルールエリア` で定義されている。
  * シート名は他のものに変えてもよい。
  * 範囲名はワークシート関数で参照するだけなので、両方を同一のものに変更するなら、別の名称でも構わない。範囲名を使わない、生のアドレスでもよい。
  * ルールエリアの範囲ごと、切り取りと貼り付けをすれば、どのシートにでも移せる。シートの左上に置く必要もなく、好きな位置に置いてよい。計算データのシートと１枚にまとめてもよい。
  * ルールの行やカラムを増減した時は、範囲名もそれにあわせて修正する必要がある。（上手に増減すれば、範囲名は追随する）

### ルールエリアの行 / Rows of RuleArea ###
  * 先頭４行は、読むための見出しで、実際のルールは５行目から始まる。
  * ルールは上から順に評価する。
  * 各行のルールを評価する際、フラグが一致し、指定項目がすべて埋まっていたら、そのルールを即座に採用する。
  * したがって、優先度の高い順（条件が厳しい順）にルールを書く必要がある。

### ルールエリアのカラム / Columns of RuleArea ###
  * １桁目が、フラグ。
  * ２桁目が、単価。
  * ３桁目以降に、２軸のルールが並ぶ。
  * ３桁目以降は、 `FindPrice` 関数の第３引数と同じ項目なので、物件データシートの項目をそのまま順に記載したものとなる。
  * 上の理由により、フラグなどの不要な項目も入っている。
  * ルール指定方法は、該当項目のセルに `1` を入れる。 このとき、この項目が必須になり、この欄が埋まっているデータのみに、この単価が使われる。実際には `1` 以外でも、何かが入っていれば指定したことになる。

### 単価の上書き / overwriting the price ###
  * `Sheet1` ワークシートの `A:B列` には、 `FindPrice` 関数による計算と、それを上書きする見本が載っている。
  * `B列` は、すべて、 `FindPrice` による計算値になっていて、 `B4` などが、該当ルール無しで 0 になっている。
  * `A列` は、基本的には `=B4` のような式で、 `B列` と同一の数値になっているが、ここを上書きして `4` などと入れれば、セルが赤色になる。
  * これは条件書式を使っているので、セルと一緒にコピーされるし、エクセルメニューから変更もできる。

### 他のブックで使う / work with another Workbook ###
  * コピーして、このブックを増やして使う方法と、このブックを外部参照して使う方法がある。

  1. コピーする方法
    * ブックをコピーして、 `Sheet1` を書き換えればよい。
    * 計算ルールが毎月違う、という場合、この方法で、新しい計算に対応したブックを作っていくことができる。
  1. 外部参照する方法
    * =find\_price\_in\_matrix.xls!findprice(find\_price\_in\_matrix.xls!ルールエリア,C4,B4:M4)
    * このように、エクセルの通常の外部参照と同じ方法で、別のブックから、ワークシート関数とルールエリアを呼び出して使うことができる。
    * １つのルールで複数のブックを計算したい場合や、計算した数値を値コピーして確定してしまう場合には、この方法で手軽に計算だけ利用できる。

## 2011/4/27 の修正 ##

  * 操作ボタンを出すようにし、ボタン操作で集計シートを作る。
  * 操作のイメージ
    1. 今の関数、 `=FindPrice(ルールエリア,F3,E3:P3)` の３つの引数に相当する範囲を、それぞれに応じたボタン操作で設定する。
    1. `集計` ボタンを押すと、データ明細のあるブックにシートを追加し、そこで集計する。
  * 集計表には次の項目が並ぶ。
    1. 件数、単価、２軸ルールの項目名
    1. 他に集計作業に必要な列も追加される。

  * Excel 2000 Tool Bar
> > ![http://3.bp.blogspot.com/-mSHoWf6CtQE/TbdmaDQWuZI/AAAAAAAAAFA/W37nyqxpGsY/s1600/shot2.png](http://3.bp.blogspot.com/-mSHoWf6CtQE/TbdmaDQWuZI/AAAAAAAAAFA/W37nyqxpGsY/s1600/shot2.png)
    1. 計算対象のデータ範囲を指定する。 `FindPrice` 関数の第３引数を、全データ行に拡張した範囲。 データの左上セルにカーソルを置いてからボタンを押すと、範囲を自動選択する。 下のサンプルでは、 E3 セルだけを選んでからボタンを押すと、自動選択となる。 このボタンを押すと、残りの引数、フラグ列とルールエリアを自動設定する。
    1. フラグがどの列にあるかを指定する。 `FindPrice` 関数の第２引数。 見出しが `フラグ` になっている項目を自動選択するので、それと異なる列を選びたいときに、このボタンを使う。
    1. ルールエリアを指定する。 `FindPrice` 関数の第１引数。 デフォルトの範囲名 `ルールエリア` と違う場所を使うときや、他のブックにルールがある場合に、このボタンを使う。
    1. 集計を実行する。 データ範囲として指定したブックに新規ワークシートを追加して、集計表を作成する。
  * Ready to Calculate
> > ![http://3.bp.blogspot.com/-hwfbuS30OUY/TbdmadE_FWI/AAAAAAAAAFE/mQILInO5DnE/s1600/shot3.png](http://3.bp.blogspot.com/-hwfbuS30OUY/TbdmadE_FWI/AAAAAAAAAFE/mQILInO5DnE/s1600/shot3.png)
    1. ３つのボタンすべてが on になったら、集計ボタンを押す。
  * Data Area selected
> > ![http://4.bp.blogspot.com/-ejOO79Bt1Sg/TbdmbAQbVkI/AAAAAAAAAFI/nt57ewWsdYg/s1600/shot4.png](http://4.bp.blogspot.com/-ejOO79Bt1Sg/TbdmbAQbVkI/AAAAAAAAAFI/nt57ewWsdYg/s1600/shot4.png)
    1. 見本のシートでは、 E3 セルを選んで、データ範囲ボタンを押すと、自動選択でこのように拡張される。
  * Summary Report created
> > ![http://2.bp.blogspot.com/-rKzLkmubDqA/Tbdmbgyk8PI/AAAAAAAAAFM/5yZ4nywa7d8/s1600/shot5.png](http://2.bp.blogspot.com/-rKzLkmubDqA/Tbdmbgyk8PI/AAAAAAAAAFM/5yZ4nywa7d8/s1600/shot5.png)
    1. 集計で作成されるレポートの見本。
    1. エクセルの集計表機能で、レベル２に絞り込んだ表を作成する。
    1. 掘り下げて、明細レベルを確認することができる。
    1. データ数が多い場合、作成に１分程度かかる場合もある。
  * Drill down to check details
> > ![http://4.bp.blogspot.com/-VMsIlu2b9oA/Tbdmb4HG54I/AAAAAAAAAFQ/4LMCuTfVsaI/s1600/shot6.png](http://4.bp.blogspot.com/-VMsIlu2b9oA/Tbdmb4HG54I/AAAAAAAAAFQ/4LMCuTfVsaI/s1600/shot6.png)
    1. 左端(A列)は、ルールエリアの各ルールを、行番号として示している。これで集計している。
    1. ルールに該当しないものは、ルール番号 0 として、先頭で集計される。
    1. 明細では、件数欄(B列)に、対象データの行番号が入っていて、これでデータを特定できる。
    1. 明細の右側には、対象データがコピーされていて、そちらで確認することもできる。

  * 処理内容についてのメモ
    1. ワークシート関数 `FindPrice` と同等の処理をして、ルールデータを返す `FindPriceRow` 関数を使用して集計している。
      * ２つの関数は内部が共通なので、完全に同等の処理をする。
    1. データを行単位で、関数による計算結果を反映させる処理をし、ソート、集計する。
      * データ行でソートしなおせば、ワークシート関数による計算と本質的に同じものが出る。
      * 行単位処理なので、行数に比例して時間がかかる。行単位までの最適化しか行っていない。

# Code #

```
'ssf-begin
';

'workbook
'   name;find_price_in_matrix.xls/F3FindPriceInMatrix

'book-identity
'  title;find_price_in_matrix
'  description;修正可能なマトリックス単価表から単価を探す

'require

'cells-name
'       ;=find_price_in_matrix!R15C1
'       ;find_price_in_matrix!_ButtonCaption
'       ;=find_price_in_matrix!R3C2
'       ;find_price_in_matrix!_Comment
'       ;=find_price_in_matrix!R6C2
'       ;find_price_in_matrix!_Contributor
'       ;=find_price_in_matrix!R4C2
'       ;find_price_in_matrix!_Copyright
'       ;=find_price_in_matrix!R5C2
'       ;find_price_in_matrix!_License
'       ;=find_price_in_matrix!R2C2
'       ;find_price_in_matrix!_LocalComment
'       ;=find_price_in_matrix!R1C2
'       ;find_price_in_matrix!_PublicName
'       ;=find_price_in_matrix!R7C2
'       ;find_price_in_matrix!_Since
'       ;=find_price_in_matrix!R10C2
'       ;find_price_in_matrix!_Tag
'       ;=find_price_in_matrix!R9C2
'       ;find_price_in_matrix!_Timestamp
'       ;=find_price_in_matrix!R8C2
'       ;find_price_in_matrix!_Url
'       ;=PriceMatrix!R1C1:R30C14
'       ;ルールエリア

'worksheet
'   name;Sheet1

'cells-formula
'  address;A1:V18
'     skip;5
'         ;21
'         ;22
'         ;23
'         ;24
'         ;25
'         ;26
'         ;27
'         ;28
'         ;29
'         ;30
'         ;31
'         ;32
'         ;33
'         ;34
'         ;35
'         ;36
'         ;37
'         ;確定単価
'         ;単価計算
'     skip;2
'         ;指定サイトURL
'         ;フラグ
'         ;電話
'         ;住所
'         ;fax
'         ;確認
'         ;記号
'         ;ポイント
'         ;1回目確認URL
'         ;備考
'         ;名前
'         ;事務所
'     skip;3
'         ;2回目確認URL
'     skip;2
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;1
'         ;http://www.yahoo.co.jp/
'         ;3
'         ;888-888-8888
'         ;文京区本郷７－３－１
'         ;777-777-7777
'     skip;3
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;4
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;2
'         ;http://www.yahoo.co.jp/
'         ;3
'     skip;1
'         ;文京区本郷７－３－１
'         ;777-777-7777
'     skip;3
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;3
'         ;http://www.yahoo.co.jp/
'         ;9
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;4
'         ;http://www.yahoo.co.jp/
'         ;4
'     skip;1
'         ;文京区本郷７－３－１
'         ;777-777-7777
'     skip;3
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;5
'         ;http://www.yahoo.co.jp/
'         ;3
'         ;888-888-8888
'     skip;5
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;6
'         ;http://www.yahoo.co.jp/
'         ;0
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;7
'         ;http://www.yahoo.co.jp/
'         ;0
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;8
'         ;http://www.yahoo.co.jp/
'         ;1
'         ;888-888-8888
'         ;文京区本郷７－３－１
'         ;777-777-7777
'     skip;3
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;9
'         ;http://www.yahoo.co.jp/
'         ;1
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;10
'         ;http://www.yahoo.co.jp/
'         ;0
'     skip;2
'         ;777-777-7777
'     skip;13
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;11
'         ;http://www.yahoo.co.jp/
'         ;2
'     skip;6
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;12
'         ;http://www.yahoo.co.jp/
'         ;9
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;13
'     skip;1
'         ;9
'     skip;16
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;14
'         ;http://www.yahoo.co.jp/
'         ;2
'     skip;2
'         ;777-777-7777
'     skip;3
'         ;http://www.yahoo.co.jp/
'     skip;9
'         ;=RC[1]
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;15
'         ;http://www.yahoo.co.jp/
'         ;0
'     skip;16
'         ;1.5
'         ;=findprice(ルールエリア,RC[4],RC[3]:RC[14])
'     skip;1
'         ;16
'         ;http://www.yahoo.co.jp/
'         ;1
'     skip;1
'         ;文京区本郷７－３－１

'cells-height
'   unit;pt
'  address;A2
'         ;108

'cells-h-align
'  address;A2:P2
'   repeat;2
'         ;center
'     skip;2
'   repeat;12
'         ;center

'cells-v-align
'  address;A2:P2
'   repeat;2
'         ;center
'     skip;2
'   repeat;12
'         ;center

'worksheet
'   name;PriceMatrix

'cells-formula
'  address;A1:N3
'         ;１軸
'         ;単価
'         ;２軸ルール
'     skip;14
'         ;1
'         ;2
'         ;3
'         ;4
'         ;5
'         ;6
'         ;7
'         ;8
'         ;9
'         ;10
'         ;11
'         ;フラグ
'     skip;1
'         ;指定サイトURL
'         ;フラグ
'         ;電話
'         ;住所
'         ;fax
'         ;確認
'         ;記号
'         ;ポイント
'         ;1回目確認URL
'         ;備考
'         ;名前
'         ;事務所
'  address;A5:K30
'     skip;1
'         ;2
'         ;1
'     skip;3
'         ;1
'     skip;5
'         ;1.5
'         ;1
'     skip;8
'         ;1
'         ;3.5
'         ;1
'     skip;1
'   repeat;3
'         ;1
'     skip;4
'         ;1
'         ;3
'         ;1
'     skip;1
'   repeat;2
'         ;1
'     skip;5
'         ;1
'         ;3
'         ;1
'     skip;1
'         ;1
'     skip;1
'         ;1
'     skip;4
'         ;1
'         ;2.5
'         ;1
'     skip;1
'         ;1
'     skip;6
'         ;2
'         ;3
'         ;1
'     skip;3
'         ;1
'     skip;3
'         ;1
'         ;2
'         ;2.5
'         ;1
'     skip;7
'         ;1
'         ;3
'         ;5
'         ;1
'     skip;1
'   repeat;3
'         ;1
'     skip;3
'         ;1
'         ;3
'         ;4.5
'         ;1
'     skip;1
'   repeat;2
'         ;1
'     skip;4
'         ;1
'         ;3
'         ;4.5
'         ;1
'     skip;1
'         ;1
'     skip;1
'         ;1
'     skip;3
'         ;1
'         ;3
'         ;4
'         ;1
'     skip;1
'         ;1
'     skip;5
'         ;1
'         ;9
'         ;2
'         ;1
'     skip;8
'         ;9
'         ;1.5
'     skip;9
'         ;4
'         ;5
'         ;1
'     skip;1
'   repeat;3
'         ;1
'     skip;3
'         ;1
'         ;4
'         ;4.5
'         ;1
'     skip;1
'   repeat;2
'         ;1
'     skip;4
'         ;1
'         ;4
'         ;4.5
'         ;1
'     skip;1
'         ;1
'     skip;1
'         ;1
'     skip;3
'         ;1
'   repeat;2
'         ;4
'         ;1
'     skip;1
'         ;1
'     skip;5
'         ;1
'         ;4
'         ;3
'         ;1
'     skip;3
'         ;1
'     skip;3
'         ;1
'         ;4
'         ;2.5
'         ;1
'     skip;7
'         ;1
'         ;4
'         ;3.5
'         ;1
'     skip;1
'   repeat;3
'         ;1
'     skip;4
'         ;4
'         ;3
'         ;1
'     skip;1
'   repeat;2
'         ;1
'     skip;5
'         ;4
'         ;3
'         ;1
'     skip;1
'         ;1
'     skip;1
'         ;1
'     skip;4
'         ;4
'         ;2.5
'         ;1
'     skip;1
'         ;1
'     skip;6
'         ;4
'         ;2
'         ;1
'     skip;8
'         ;4
'         ;1.5

'cells-background-color
'  address;A1:N1
'         ;#FFFF00
'         ;#FF99CC
'   repeat;12
'         ;#FFFF00
'  address;A3:N3
'         ;#FFFF00
'     skip;1
'   repeat;12
'         ;#FFFF00
'  address;A5:N30
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;13
'         ;#FFFF00
'         ;#FF99CC
'   repeat;12
'         ;#FFFF00

'worksheet
'   name;find_price_in_matrix/BaumMain

'cells-formula
'  address;A1:B10
'         ;名称
'         ;find_price_in_matrix
'         ;コメント
'         ;修正可能なマトリックス単価表から単価を探す
'         ;comment
'         ;find a unit price from a flexible price matrix
'         ;著作権
'         ;="Copyright (C) " &R[3]C & "-" & YEAR(R[5]C) & " " & R[2]C
'         ;ライセンス
'         ;自律, 自由, 公正, http://cowares.nobody.jp
'         ;作者
'         ;Tomizono - kobobau.com
'         ;初版
'         ;2011
'         ;配布元
'         ;http://code.google.com/p/cowares-excel-hello/wiki/find_price_in_matrix
'         ;更新
'         ;40660.3333333333
'         ;keyword
'         ;vba,excel
'  address;A13:J13
'         ;ボタンの表示
'         ;ボタンの機能
'         ;Tag
'         ;Parameter
'         ;ControlType
'         ;Style
'         ;Width
'         ;Group
'         ;Action
'         ;Initialize ..
'  address;A15:H18
'         ;データ範囲
'         ;集計の対象とするデータ範囲を指定する。セルが１つしか選ばれていなければ、そのセルを左上端とした範囲で自動設定する。フラグ列を自動設定する。
'         ;data
'         ;=_ButtonCaption
'         ;1
'         ;2
'     skip;1
'         ;1
'         ;フラグ列
'         ;自動設定されたフラグ列が違う場合に、このボタンで指定する。
'         ;flag
'         ;=RC[-3]
'         ;1
'         ;2
'     skip;2
'         ;ルールエリア
'         ;標準の範囲名「ルールエリア」以外の範囲を使う場合に、このボタンで指定する。
'         ;rulearea
'         ;=RC[-3]
'         ;1
'         ;2
'     skip;2
'         ;集計
'         ;指定された範囲から集計表を作成する。
'         ;sumup
'     skip;1
'         ;1
'         ;2
'     skip;1
'         ;1

'cells-numberformat
'  address;B9
'         ;m/d/yyyy h:mm

'cells-width
'   unit;pt
'  address;B1
'         ;96.75

'module
'   name;FindPriceInMatrix
'{{{
Option Explicit

Const MxRstart = 5
Const MxRname = 3
Const MxCstart = 3
Const MxCprice = 2
Const MxCflag = 1

Public Function FindPrice(RuleArea As Range, ByVal Flag As Long, Rules As Range) As Variant
    Dim out As Variant
    Dim Mx As Variant
    Dim MxRow As Variant
    Dim D As Variant
    
    out = 0
    
    SmartArray RuleArea, Rules, Mx, D
    
    For Each MxRow In Mx
        If IsMet1st(MxRow, Flag) Then
            If IsMet2nd(MxRow, D) Then
                out = SelectPrice(MxRow)
                Exit For
            End If
        End If
    Next
    
    FindPrice = out
End Function

Public Function FindPriceRow(RuleArea As Range, ByVal Flag As Long, Rules As Range) As Variant
    Dim out As Variant
    Dim Mx As Variant
    Dim MxRow As Variant
    Dim D As Variant
    Dim RowNumber As Long
    
    out = Array(0, Empty)
    RowNumber = RuleArea.Cells(1).Row + MxRstart - 1
    
    SmartArray RuleArea, Rules, Mx, D
    
    For Each MxRow In Mx
        If IsMet1st(MxRow, Flag) Then
            If IsMet2nd(MxRow, D) Then
                out = Array(RowNumber, MxRow)
                Exit For
            End If
        End If
        RowNumber = RowNumber + 1
    Next
    
    FindPriceRow = out
End Function

Public Function GetRuleNames(RuleArea As Range) As Variant
    Dim itemA() As Variant
    Dim j As Long
    
    ReDim itemA(0 To RuleArea.Columns.Count - MxCstart)
    
    For j = 0 To UBound(itemA)
        itemA(j) = RuleArea(MxRname, j + MxCstart)
    Next
    
    GetRuleNames = itemA
End Function

Private Function IsMet1st(MxRow As Variant, Flag As Long) As Boolean
    IsMet1st = (Flag = MxRow(1))
End Function

Private Function IsMet2nd(MxRow As Variant, Data As Variant) As Boolean
    Dim i As Long
    
    IsMet2nd = False
    
    For i = 0 To UBound(MxRow(2))
        If Not IsEmpty(MxRow(2)(i)) Then
            If IsEmpty(Data(i)) Then Exit Function
        End If
    Next
    
    IsMet2nd = True
End Function

Private Function SelectPrice(MxRow As Variant) As Variant
    SelectPrice = CDec(MxRow(0))
End Function

Private Function SmartArray(RuleArea As Range, Data As Range, _
        ByRef VarRuleArea As Variant, ByRef VarData As Variant) As Long
    Dim outA() As Variant
    Dim outB() As Variant
    Dim itemA() As Variant
    Dim i As Long
    Dim j As Long
    Dim UpData As Long
    Dim out As Long
    
    ReDim outA(0 To RuleArea.Rows.Count - MxRstart)
    ReDim itemA(0 To RuleArea.Columns.Count - MxCstart)
    ReDim outB(0 To UBound(itemA))
    out = UBound(itemA) + 1
    
    For i = 0 To UBound(outA)
        For j = 0 To UBound(itemA)
            itemA(j) = RuleArea(i + MxRstart, j + MxCstart)
        Next
        outA(i) = Array(RuleArea(i + MxRstart, MxCprice).Value, _
                        RuleArea(i + MxRstart, MxCflag).Value, _
                        itemA)
    Next
    
    UpData = Data.Cells.Count - 1
    For i = 0 To UBound(outB)
        If i > UpData Then Exit For
        outB(i) = Data.Cells(i + 1)
    Next
    
    VarRuleArea = outA
    VarData = outB
    SmartArray = out
End Function


'=== unit test begin ===
#If False Then

Sub test_IsMet1st()
    Debug.Assert IsMet1st(Array(100, 0), 0)
    Debug.Assert Not IsMet1st(Array(100, 0), 1)
    Debug.Assert IsMet1st(Array(100, 1), 1)
    Debug.Assert Not IsMet1st(Array(100, 1), 2)
End Sub

Sub test_IsMet2nd()
    Debug.Assert IsMet2nd(Array(0, 0, Array(1, 1)), Array("a", "a"))
    Debug.Assert IsMet2nd(Array(0, 0, Array(1, Empty, 1)), Array("a", Empty, "a"))
    Debug.Assert IsMet2nd(Array(0, 0, Array(1, Empty, 1)), Array("a", "a", "a"))
    Debug.Assert Not IsMet2nd(Array(0, 0, Array(1, Empty, 1)), Array("a", "a", Empty))
End Sub

Sub test_SelectPrice()
    Debug.Assert SelectPrice(Array(12.3)) = 12.3
End Sub

Sub test_SmartArray()
    Dim x As Range
    Dim y As Range
    Dim a As Variant
    Dim b As Variant
    Dim C As Long
    
    Set x = PriceMatrix.Range("A1:L12")
    Set y = Sheet1.Range("D3:M3")
    
    C = SmartArray(x, y, a, b)
    
    Stop
End Sub

Sub test_FindPrice()
    Dim x As Range
    Dim y As Range
    
    Set x = PriceMatrix.Range("A1:L12")
    Set y = Sheet1.Range("D3:M3")
    
    Debug.Print FindPrice(x, 0, y)
    Debug.Print FindPrice(x, 1, y)
    Debug.Print FindPrice(x, 9, y)
    Debug.Print TypeName(FindPrice(x, 0, y))
    
End Sub

Sub test_FindPriceRow()
    Dim x As Range
    Dim y As Range
    
    Set x = PriceMatrix.Range("A1:L12")
    Set y = Sheet1.Range("D3:M3")
    
    Debug.Print FindPriceRow(x, 0, y)(0)
    Debug.Print FindPriceRow(x, 1, y)(0)
    Debug.Print FindPriceRow(x, 9, y)(0)
    Debug.Print TypeName(FindPriceRow(x, 0, y))
    
End Sub

Sub test_GetRuleNames()
    Dim x As Range
    Dim y As Variant
    
    Set x = PriceMatrix.Range("A1:L12")
    y = GetRuleNames(x)
    
    Stop
End Sub

#End If
'=== unit test end ===

'}}}

'code
'   name;BaumMain
'{{{
Option Explicit

' BaumMain addin for ToolBarV2

' using a excel worksheet as a property holder

' we do not support popup on excel sheet at this moment
' no ideas how to describe it wisely on 2 dimensional sheet

Private Helper As ToolBarV2
'Public Helper As ToolBarV2
' we cannot have a public variable in Excel Worksheet

Friend Function VBProjectName() As String
    ' VBProject.Name can't be accessed for the default settings.
    VBProjectName = "F3FindPriceInMatrix"
End Function

Friend Function GetHelper() As ToolBarV2
    Set GetHelper = Helper
End Function

'=== default main procedures begin ===


' this will called by pressing a button
Public Sub BarMain(Optional oWho As Object = Nothing)
    If Helper Is Nothing Then
        BarInitialize
        MsgBox "ツールバーを修復しました。もう一度操作してください。", vbExclamation, BaumMain.Name
    Else
        Helper.BarMain Me
    End If
End Sub

Public Sub OnButtonToggle()
    Helper.OnButtonToggle
End Sub

' followings need to be public, because they are called from outside by the Helper
' we also can copy the Helper.BarMain code here, and let the followings be private.

Public Sub Menu_data(oAC As Object)
    FindPriceInMatrixSum.Set_data Selection, Helper, oAC, Helper.GetButton("flag"), Helper.GetButton("rulearea")
End Sub

Public Sub Menu_flag(oAC As Object)
    FindPriceInMatrixSum.Set_flag Selection, Helper, oAC
End Sub

Public Sub Menu_rulearea(oAC As Object)
    FindPriceInMatrixSum.Set_rulearea Selection, Helper, oAC
End Sub

Public Sub Menu_sumup(oAC As Object)
    FindPriceInMatrixSum.Exec_sumup
End Sub


'=== default main procedures end ===
'=== button data begin ===

Public Property Get ButtonData() As Variant
    ButtonData = ConvertRangeToArray(Application.Intersect(GetButtonRow, GetButtonCol))
End Property

Public Property Get ButtonParent() As Variant
    ButtonParent = Array(VBProjectName & "." & Me.CodeName)
End Property

' above simple property codes are supported by the following range helpers

Private Function GetButtonRow(Optional Address As String = "_ButtonCaption") As Range
    Dim out As Range
    Dim StartAt As Range
    
    Set StartAt = Me.Range(Address)
    If IsEmpty(StartAt.Offset(1, 0).Value) Then
        Set out = StartAt
    Else
        Set out = Me.Range(StartAt, StartAt.End(xlDown))
    End If
    
    Set GetButtonRow = out.EntireRow
End Function

Private Function GetButtonCol(Optional Address As String = "_ButtonCaption") As Range
    Dim StartAt As Range
    Set StartAt = Me.Range(Address)
    Set GetButtonCol = Me.Range(StartAt, StartAt.SpecialCells(xlCellTypeLastCell)).EntireColumn
End Function

Private Function ConvertRangeToArray(Ra As Range) As Variant
    Dim out() As Variant
    Dim i As Long
    
    ReDim out(0 To Ra.Rows.Count - 1)
    For i = 0 To UBound(out)
        out(i) = Ra.Rows(i + 1).Value
    Next
    
    ConvertRangeToArray = out
End Function


'=== button data end ===
'=== constructor / destructor begin ===


Private Function BarName() As String
    BarName = Me.Name & Me.Range("_PublicName").Text & Me.Range("_Timestamp").Text
End Function

Public Sub BarInitialize()
    Dim vMe As Variant
    Set vMe = Me
    Set Helper = New ToolBarV2
    Helper.SetName BarName
    Helper.NewBar vMe
End Sub

Public Sub BarTerminate()
    On Error Resume Next
    Helper.DelBar
    Set Helper = Nothing
End Sub


'=== constructor / destructor end ===

'}}}

'class
'   name;ToolBarV2
'{{{
Option Explicit

' Generate an application toolbar

Private MyBar As Office.CommandBar
Private MyName As String
Private MyApp As Application


'=== main procedures helper begin ===


' this will called by pressing a button
Friend Sub BarMain(Optional oWho As Object = Nothing)
    Dim oAC As Object   ' this is the button itself pressed
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    ' switch to a main menu procedure
    Main oAC, SomebodyOrMe(oWho)
    Set oAC = Nothing
End Sub

' main menu procedure. if you delete this, a public Main in Standard Module will be called, maybe.
Private Sub Main(oAC As Object, Optional oWho As Object = Nothing)
    ' use a button tag to switch a procedure to be called as "Menu_xx"
    CallByName SomebodyOrMe(oWho), "Menu_" & oAC.Tag, VbMethod, oAC
End Sub

Public Sub Menu_about(oAC As Object)
    MsgBox TypeName(Me), vbOKOnly, "Sample of procedure called by the Main"
End Sub

Friend Sub OnButtonToggle()
    Dim oAC As Object   ' toggle this button
    Set oAC = Application.CommandBars.ActionControl
    If oAC Is Nothing Then Exit Sub
    
    ButtonSwitchToggle oAC
    Set oAC = Nothing
End Sub

Private Function SomebodyOrMe(oWho As Object) As Object
    If oWho Is Nothing Then
        Set SomebodyOrMe = Me
    Else
        Set SomebodyOrMe = oWho
    End If
End Function


'=== main procedures helper end ===
'=== event procedures begin ===


Private Sub Class_Initialize()
    Set MyApp = Application
    MyName = CStr(Timer)    ' random name, maybe uniq
End Sub

Private Sub Class_Terminate()
    Set MyApp = Nothing
End Sub


'=== event procedures end ===
'=== construction and destruction begin ===


Public Sub NewBar(ParamArray Addins() As Variant)
    DelBar
    Set MyBar = CreateBar(MyApp, MyName)
    AddAddins MyBar, CVar(Addins)
    ShowBar MyBar
End Sub

Public Sub DelBar()
    DeleteBar MyBar
    Set MyBar = Nothing
End Sub

Public Sub SetApplication(oApp As Application)
    Set MyApp = oApp
End Sub

Public Sub SetName(NewName As String)
    MyName = NewName
End Sub

Public Property Get Bar() As Office.CommandBar
    Set Bar = MyBar
End Property


'=== construction and destruction end ===
'=== bar generator begin ===


Public Function CreateBar(oApp As Application, BarName As String) As Office.CommandBar
    RemoveExistingBar oApp, BarName
    Set CreateBar = oApp.CommandBars.Add(Name:=BarName, Temporary:=True)
End Function

Public Sub RemoveExistingBar(oApp As Application, BarName As String)
    On Error Resume Next
    oApp.CommandBars(BarName).Delete
End Sub

Public Sub DeleteBar(Bar As Object)
    On Error Resume Next
    Bar.Delete
End Sub

Public Sub ShowBar(Bar As Object, Optional Position As Long = msoBarTop, Optional Height As Long = 0)
    Bar.Visible = True
    Bar.Position = Position
    If Height > 0 Then Bar.Height = Bar.Height * Height
End Sub


'=== bar generator end ===
'=== handle addins begin ===


Public Function WithAddins(ParamArray Addins() As Variant) As Long
    WithAddins = AddAddins(MyBar, CVar(Addins))
End Function

Public Function AddAddins(Bar As Object, Addins As Variant) As Long
    Dim Addin As Variant
    Dim LastButtonIndex As Long
    
    For Each Addin In Addins
        LastButtonIndex = AddButtons(Bar, Addin.ButtonData, Addin.ButtonParent)
    Next
    
    AddAddins = LastButtonIndex
End Function


'=== handle addins end ===
'=== button generator begin ===


Public Function AddButtons(Bar As Object, Data As Variant, Parent As Variant) As Long
    Dim LastButtonIndex As Long
    Dim SingleData As Variant
    
    For Each SingleData In Data
        LastButtonIndex = Add(Bar, MakeAButtonData(SingleData, Parent))
    Next
    
    AddButtons = LastButtonIndex
End Function

Public Function Add(Bar As Object, Data As Variant) As Long
    Dim ButtonA As CommandBarControl
    
    Set ButtonA = Bar.Controls.Add(Type:=ButtonControlType(Data), Temporary:=True)
    With ButtonA
        Select Case ButtonControlType(Data)
        Case msoControlEdit                         '2      ' textbox
        Case msoControlDropdown, msoControlComboBox '3, 4   ' list and combo
            SetButtonItems ButtonA, Data
            SetButtonStyle ButtonA, Data
        Case msoControlPopup                        '10     ' popup
            SetButtonPopup ButtonA, Data
        Case msoControlButton                       '1      ' Button
            SetButtonStyle ButtonA, Data
            SetButtonState ButtonA, Data
        End Select
        SetButtonWidth ButtonA, Data
        SetButtonGroup ButtonA, Data
        .OnAction = ButtonAction(Data)
        .Caption = ButtonCaption(Data)
        .TooltipText = ButtonDescription(Data)
        .Tag = ButtonTag(Data)
        .Parameter = ButtonParameter(Data)
    End With
    
    Add = ButtonA.Index
    Set ButtonA = Nothing
End Function

Public Sub Remove(Bar As Object, Items As Variant)
    On Error Resume Next
    Dim Item As Variant
    
    If IsArray(Item) Then
        For Each Item In Items
            Remove Bar, Item
        Next
    Else
        Bar.Controls(Item).Delete
    End If
End Sub


'=== button generator end ===
'=== button data structure begin ===


' generator / selector

' Data(): Array of button data
' Parent(): Array of button parent information (bar and properties)
'           Parent(0) is reserved for addin key


Public Function MakeAButtonData(Data As Variant, Parent As Variant) As Variant
    MakeAButtonData = Array(NormalizeArray(Data), Parent)
End Function

Public Function DataAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    DataAButtonData = AButtonData(0)
End Function

Public Function ParentAButtonData(AButtonData As Variant) As Variant
    On Error Resume Next
    ParentAButtonData = AButtonData(1)
End Function

Public Function KeyAButtonData(AButtonData As Variant) As String
    On Error Resume Next
    KeyAButtonData = ParentAButtonData(AButtonData)(0)
End Function

Public Function ItemAButtonData(AButtonData As Variant, ByVal Item As Long, _
            Optional FallBack As Variant = Empty) As Variant
    On Error Resume Next
    Dim out As Variant
    
    out = DataAButtonData(AButtonData)(Item)
    If IsEmpty(out) Then out = FallBack
    
    ItemAButtonData = out
End Function


'=== button data structure end ===
'=== button data struncture detail begin ===


Public Function ButtonCaption(Data As Variant) As String
    ButtonCaption = ItemAButtonData(Data, 0)
End Function

Public Function ButtonDescription(Data As Variant) As String
    ButtonDescription = ItemAButtonData(Data, 1)
End Function

Public Function ButtonTag(Data As Variant) As String
    ButtonTag = ItemAButtonData(Data, 2, ButtonCaption(Data))
End Function

Public Function ButtonParameter(Data As Variant) As String
    ButtonParameter = ItemAButtonData(Data, 3)
End Function

Public Function ButtonControlType(Data As Variant) As Long
    'MsoControlType
    On Error Resume Next
    ButtonControlType = Val(ItemAButtonData(Data, 4, msoControlButton))
End Function

Public Function ButtonStyle(Data As Variant) As Long
    'MsoButtonStyle
    On Error Resume Next
    ButtonStyle = Val(ItemAButtonData(Data, 5, msoButtonCaption))
End Function

Public Function ButtonWidth(Data As Variant) As Long
    ' we use 45 units here
    On Error Resume Next
    Const UnitWidth = 45
    ButtonWidth = Val(ItemAButtonData(Data, 6)) * UnitWidth
End Function

Public Function ButtonGroup(Data As Variant) As Boolean
    ' put group line on its left
    ButtonGroup = Not IsEmpty(ItemAButtonData(Data, 7))
End Function

Public Function ButtonAction(Data As Variant) As String
    On Error Resume Next
    ' Standard Method Name to be kicked with the button
    Const BarMain = "BarMain"
    Dim FullName As String
    
    If KeyAButtonData(Data) = "" Then
        FullName = BarMain
    Else
        FullName = KeyAButtonData(Data) & "." & BarMain
    End If
    
    ButtonAction = ItemAButtonData(Data, 8, FullName)
End Function

Public Function ButtonItems(Data As Variant) As Variant
    Dim pan As Variant
    Dim i As Long
    
    On Error GoTo DONE
    pan = Empty
    i = 9
    
    Do Until IsEmpty(ItemAButtonData(Data, i))
        pan = Array(ItemAButtonData(Data, i), pan)
        i = i + 1
    Loop
    
DONE:
    ButtonItems = pan
End Function


'=== button data struncture detail end ===
'=== button tools for data begin ===


Public Sub SetButtonWidth(ButtonA As CommandBarControl, Data As Variant)
    If ButtonWidth(Data) > 0 Then ButtonA.Width = ButtonWidth(Data)
End Sub

Public Sub SetButtonStyle(ButtonA As Object, Data As Variant)
    On Error Resume Next
    ' Each Button does not accept each style, but we won't check them.
    If ButtonStyle(Data) <> 0 Then ButtonA.Style = ButtonStyle(Data)
End Sub

Public Sub SetButtonGroup(ButtonA As CommandBarControl, Data As Variant)
    If ButtonGroup(Data) Then ButtonA.BeginGroup = True
End Sub

Public Sub SetButtonItems(ButtonA As Object, Data As Variant)
    Dim pan As Variant
    Dim HasItem As Boolean
    
    pan = ButtonItems(Data)
    HasItem = False
    
    Do Until IsEmpty(pan)
        ButtonA.AddItem pan(0), 1
        pan = pan(1)
        HasItem = True
    Loop
    If HasItem Then ButtonA.ListIndex = 1
End Sub

Public Sub SetButtonPopup(ButtonA As CommandBarControl, Data As Variant)
    Dim MyChild As Variant
    
    MyChild = StackToArray(ButtonItems(Data))
    If UBound(MyChild) >= 0 Then Add ButtonA, MyChild
End Sub

Public Sub SetButtonState(ButtonA As Object, Data As Variant)
    If Not IsEmpty(ButtonItems(Data)) Then ButtonA.State = msoButtonDown
End Sub


'=== button tools for data end ===
'=== button tools for control object begin ===


Public Sub ComboAddHistory(oAC As Object, Optional AtBottom As Boolean = False)
    If oAC.ListIndex > 0 Then Exit Sub
    
    If AtBottom Then
        oAC.AddItem oAC.Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem oAC.Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Sub ListAddHistory(oAC As Object, Text As String, Optional AtBottom As Boolean = False)
    If AtBottom Then
        oAC.AddItem Text
        oAC.ListIndex = oAC.ListCount
    Else
        oAC.AddItem Text, 1
        oAC.ListIndex = 1
    End If
End Sub

Public Function ListFindIndex(oAC As Object, Text As String) As Long
    Dim i As Long
    For i = 1 To oAC.ListCount
        If oAC.List(i) = Text Then
            ListFindIndex = i
            Exit Function
        End If
    Next
    ListFindIndex = 0
End Function

Public Function ControlText(oAC As Object) As String
    ControlText = oAC.Text
End Function

Public Sub ButtonSwitchOn(oAC As Object)
    oAC.State = msoButtonDown
End Sub

Public Sub ButtonSwitchOff(oAC As Object)
    oAC.State = msoButtonUp
End Sub

Public Function ButtonSwitchToggle(oAC As Object) As Boolean
    ButtonSwitchToggle = (Not IsButtonStateOn(oAC))
    If ButtonSwitchToggle Then
        ButtonSwitchOn oAC
    Else
        ButtonSwitchOff oAC
    End If
End Function

Public Function IsButtonStateOn(oAC As Object) As Boolean
    IsButtonStateOn = (oAC.State = msoButtonDown)
End Function

Public Function ButtonFindByTag(oAC As Object, Tag As Variant) As CommandBarControl
    If oAC Is Nothing Then Exit Function
    If TypeName(oAC) = "CommandBar" Then
        Set ButtonFindByTag = oAC.FindControl(Tag:=Tag)
    Else
        Set ButtonFindByTag = oAC.Parent.FindControl(Tag:=Tag)
    End If
End Function


'=== button tools for control object end ===
'=== button tools for mybar begin ===


Public Function GetButton(TagOrIndex As Variant) As Office.CommandBarControl
    On Error Resume Next
    Select Case TypeName(TagOrIndex)
    Case "Long", "Integer", "Byte", "Double", "Single"
        Set GetButton = MyBar.Controls(TagOrIndex)
    Case Else
        Set GetButton = ButtonFindByTag(MyBar, TagOrIndex)
    End Select
End Function

Public Function GetControlText(TagOrIndex As Variant) As String
    Dim out As String
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    Select Case oAC.Type
    Case msoControlEdit, msoControlDropdown, msoControlComboBox
        out = oAC.Text
    Case Else   ' msoControlButton, msoControlPopup
        out = oAC.Caption
    End Select
    
    Set oAC = Nothing
    GetControlText = out
End Function

Public Function SetControlText(TagOrIndex As Variant, ByVal Text As String) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    Dim Index As Long
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then
        out = False
    Else
        Select Case oAC.Type
        Case msoControlEdit
            oAC.Text = Text
        Case msoControlDropdown
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                ListAddHistory oAC, Text
            Else
                oAC.ListIndex = Index
            End If
        Case msoControlComboBox
            Index = ListFindIndex(oAC, Text)
            If Index = 0 Then
                oAC.Text = Text
                ComboAddHistory oAC
            Else
                oAC.ListIndex = Index
            End If
        Case Else
            oAC.Caption = Text
        End Select
        Set oAC = Nothing
        out = True
    End If
    
    SetControlText = out
End Function

Public Function GetControlState(TagOrIndex As Variant) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        ' return True when the button is pushed down
        out = IsButtonStateOn(oAC)
    End If
    
    Set oAC = Nothing
    GetControlState = out
End Function

Public Function SetControlState(TagOrIndex As Variant, ByVal State As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Type = msoControlButton Then
        If IsButtonStateOn(oAC) <> State Then
            If State Then
                ButtonSwitchOn oAC
            Else
                ButtonSwitchOff oAC
            End If
            ' return True when the status is strictly changed
            out = True
        End If
    End If
    
    Set oAC = Nothing
    SetControlState = out
End Function

Public Function GetControlVisible(TagOrIndex As Variant) As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    GetControlVisible = oAC.Visible
End Function

Public Function SetControlVisible(TagOrIndex As Variant, ByVal Visible As Boolean) As Boolean
    Dim out As Boolean
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    out = False
    If oAC.Visible <> Visible Then
        oAC.Visible = Visible
        ' return True when the visible is strictly changed
        out = True
    End If
    
    SetControlVisible = out
End Function

Public Function IncControlWidth(TagOrIndex As Variant, ByVal Width As Long) As Long
    Dim out As Long
    Dim oAC As Office.CommandBarControl
    
    Set oAC = GetButton(TagOrIndex)
    If oAC Is Nothing Then Exit Function
    
    On Error Resume Next
    oAC.Width = oAC.Width + Width
    ' return the width accepted (tips: setting 0 to width makes it becomes default)
    out = oAC.Width
    
    IncControlWidth = out
End Function


'=== button tools for mybar end ===
'=== helper functions begin ===


Public Function NormalizeArray(x As Variant) As Variant
    On Error Resume Next
    Dim out() As Variant
    Dim i As Long
    Dim L1 As Long
    Dim L2 As Long
    Dim U1 As Long
    Dim U2 As Long
    
    L1 = 0
    L2 = 0
    U1 = -1
    U2 = -1
    
    L1 = LBound(x)
    L2 = LBound(x, 2)   ' error unless 2 dimensions
    U1 = UBound(x)
    U2 = UBound(x, 2)   ' error unless 2 dimensions
    
    If U1 < L1 Then
        NormalizeArray = Array()
        Exit Function
    End If
    
    If U2 = -1 Then
        ReDim out(0 To U1 - L1)
        For i = 0 To UBound(out)
            out(i) = x(i + L1)
        Next
    Else
        ReDim out(0 To U2 - L2)
        For i = 0 To UBound(out)
            out(i) = x(L1, i + L2)
            ' we pick up the 1st line only
        Next
    End If
    
    NormalizeArray = out
End Function

Public Function StackToArray(pan As Variant) As Variant
    Dim out() As Variant
    Dim x As Variant
    Dim i As Long
    Dim Counter As Long
    
    x = Empty
    Counter = 0
    Do Until IsEmpty(pan)
        x = Array(pan(0), x)
        pan = pan(1)
        Counter = Counter + 1
    Loop
    
    If Counter = 0 Then
        StackToArray = Array()
        Exit Function
    End If
    
    ReDim out(0 To Counter - 1)
    i = 0
    Do Until IsEmpty(x)
        out(i) = x(0)
        x = x(1)
        i = i + 1
    Loop
    
    StackToArray = out
End Function


'=== helper functions end ===

'}}}

'module
'   name;FindPriceInMatrixSum
'{{{
Option Explicit

Private TargetData As Range
Private FlagColumn As Range
Private RuleArea As Range

Public Sub Set_data(Target As Object, Helper As ToolBarV2, oAC As Object, oAC_flag As Object, oAC_rulearea As Object)
    ' use magic word to determin data rows and flag column, or use default numbers
    Const FlagName = "フラグ"
    Const FlagCol = 2
    Const ColOffset = 1
    Const DataRow = 3
    
    Dim MagicCell As Range
    Dim RowBegin As Long
    Dim FlagAt As Long
    
    If TypeName(Target) <> "Range" Then Exit Sub
    
    If IsASingleCell(Target) Then
        ' assume the top-left cell is given
        Set TargetData = CurrentRegionKeepTopLeft(Target)
    Else
        Set TargetData = Target
    End If
    
    TargetData.Select
    oAC.Caption = TargetData.Address(False, False, xlA1, True)
    Helper.ButtonSwitchOn oAC
    
    ' try to find the flag column
    Set MagicCell = TargetData.EntireColumn.Cells.Find(What:=FlagName, LookAt:=xlPart, SearchOrder:=xlByRows)
    If MagicCell Is Nothing Then
        FlagAt = FlagCol
    Else
        FlagAt = MagicCell.Column - TargetData.Cells(1).Column + 1
    End If
    
    ' set standard flag column
    Set_flag TargetData.Columns(FlagAt), Helper, oAC_flag
    
    ' set standard rulearea
    Set RuleArea = ParseRangeName(oAC_rulearea.Parameter, ThisWorkbook)
    oAC_rulearea.Caption = oAC_rulearea.Parameter
    If RuleArea Is Nothing Then
        Helper.ButtonSwitchOff oAC_rulearea
    Else
        Helper.ButtonSwitchOn oAC_rulearea
    End If
End Sub

Public Sub Set_flag(Target As Object, Helper As ToolBarV2, oAC As Object)
    If TypeName(Target) <> "Range" Then Exit Sub
    Set FlagColumn = Target.EntireColumn
    Helper.ButtonSwitchOn oAC
    oAC.Caption = FlagColumn.Address(False, False, xlA1, False)
End Sub

Public Sub Set_rulearea(Target As Object, Helper As ToolBarV2, oAC As Object)
    If TypeName(Target) <> "Range" Then Exit Sub
    Set RuleArea = Target
    Helper.ButtonSwitchOn oAC
    oAC.Caption = RuleArea.Address(False, False, xlA1, True)
End Sub

Public Sub Exec_sumup()
    Dim Ws As Worksheet
    Dim Ra As Range
    
    On Error GoTo HasError
    
    If IsNotReady Then Exit Sub
    Set Ws = GetNewSheet(TargetData.Worksheet)
    SetEntireFormat Ws
    
    Application.ScreenUpdating = False
    Set Ra = MakeList(Ws, RuleArea, FlagColumn, TargetData)
    SumUp Ra
    Application.ScreenUpdating = True
    Exit Sub
    
HasError:
    Application.ScreenUpdating = True
    MsgBox "集計できません。 " & Err.Number & " " & Err.Description, _
            vbExclamation Or vbOKOnly, "Exec_sumup"
End Sub

Private Function GetNewSheet(Ws As Worksheet) As Worksheet
    Set GetNewSheet = Ws.Parent.Worksheets.Add(Before:=Ws)
End Function

Private Function MakeList(Ws As Worksheet, Rule As Range, Flag As Range, Data As Range) As Range
    Dim x As Variant
    Dim ARow As Range
    Dim R As Long
    Dim FlagValue As Variant
    Dim DataOffset As Long
    Dim RuleNames As Variant
    
    Const ListOffset = 1
    
    DataOffset = Data.Cells(1).Row - 1
    R = 1
    RuleNames = FindPriceInMatrix.GetRuleNames(Rule)
    SetListTitle Ws.Rows(1).Cells, RuleNames
    
    For Each ARow In Data.Rows
        FlagValue = Flag.Cells(ARow.Row).Value
        x = FindPriceInMatrix.FindPriceRow(Rule, FlagValue, ARow)
        CopyListData Ws.Rows(R + ListOffset).Cells, x, RuleNames, R + DataOffset, ARow, FlagValue
        R = R + 1
    Next
    
    Set MakeList = Ws.UsedRange
End Function

Private Sub CopyListData(Ra As Range, FoundRule As Variant, RuleNames As Variant, _
        RowNumber As Long, DataRow As Range, Flag As Variant)
    Dim RuleNo As Long
    Dim RuleRow As Variant
    Dim i As Long
    Dim MaxCol As Long
    Dim buff() As Variant
    
    Const RuleOutputColBegin = 5
    
    RuleNo = FoundRule(0)
    RuleRow = FoundRule(1)
    ' RuleRow : array(price, flag, array of rules)
    
    If RuleNo = 0 Then
        MaxCol = 4
        ReDim buff(1 To MaxCol)
        buff(3) = 0
        buff(4) = "×"  ' indicates clearly no flags available
    Else
        'Debug.Assert Flag = RuleRow(1)
        MaxCol = RuleOutputColBegin + UBound(RuleNames)
        ReDim buff(1 To MaxCol)
        buff(3) = RuleRow(0)    ' price
        buff(4) = RuleRow(1)    ' flag, not in data but in rule
        For i = 0 To UBound(RuleNames)
            buff(RuleOutputColBegin + i) = IIf(IsEmpty(RuleRow(2)(i)), vbNullString, RuleNames(i))
        Next
    End If
    buff(1) = RuleNo
    buff(2) = RowNumber     ' to be count
    
    Ra.Worksheet.Range(Ra.Columns(1), Ra.Columns(MaxCol)).Value = buff
    
    ' verbose
    i = RuleOutputColBegin + UBound(RuleNames) + 2
    Ra.Worksheet.Range(Ra.Columns(i), Ra.Columns(i + DataRow.Columns.Count - 1)).Value = DataRow.Value
End Sub

Private Sub SetListTitle(Ra As Range, RuleNames As Variant)
    Dim cEnd As String
    cEnd = Ra(5 + UBound(RuleNames)).Address(False, False, xlA1, False)
    
    Ra.Range("E1:" & cEnd).Value = RuleNames
    Ra.Range("A1:" & cEnd).Interior.ColorIndex = 43 ' ライム
    
    With Ra.Range("A1:D1")
        .Value = Array("ルール行番号", "件数", "単価", "フラグ")
        .Interior.ColorIndex = 4                    ' 明るい緑
        .HorizontalAlignment = xlHAlignCenter
    End With
End Sub

Private Sub SetEntireFormat(Ws As Worksheet)
    Ws.Cells.ShrinkToFit = True
    Ws.Columns("D").HorizontalAlignment = xlHAlignCenter
End Sub

Private Sub SumUp(DataList As Range)
    Dim Ra As Range
    Dim R As Long
    Dim C As Long
    Dim RememberFormula As String
    
    ' sort
    DataList.Sort _
        Key1:=Range("A2"), Order1:=xlAscending, _
        Key2:=Range("B2"), Order2:=xlAscending, _
        Header:=xlYes
    
    ' summarize
    DataList.Cells(1).CurrentRegion.Subtotal _
        GroupBy:=1, Function:=xlCount, TotalList:=Array(2), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    DataList.Worksheet.Outline.ShowLevels RowLevels:=2
    
    ' adjust total
    Set Ra = DataList.Cells(1).CurrentRegion.SpecialCells(xlVisible)
    For R = 2 To Ra.Areas.Count
        With Ra.Areas(R)
            RememberFormula = .Range("B1").FormulaR1C1
            .Rows(1).FormulaR1C1 = "=R[-1]C"
            .Range("B1").FormulaR1C1 = RememberFormula
            RememberFormula = Replace(RememberFormula, 3, 1, 1, 1)  ' counta(3) to average(1)
            .Range("C1").FormulaR1C1 = RememberFormula
            
            .Rows(1).NumberFormat = "General;;"                     ' avoid to show 0(empty)
            .Range("A1,D1").NumberFormat = "General"
            .Columns("B").NumberFormat = "General""件"""
            .Range("C1").NumberFormat = "General""円"""
            
            .Interior.ColorIndex = 40       ' ベージュ
        End With
    Next
End Sub

Private Function IsNotReady() As Boolean
    Dim x As Variant
    
    IsNotReady = True
    For Each x In Array(Array(TargetData, "データ範囲"), _
                        Array(FlagColumn, "フラグ列"), _
                        Array(RuleArea, "ルールエリア"))
        If Not EnsureTheRange(x(0), x(1)) Then Exit Function
    Next
    IsNotReady = False
End Function

Private Function EnsureTheRange(ByVal Ra As Range, Optional ByVal Tag As String = "") As Boolean
    On Error GoTo HasError
    
    If Ra Is Nothing Then GoTo NotSpecified
    Tag = Ra.Address    ' make sure the range exists
    EnsureTheRange = True
    Exit Function
    
HasError:
    If Tag <> "" Then Tag = Tag & " Error " & Err.Number & " " & Err.Description
    Err.Clear
NotSpecified:
    If Tag <> "" Then MsgBox "範囲が不正: " & Tag, vbExclamation Or vbOKOnly
    EnsureTheRange = False
End Function

Private Function ParseRangeName(RangeName As String, Wb As Workbook) As Range
    Dim AddressString As String
    Dim Parsing As Variant
    
    On Error GoTo HasError
    
    AddressString = Wb.Names(RangeName).RefersTo
    '=PriceMatrix!$A$1:$N$30 is expected
    If Left(AddressString, 1) = "=" Then AddressString = Mid(AddressString, 2)
    Parsing = Split(AddressString, "!", 2)
    Set ParseRangeName = Wb.Worksheets(Parsing(0)).Range(Parsing(1))
    Exit Function
    
HasError:
    Set ParseRangeName = Nothing
End Function

Private Function CurrentRegionKeepTopLeft(Ra As Range) As Range
    Dim Address As Variant
    
    Address = Split(Ra.CurrentRegion.Address(False, False, xlA1, False), ":", 2)
    If UBound(Address) = 0 Then
        Set CurrentRegionKeepTopLeft = Ra
    Else
        Set CurrentRegionKeepTopLeft = Ra.Worksheet.Range(Ra.Cells(1).Address(False, False, xlA1, False) & ":" & Address(1))
    End If
End Function

' this function is required to avoid overflow errors on excel 2007 Cells.Count
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

'}}}

'code
'   name;ThisWorkbook
'{{{
Option Explicit

Private Sub Workbook_Open()
    BaumMain.BarInitialize
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    BaumMain.BarTerminate
End Sub

'}}}

'ssf-end


```