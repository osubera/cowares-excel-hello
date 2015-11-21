

# Introduction #

  * generate a square excel sheet in inches

## 概要 ##
  * エクセルでセル幅や高さをミリ単位で設定した方眼を作る

# Details #


## 説明 ##
  * エクセルワークシートの行の高さや列の幅を、mm単位で指定します。
  * 印刷される大きさを細かく指定したい場合に便利です。
  * 単位は、ミリ・センチ・インチ・寸で指定できます。方眼紙のようなシートが簡単につくれます。

# Downloads #

  * [downloads / ダウンロード](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=square_cells)


# Snapshots #

  * Excel 2000 toolbar
> > ![http://3.bp.blogspot.com/-V3KcnII8enk/TbKgZl6oNBI/AAAAAAAAAE4/DXC0YowrM2A/s1600/shot1.png](http://3.bp.blogspot.com/-V3KcnII8enk/TbKgZl6oNBI/AAAAAAAAAE4/DXC0YowrM2A/s1600/shot1.png)
  * Excel 2007 ribbon
> > ![http://4.bp.blogspot.com/-IENvLy4xxYQ/TbKgaIyvraI/AAAAAAAAAE8/-b7FTtiKX7Y/s1600/shot2.png](http://4.bp.blogspot.com/-IENvLy4xxYQ/TbKgaIyvraI/AAAAAAAAAE8/-b7FTtiKX7Y/s1600/shot2.png)


# How to use #


## 使い方 ##

### 単位変換の方法 ###

  * 変換の基礎関数
```
' mm をセル幅に換算する。
' ポイント幅で返す。
    mm2Haba = x / 5 * 13.5 * 1.147028154 * 0.98167218

' mm をセル高さに換算する。
    mm2Takasa = x / 5 * 13.5 * 1.147028154
```
  * 幅、高さとも、ポイントに換算する。エクセルでは幅はポイントでなく 0 の文字数を使用するが、基礎関数はそれを区別せず、高さで使うポイントに変換する。
  * ただし、通常のハードウェアで縦と横が完全に一致することは期待できないため、幅の側に調整比率をかけている。
  * 上記の数値は、標準セル高さの13.5ポイントが、約5mmという期待値から実際にどの程度ずれているか、を実測して決めている。
  * セル幅をポイントで直接変更はできないが、文字数指定で変更した後のポイント幅を知ることはできる。
  * セル幅はとポイントの関係は線型ではないが、比の期待値は 6 から 9.375 程度になる。
  * これを元に、概算値から補間による修正を行い、要求されたセル幅を得る。
  * この数値はあくまで、１つのマシンでの実測値なので、実際に使用するマシンでの調整結果を各々に乗じる必要がある。

# Code #

```
'ssf-begin
';

'workbook
'   name;square_cells_2k.xls/F3SquareCells

'book-identity
'  title;きっちり方眼
'  description;セル幅や高さをミリ単位で設定した方眼を作る

'require

'worksheet
'   name;きっちり方眼/BaumMain

'cells-formula
'  address;A1:B10
'         ;名称
'         ;square_cells
'         ;コメント
'         ;セル幅や高さをミリ単位で設定した方眼を作る
'         ;comment
'         ;generate a square sheet in inches
'         ;著作権
'         ;="Copyright (C) " &R[3]C & "-" & YEAR(R[5]C) & " " & R[2]C
'         ;ライセンス
'         ;自律, 自由, 公正, http://cowares.nobody.jp
'         ;作者
'         ;Tomizono - kobobau.com
'         ;初版
'         ;2002
'         ;配布元
'         ;http://cowares.blogspot.com/search/label/baum
'         ;更新
'         ;40656.6894444444
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
'  address;A15:M19
'         ;対象
'         ;高さだけや幅だけを変更することも、両方を同時に変更することもできます。
'         ;target
'     skip;1
'         ;3
'     skip;1
'   repeat;2
'         ;1
'     skip;1
'         ;両方
'         ;幅
'         ;高さ
'     skip;1
'         ;大きさ
'         ;大きさを指定します。
'         ;size
'     skip;1
'         ;4
'     skip;4
'         ;10
'     skip;3
'         ;単位
'         ;単位を指定します。
'         ;unit
'     skip;1
'         ;3
'     skip;1
'         ;1
'     skip;2
'         ;mm
'         ;cm
'         ;inch
'         ;寸
'         ;高精度
'         ;高精度と標準を切り替えます。
'         ;high
'     skip;1
'         ;1
'         ;2
'     skip;2
'         ;F3SquareCells.BaumMain.OnButtonToggle
'     skip;4
'         ;=R1C2 & " について"
'         ;このシートを表示する。
'         ;about
'     skip;1
'         ;1
'         ;2
'     skip;1
'         ;1
'  address;B24
'         ;エクセルブック、ワード、アクセスのカスタムメイドやウェブシステムの開発などをリーズナブルな価格で承っております。
'  address;C27:G33
'         ;導入する利点
'     skip;1
'         ;発注について
'     skip;2
'         ;事業者で
'     skip;1
'         ;発注から公開まで
'     skip;1
'         ;プライバシー
'         ;大企業で
'     skip;1
'         ;料金
'     skip;1
'         ;取引上の注意
'         ;公益法人で
'     skip;1
'         ;公開（納品）
'     skip;2
'         ;個人で
'     skip;1
'         ;基本的なルール
'     skip;2
'         ;IT企業で
'     skip;1
'         ;細かいルール
'     skip;4
'         ;なぜ無料ソフトを買う？
'  address;C35
'         ;すぐに注文する
'  address;B39
'         ;簡単な指示、安い料金、早い結果
'  address;B41
'         ;料金 - コ・ウェア・ライセンスのシステム開発
'  address;B43
'         ;基本料金表
'  address;B45:B46
'         ;３つのサイズから選ぶだけで簡単。
'         ;前払いが原則。
'  address;B48:D50
'         ;Sサイズ
'         ;1,000円
'         ;（税込 1,050円）
'         ;Mサイズ
'         ;10,000円
'         ;（税込 10,500円）
'         ;Lサイズ
'         ;100,000円
'         ;（税込 105,000円）
'  address;B52
'         ;サイズとは
'  address;B54
'         ;開発にかかる時間や難易度を、おおまかに３つのサイズで分類します。
'  address;B56
'         ;Sサイズ
'  address;B58:C62
'         ;1. 小型、 Small 、 partial
'     skip;1
'         ;2. ちょっとしたコード素片や、ワークシートの一部分など。
'     skip;1
'         ;3. 手に負えない、書き方のわからないコードだけを知りたいときに。
'     skip;1
'         ;4. 例）数行の VBA コード。次のリンク先の１つ目のコード
'     skip;2
'         ;http://code.google.com/p/cowares-excel-hello/wiki/hello_key_value
'  address;B64
'         ;Mサイズ
'  address;B66:C70
'         ;1. 中型、 Medium 、 functional
'     skip;1
'         ;2. 完成した関数やワークシート。
'     skip;1
'         ;3. 単一の機能が、とりあえず動くレベルのものが欲しいときに。
'     skip;1
'         ;4. 例）マクロを実行すれば一つの動作を行う VBA コード。
'     skip;2
'         ;http://code.google.com/p/cowares-excel-hello/wiki/annual_list
'  address;B72
'         ;Lサイズ
'  address;B74:B77
'         ;1. 大型、 Large 、 integrated
'         ;2. 実用的なアプリケーション。
'         ;3. 複数の機能や、条件設定による動作切り替えや画面遷移も含むときに。
'         ;4. 例）ユーザーインターフェースを持ち、ツールとして利用できる。
'  address;B81
'         ;公開 - コ・ウェア・ライセンスのシステム開発
'  address;B83
'         ;公開が納品です
'  address;B85:B92
'         ;所定の公開場所に成果物をアップロードする方法を採ります。
'         ;仕様を決める段階から公開URLを使います。
'         ;公開のタイミングで連絡はしますが、ファイル添付などはしません。
'         ;公開先からのセルフダウンロードでお願いします。
'         ;本人はもちろん、同僚や友達、その他大勢の人がダウンロードして利用できます。
'         ;マクロコードをテキストで公開するので、セキュリティの強い職場で、マクロ付きブックのダウンロード規制がある環境でも心配ありません。
'         ;公開後のコード修正等、追加情報も当該URLから派生していきます。
'         ;URLは永久に変わらないものではありません。
'  address;B94
'         ;主な公開先URL
'  address;B96:B97
'         ;http://cowares.blogspot.com
'         ;http://code.google.com/p/cowares-excel-hello/
'  address;B101
'         ;なぜ無料ソフトを買うのか？ - コ・ウェア・ライセンスのシステム開発
'  address;B103
'         ;そのお金は何に払っているのでしょうか
'  address;B105:G105
'         ;無料のもの
'     skip;1
'         ;買うもの
'     skip;2
'         ;買わないもの
'  address;B107:G110
'         ;ライセンス
'     skip;1
'         ;エンジニアの働き
'     skip;2
'         ;保証
'         ;コピー
'     skip;1
'         ;世界への貢献
'     skip;2
'         ;役員の働き
'     skip;5
'         ;事務員の働き
'     skip;5
'         ;営業スマイル
'  address;C112:D113
'     skip;1
'         ;コ・ウェアの料金
'         ;通常のシステム開発で払うお金
'  address;C117
'         ;すぐに注文する

'cells-numberformat
'  address;B9
'         ;m/d/yyyy h:mm

'cells-width
'   unit;zero
'  address;B1
'         ;15.5

'cells-height
'  address;A24
'         ;14.25
'  address;A35
'         ;24.75
'  address;A41
'         ;21
'  address;A43
'         ;14.25
'  address;A47:A48
'   repeat;2
'         ;14.25
'  address;A50:A52
'   repeat;3
'         ;14.25
'  address;A81
'         ;21
'  address;A83
'         ;14.25
'  address;A94
'         ;14.25
'  address;A101
'         ;21
'  address;A103
'         ;14.25
'  address;A117
'         ;24.75

'cells-background-color
'  address;A24:M24
'   repeat;13
'         ;#FF6600
'  address;C27:H27
'   repeat;2
'         ;#FFCC99
'   repeat;4
'         ;#CCFFCC
'  address;C35:F35
'   repeat;4
'         ;#99CC00
'  address;B41:L41
'   repeat;11
'         ;#333399
'  address;B48:B50
'         ;#CCFFCC
'         ;#FFFF99
'         ;#FFCC99
'  address;B56
'         ;#CCFFCC
'  address;B64
'         ;#FFFF99
'  address;B72
'         ;#FFCC99
'  address;B81:L81
'   repeat;11
'         ;#333399
'  address;B101:L101
'   repeat;11
'         ;#333399
'  address;B105:H105
'   repeat;2
'         ;#FF99CC
'   repeat;2
'         ;#CCFFCC
'   repeat;3
'         ;#FF99CC
'  address;B112:H113
'     skip;2
'   repeat;2
'         ;#00FF00
'     skip;3
'   repeat;7
'         ;#FF00FF
'  address;C117:F117
'   repeat;4
'         ;#99CC00

'cells-color
'  address;C28:G33
'         ;#0000FF
'     skip;1
'         ;#0000FF
'     skip;1
'   repeat;2
'         ;#0000FF
'     skip;1
'         ;#0000FF
'     skip;1
'   repeat;2
'         ;#0000FF
'     skip;1
'         ;#0000FF
'     skip;2
'         ;#0000FF
'     skip;1
'         ;#0000FF
'     skip;2
'         ;#0000FF
'     skip;1
'         ;#0000FF
'     skip;4
'         ;#0000FF
'  address;C35:F35
'   repeat;4
'         ;#0000FF
'  address;B39
'         ;#FF00FF
'  address;B41:L41
'   repeat;11
'         ;#FFCC00
'  address;C61:J62
'     skip;2
'   repeat;14
'         ;#0000FF
'  address;C70:H70
'   repeat;6
'         ;#0000FF
'  address;B81:L81
'   repeat;11
'         ;#FFCC00
'  address;B86
'  address;B89:B90
'  address;B96:B97
'   repeat;2
'         ;#0000FF
'  address;B101:L101
'   repeat;11
'         ;#FFCC00
'  address;C117:F117
'   repeat;4
'         ;#0000FF

'cells-font-size
'  address;B24
'         ;12
'  address;C35:F35
'   repeat;4
'         ;12
'  address;B41
'         ;18
'  address;B43
'         ;12
'  address;B52
'         ;12
'  address;B81
'         ;18
'  address;B83
'         ;12
'  address;B94
'         ;12
'  address;B101
'         ;18
'  address;B103
'         ;12
'  address;C117:F117
'   repeat;4
'         ;12

'cells-font-bold
'  address;B24
'         ;yes
'  address;C35:F35
'   repeat;4
'         ;yes
'  address;B41
'         ;yes
'  address;B43
'         ;yes
'  address;B48:B50
'   repeat;3
'         ;yes
'  address;B52
'         ;yes
'  address;B56
'         ;yes
'  address;B64
'         ;yes
'  address;B72
'         ;yes
'  address;B81
'         ;yes
'  address;B83
'         ;yes
'  address;B94
'         ;yes
'  address;B101
'         ;yes
'  address;B103
'         ;yes
'  address;B105:G105
'         ;yes
'     skip;1
'   repeat;2
'         ;yes
'     skip;1
'         ;yes
'  address;C117:F117
'   repeat;4
'         ;yes

'cells-h-align
'  address;C35:F35
'   repeat;4
'         ;center
'  address;C48:C50
'   repeat;3
'         ;right
'  address;C61:J62
'     skip;2
'   repeat;6
'         ;center
'   repeat;7
'         ;left
'         ;center
'  address;C70:H70
'   repeat;6
'         ;left
'  address;D105:E105
'   repeat;2
'         ;center
'  address;D112:E112
'   repeat;2
'         ;center
'  address;C117:F117
'   repeat;4
'         ;center

'cells-v-align
'  address;C35:F35
'   repeat;4
'         ;center
'  address;B43:J77
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;4
'         ;center
'     skip;1
'   repeat;6
'         ;center
'     skip;1
'   repeat;14
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;4
'   repeat;5
'         ;center
'     skip;4
'   repeat;6
'         ;center
'     skip;2
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'     skip;3
'   repeat;6
'         ;center
'  address;B81:G97
'   repeat;102
'         ;center
'  address;B101:B103
'   repeat;3
'         ;center
'  address;B105:G113
'   repeat;11
'         ;center
'     skip;1
'   repeat;28
'         ;center
'     skip;2
'   repeat;4
'         ;center
'     skip;2
'   repeat;4
'         ;center
'  address;C117:F117
'   repeat;4
'         ;center

'worksheet
'   name;微調整/Calib

'cells-formula
'  address;A1:E1
'         ;戻る
'     skip;3
'         ;10
'  address;A4:Y6
'     skip;25
'         ;10
'     skip;29
'         ;1
'     skip;1
'         ;2
'     skip;1
'         ;3
'     skip;1
'         ;4
'     skip;1
'         ;5
'     skip;1
'         ;6
'     skip;1
'         ;7
'     skip;1
'         ;8
'     skip;1
'         ;9
'         ;10
'     skip;1
'         ;cm
'  address;F8
'         ;2
'  address;F10:I18
'         ;3
'     skip;2
'         ;微調整の方法
'     skip;3
'         ;1. このシートを印刷。
'         ;4
'     skip;2
'         ;2. 上と左にある、10cmの太線を、物差しで測る。
'     skip;3
'         ;3. 測った長さを、それぞれE1とA5セルに記入。
'         ;5
'     skip;2
'         ;  (横線の長さが10.25cmなら、E1に10.25と入力)
'     skip;3
'         ;4. 設定ボタンを押す。
'         ;6
'     skip;2
'         ;5. 再度、印刷して確認する。
'     skip;3
'         ;6. ブックを保存する。
'         ;7
'  address;F20:L20
'         ;8
'     skip;5
'         ;設定
'  address;F22
'         ;9
'  address;E24:G24
'         ;10
'     skip;1
'         ;cm
'  address;F38:G38
'         ;1
'         ;cm
'  address;F40
'         ;0.5
'  address;F45:G45
'         ;1
'         ;inch
'  address;F50
'         ;0.5
'  address;F55:G55
'         ;1
'         ;寸
'  address;Z57:Z58
'   repeat;2
'         ;1

'cells-numberformat
'  address;F40:G43
'   repeat;8
'         ;# ?/?
'  address;F50:G53
'   repeat;8
'         ;# ?/?
'  address;Z57:Z58
'   repeat;2
'         ;""

'cells-width
'   unit;zero
'  address;A1:Z1
'   repeat;26
'         ;24.75

'cells-height
'  address;A27:A36
'   repeat;10
'         ;31
'  address;A39:A58
'   repeat;10
'         ;78.5
'   repeat;10
'         ;94

'cells-background-color
'  address;A1:F2
'   repeat;2
'         ;#CCFFCC
'     skip;2
'   repeat;2
'         ;#FFFF99
'   repeat;2
'         ;#CCFFCC
'     skip;2
'   repeat;2
'         ;#FFFF99
'  address;A5:B6
'   repeat;4
'         ;#FFFF99
'  address;H10:Y18
'   repeat;162
'         ;#CCFFFF
'  address;L20:O21
'   repeat;8
'         ;#FFFF99
'  address;Z57:Z58
'   repeat;2
'         ;#FF0000
'  address;Z60
'         ;#FF0000

'cells-color
'  address;A1:B2
'   repeat;4
'         ;#0000FF
'  address;I10
'         ;#0000FF
'  address;L20:O21
'   repeat;8
'         ;#FF0000

'cells-font-size
'  address;A1:Z60
'   repeat;243
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;25
'         ;110
'     skip;1
'   repeat;1108
'         ;110

'cells-font-bold
'  address;I10
'         ;yes
'  address;L20:O21
'   repeat;8
'         ;yes

'cells-font-italic
'  address;I10
'         ;yes

'cells-h-align
'  address;A1:B2
'   repeat;4
'         ;center
'  address;L20:O21
'   repeat;8
'         ;center

'cells-v-align
'  address;A1:B2
'   repeat;4
'         ;center
'  address;L20:O21
'   repeat;8
'         ;center

'cells-name
'       ;=きっちり方眼!R15C1
'       ;きっちり方眼!_ButtonCaption
'       ;=きっちり方眼!R3C2
'       ;きっちり方眼!_Comment
'       ;=きっちり方眼!R6C2
'       ;きっちり方眼!_Contributor
'       ;=きっちり方眼!R4C2
'       ;きっちり方眼!_Copyright
'       ;=きっちり方眼!R5C2
'       ;きっちり方眼!_License
'       ;=きっちり方眼!R2C2
'       ;きっちり方眼!_LocalComment
'       ;=きっちり方眼!R1C2
'       ;きっちり方眼!_PublicName
'       ;=きっちり方眼!R7C2
'       ;きっちり方眼!_Since
'       ;=きっちり方眼!R10C2
'       ;きっちり方眼!_Tag
'       ;=きっちり方眼!R9C2
'       ;きっちり方眼!_Timestamp
'       ;=きっちり方眼!R8C2
'       ;きっちり方眼!_Url

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
    VBProjectName = "F3SquareCells"
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
    If Helper Is Nothing Then
        BarInitialize
        MsgBox "ツールバーを修復しました。もう一度操作してください。", vbExclamation, BaumMain.Name
    Else
        Helper.OnButtonToggle
    End If
End Sub

' followings need to be public, because they are called from outside by the Helper
' we also can copy the Helper.BarMain code here, and let the followings be private.

Public Sub Menu_target(oAC As Object)
End Sub

Public Sub Menu_size(oAC As Object)
    Helper.ComboAddHistory oAC, False
    SquareCells.Touch oAC, MakeData(oAC)
End Sub

Public Sub Menu_unit(oAC As Object)
End Sub

Public Sub Menu_high(oAC As Object)
End Sub

Public Sub Menu_about(oAC As Object)
    If ThisWorkbook.IsAddin Then
        Dim Wb As Workbook
        Set Wb = Workbooks.Add
        Me.Copy Before:=Wb.Sheets(1)
        Wb.Saved = True
        Set Wb = Nothing
    Else
        Me.Activate
    End If
End Sub

Private Function MakeData(oAC As Object) As Variant
    MakeData = Array( _
        Val(Helper.GetControlText("size")), _
        Helper.GetButton("target").ListIndex, _
        Helper.GetButton("unit").ListIndex, _
        Helper.GetControlState("high"))
End Function

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

'module
'   name;SquareCells
'{{{
Option Explicit

' うちらのシート(きっちり方眼)4
' Copyright (C) 2002 Tomizono
' 2002.4.18

' 2002.4.19
' 高精度に対応
' 幅補正改善

' 2011.4.23
' オープンライセンスに変更
' 自律, 自由, 公正, http://cowares.nobody.jp
' ツールバーV2に移植
' Excel 2007 に対応

#Const EnableCalibration = True     ' キャリブレーション機能を使う

Const HOCHRES As Double = 10        ' 高精度の倍率
Const BAUCAP As String = "ぱそ工房ばう"              ' 共通のタイトルバー

Private Sub ExtractData(Data As Variant, ByRef Size As Double, ByRef Target As Long, _
        ByRef Unit As Long, ByRef High As Boolean)
    Size = Data(0)
    Target = Data(1)
    Unit = Data(2)
    High = Data(3)
End Sub

Public Sub Touch(oAC As Object, Data As Variant)
    Dim Size As Double
    Dim iTaisho As Long
    Dim Unit As Long
    Dim High As Boolean
    Dim x As Double
    Dim Ra As Range
    
    On Error GoTo Err1
    Application.ScreenUpdating = False
    
    ' 条件読み出し
    ExtractData Data, Size, iTaisho, Unit, High
    x = Kanzan2mm(Size, Unit)          ' mm
    If x <= 0 Then
        MsgBox "数値が小さすぎて扱えません。", , BAUCAP
        Exit Sub
    End If
    Set Ra = Selection
    If High Then   ' 高精度
        setupHochPage Ra.Worksheet
        x = x * HOCHRES
    End If
    
    ' 設定
    Select Case iTaisho
    Case 1, 2       ' 幅
        applyWidth2Range Ra, mm2Haba(x)
    End Select
    
    Select Case iTaisho
    Case 1, 3       ' 高さ
        applyHeight2Range Ra, mm2Takasa(x)
    End Select
    Set Ra = Nothing
    
    Application.ScreenUpdating = True
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox Err.Description, , BAUCAP
End Sub

Private Sub applyHeight2Range(Ra As Range, x As Double)
' 指定範囲のセル高さを変更する。
' 単一セルなら、シート全体が対象
    On Error GoTo Err1
    If IsASingleCell(Ra) Then
        Ra.Worksheet.Cells.RowHeight = x
    Else
        Ra.EntireRow.RowHeight = x
    End If
    Exit Sub
Err1:
    MsgBox Err.Description, , BAUCAP
End Sub

Private Function determinColumnWidth(Ra As Range, xWidth As Double) As Double
' Width をもらい ColumnWidth を決める。
' テストに使える単一セルをもらう。
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim y1 As Double, y2 As Double, y3 As Double
    Const DZero As Double = 0.000001
    Const EZero As Double = 0.255
    Dim i As Long
    Const k1 As Double = 9.375
    Const k2 As Double = 6
    On Error GoTo Err1
    y1 = xWidth / k1
    y2 = xWidth / k2
    determinColumnWidth = (y1 + y2) / 2     ' デフォルト値
    With Ra
        .ColumnWidth = y1
        x1 = .Width
        .ColumnWidth = y2
        x2 = .Width
    End With
    For i = 1 To 3
        If Abs(x2 - x1) < DZero Then Exit For               ' 収束
        y3 = y1 + (xWidth - x1) * (y2 - y1) / (x2 - x1)     ' 線型補間
        With Ra
            .ColumnWidth = y3
            x3 = .Width
        End With
        If Abs(xWidth - x3) < EZero Then Exit For            ' 収束
        If Abs(xWidth - x1) > Abs(xWidth - x2) Then
            y1 = y3
            x1 = x3
        Else
            y2 = y3
            x2 = x3
        End If
    Next
    determinColumnWidth = y3
    Exit Function
Err1:
    'MsgBox Err.Description, , BAUCAP
End Function

Private Sub applyWidth2Range(Ra As Range, x As Double)
' 指定範囲のセル幅を変更する。
' 単一セルなら、シート全体が対象
' もらうのはポイント幅
    Dim y As Double
    On Error GoTo Err1
    y = determinColumnWidth(Ra.Cells(1), x)
    If IsASingleCell(Ra) Then
        Ra.Worksheet.Cells.ColumnWidth = y
    Else
        Ra.EntireColumn.ColumnWidth = y
    End If
    Exit Sub
Err1:
    MsgBox Err.Description, , BAUCAP
End Sub

Private Function mm2Haba(x As Double) As Double
' mm をセル幅に換算する。
' ポイント幅で返す。
    On Error Resume Next
#If EnableCalibration = True Then
    mm2Haba = x / 5 * 13.5 * 1.147028154 * 0.98167218 * Calib.Range("Z57").Value
#Else
    mm2Haba = x / 5 * 13.5 * 1.147028154 * 0.98167218
#End If
End Function

Private Function mm2Takasa(x As Double) As Double
' mm をセル高さに換算する。
    On Error Resume Next
#If EnableCalibration = True Then
    mm2Takasa = x / 5 * 13.5 * 1.147028154 * Calib.Range("Z58").Value
#Else
    mm2Takasa = x / 5 * 13.5 * 1.147028154
#End If
End Function

Private Function Kanzan2mm(x As Double, Unit As Long) As Double
' 大きさと単位から、mm数値を返す。
    On Error Resume Next
    Kanzan2mm = 0
    Select Case Unit
    Case 1      ' mm: =1mm
        Kanzan2mm = x
    Case 2      ' cm: = 10mm
        Kanzan2mm = x * 10
    Case 3      ' inch: = 25.4mm
        Kanzan2mm = x * 25.4
    Case 4      ' 寸: = 30.303mm
        Kanzan2mm = x * 30.303
    End Select
End Function

Private Function setupHochPage(Sa As Worksheet) As Boolean
' 高精度のページ設定を行う。新規設定を行った場合にTrueを返す。
    Dim x As Double
    'On Error Resume Next
    setupHochPage = False
    x = CDbl(100) / HOCHRES
    With Sa
        If .PageSetup.Zoom <> x Then    ' 印刷倍率が一致すれば設定済みとする。
            .PageSetup.Zoom = x
            If Sa Is ActiveWindow.ActiveSheet Then
                .Cells.Font.Size = .Cells(1).Style.Font.Size * HOCHRES  ' 標準サイズにかける
                '.Cells.Font.Size = 11 * HOCHRES       ' 固定の方が安全?
                ActiveWindow.Zoom = x
            End If
            setupHochPage = True
        End If
    End With
End Function

Private Sub updateCalibSheet()
' キャリブレーションシートの更新(常に高精度)
    Dim Sa As Worksheet
    Dim Ra As Range
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    On Error GoTo Err1
    Application.ScreenUpdating = False
    Set Sa = Calib      ' キャリブレーションシート
    ' 設定変更
    x1 = Val(Calib.Range("Z57").Value)
    y1 = Val(Calib.Range("Z58").Value)
    If x1 <= 0 Then x1 = 1
    If y1 <= 0 Then y1 = 1
    x2 = Val(Calib.Range("E1").Value)
    y2 = Val(Calib.Range("A5").Value)
    Calib.Range("Z57").Value = x1 * 10 / x2
    Calib.Range("Z58").Value = y1 * 10 / y2
    Calib.Range("E1").Value = 10
    Calib.Range("A5").Value = 10
    ' 高さと幅の再調整
    setupHochPage Calib
    applyWidth2Range Sa.Cells, mm2Haba(5 * HOCHRES)
    applyHeight2Range Sa.Cells, mm2Takasa(5 * HOCHRES)
    applyHeight2Range Sa.Rows("$27:$36"), mm2Takasa(1 * HOCHRES)
    applyHeight2Range Sa.Rows("$39:$48"), mm2Takasa(25.4 / 10 * HOCHRES)
    applyHeight2Range Sa.Rows("$49:$58"), mm2Takasa(30.303 / 10 * HOCHRES)
    Set Sa = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "調整が終わりました。すぐに印刷して確認することができます。" & vbNewLine & _
        "調整結果を保持するために、ブックを保存してください。", , BAUCAP
    Exit Sub
Err1:
    Application.ScreenUpdating = True
    MsgBox "微調整に失敗しました。" & vbNewLine & Err.Description, , BAUCAP
End Sub

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
'   name;Calib
'{{{
Option Explicit

#Const EnableFunctions = True       ' メンテナンス用

#If EnableFunctions = True Then

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Select Case Target.Address(False, False, xlA1, False)
    Case "L20:O21"
        Calib.Cells(1).Select
        Application.Run "updateCalibSheet"
    End Select
End Sub

#End If

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