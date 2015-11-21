

# how to open the Visual Basic Editor window #

### Excel 2003 and earlier ###

  * select followings at the menu bar
    1. Tools
    1. Macro
    1. Visual Basic Editor
  * you may want to disable "Menus show recently used commands first" settings.
    1. Tools
    1. Customize
    1. Options tab
    1. uncheck "Menus show recently used commands first"
    1. click Close the Customize window

### Excel 2007 and later ###

  * follow this at first, because the Visual Basic is hidden at default
    1. click the Microsoft Office Button located at the top-left corner.
    1. Excel Options at the bottom
    1. Popular from the left menu
    1. check "Show Developer tab in the Ribbon"
    1. click OK button to close the Excel Options window
  * do followings at the Ribbon
    1. Developer tab
    1. Visual Basic icon (at the left, in Code)

### Excel 2010 ###

  * may change the Microsoft Office Button, but I'm not sure
    * maybe, File, Options, Cusomize Ribbon, or something


## Visual Basic Editor ウィンドウの開き方 ##

### Excel 2003 以前 ###

  * メニューバーで次の操作をする
    1. ツール(T)
    1. マクロ(M)
    1. Visual Basic Editor(V)
  * 「最近使用したコマンドを最初に表示する(N)」機能が邪魔な場合
    1. ツール(T)
    1. ユーザー設定(C)
    1. オプション　タブ
    1. 「最近使用したコマンドを最初に表示する(N)」のチェックを消す
    1. 「閉じる」をクリックして、ユーザー設定ウィンドウを閉じる

### Excel 2007 以降 ###

  * インストール直後はVisual Basic を非表示にしているため、表示する設定が必要
    1. 左上角の Officeボタンをクリック
    1. Excelのオプション(I) （最下行）
    1. 左側一覧の基本設定
    1. 「開発」タブをリボンに表示する(D) をチェック
    1. OK ボタンを押して、Excelのオプションウィンドウを閉じる
  * リボンで次の操作をする
    1. 開発　タブを選ぶ
    1. Visual Basic ボタンを押す （コードグループ、左端）


# how to insert a Visual Basic Module #

  * at the Visual Basic window
    * select followings at the menu bar
      1. Insert
      1. Module

## 標準モジュールを作成するには ##

  * Visual Basic ウィンドウで、次の操作をする
    * メニューから
      1. 挿入(I)
      1. 標準モジュール(M)

# how to allow access to a Visual Basic project #

  * a VBA code raises a run time error 1004 on after Excel 2003 of the default settings, when it uses the Microsoft Visual Basic for Application Extensibility 5.3 Library.
  * most [SSF](ssf.md) tools require the settings bellow.
  * references: the Microsoft KB [813969](http://support.microsoft.com/kb/813969/en-us) and [282830](http://support.microsoft.com/kb/282830/en-us)

### Excel 2003 ###

  * select followings at the menu bar
    1. Tools
    1. Macro
    1. Security
  * in the Security dialog box
    1. select Trusted Sources tab
    1. check the Trust access to Visual Basic Project
    1. click OK to close the dialog box

### Excel 2007 and later ###

  * do the followings
    1. click the Microsoft Office Button located at the top-left corner.
    1. Excel Options at the bottom
    1. click the Trust Center at the left
    1. click the Trust Center Settings button at the right
    1. click the Macro Settings at the left
    1. check the Trust access to the VBA project object model
    1. click OK to close the Trust Center window
    1. click OK to close the Excel Option window
  * check the following link to see the details.
    * [Trust access to the VBA project object model](http://translate.google.co.jp/translate?hl=en&sl=ja&tl=en&u=http%3A%2F%2Fxlsm.web.fc2.com%2Foffice2007%2Fvbproject.html)  (translated by Google from Japanese)

## Visual Basic プロジェクトへのアクセスを信頼させるには ##

  * Excel 2003 以降では、標準のままの設定だと Microsoft Visual Basic for Application Extensibility 5.3 Library を使用するコードで、実行時エラー 1004 が出る。
  * [SSF](ssf.md) 関連のツールは、この設定をしないと動かないものが多い。
  * 参考: マイクロソフトのナリッジベース [813969](http://support.microsoft.com/kb/813969/ja) と [282830](http://support.microsoft.com/kb/282830/ja)

### Excel 2003 ###

  * メニューバーで次の操作をする
    1. ツール(T)
    1. マクロ(M)
    1. セキュリティ(S)
  * セキュリティ ダイアログボックスで次の操作をする
    1. 信頼できる発行元 タブを選ぶ
    1. Visual Basic プロジェクトへのアクセスを信頼する をチェックする
    1. OK をクリックして閉じる

### Excel 2007 以降 ###

  * 次の操作をする
    1. 左上角の Officeボタンをクリック
    1. Excelのオプション(I) （最下行）
    1. 左側一覧のセキュリティセンター
    1. 右側の セキュリティセンターの設定 ボタンをクリック
    1. 左側一覧のマクロの設定
    1. VBA プロジェクト オブジェクト モデルへのアクセスを信頼する(V) をチェック
    1. OK ボタンを押して、セキュリティセンターウィンドウを閉じる
    1. OK ボタンを押して、Excelのオプションウィンドウを閉じる
  * より詳しい説明はこちら
    * [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する](http://xlsm.web.fc2.com/office2007/vbproject.html)