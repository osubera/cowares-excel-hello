

# Introduction #

  * how to get html document from web using http in vba

## 概要 ##
  * VBAで http を使ってウェブから HTML 文書を取得する

# Details #

  * we use the Microsoft HTML Object Library 4.0 to process a request.
  * this library is troublesome and chaotic, because it has many differenct interfaces depend machines while it shows exactly same version 4.0.
  * we use the `MSHTML.IHTMLDocument4` interface here, so XP and later may go well, but 2000 and earliar may not work.
  * this code has a explicit interface as IHTMLDocument4, maybe you can use HTMLDocument instead.
  * the `createDocumentFromUrl` method gets document from the specified url, and parse it into a html dom document.
  * after an HTML Document is gotten, we can use familiar dom functions to analyse the document.
  * when a Charset used to parse is wrong, we can specify a correct one to this property and "Refresh" the document.
  * it goes to an asynchronous request, so we must watch the status, or use event procedures.
  * the first code shows a simple polling.  the second uses an event procedure.

## 説明 ##
  * Microsoft HTML Object Library 4.0 でリクエストを処理する。
  * このライブラリはやっかいでカオスだ。同じバージョン 4.0 なのにマシンによって異なる、多くのインターフェースを持っている。
  * ここでは `MSHTML.IHTMLDocument4` インターフェースを使うので、 XP 以降は良いが 2000 とかもっと古いとだめ。
  * コードに IHTMLDocument4 と明確なインターフェースを指示しているところは、単に HTMLDocument と置き換えてもよい。
  * `createDocumentFromUrl` メソッドは URL から文書を取得して HTML DOM への解析まで行う。
  * HTML 文書が手に入れば、使い慣れた DOM の機能で文書を解析できる。
  * Charset が間違って使われている場合、プロパティに正しいものを指示して "Refresh" すればよい。
  * これは非同期に動作するので、ステータス監視をするか、イベントプロシジャを使わないといけない。
  * 最初のコードは単純なポーリングを使った。２番目のはイベントプロシジャ。

# How to use #

  1. use an ssf reader tool like [ssf\_reader\_primitive](ssf_reader_primitive.md) to convert a text code below into an excel book.
  1. test1() are executable examples, it lists links.

## 使い方 ##
  1. [ssf\_reader\_primitive](ssf_reader_primitive.md) のような ssf 読み込みツールを使って、下のコードをエクセルブックに変換する。
  1. test1() が実行可能な見本。リンクを抽出する。

# Code #

```

'workbook
'  name;htllo_http_get_html.xlsm

'require
'  ;{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B} 4 0 Microsoft HTML Object Library


'module
'  name;Module1
'{{{
Option Explicit

Sub test1()
    Const Timeout As Long = 5
    Const Url As String = "http://www.nic.ad.jp"
    
    Dim tp As MSHTML.IHTMLDocument4
    Dim doc As MSHTML.HTMLDocument
    Dim tags As MSHTML.IHTMLElementCollection
    Dim tag As MSHTML.IHTMLElement
    Dim TimeUp As Date

    Set tp = New MSHTML.HTMLDocument
    Set doc = tp.createDocumentFromUrl(Url, vbNullString)
    
    TimeUp = Now() + Timeout / 24 / 3600
    Do
        DoEvents
        If (doc.readyState = "complete") Then Exit Do
    Loop While (Now() < TimeUp)
    
    If doc.readyState <> "complete" Then
        Debug.Print "timeout"
        GoTo Done
    End If
    
    If LenB(doc.DocumentElement.outerHTML) = 0 Then
        Debug.Print "blank document"
        GoTo Done
    End If
    
    Debug.Print doc.Charset
    'doc.Charset = "iso-2022-jp"
    'doc.execCommand "Refresh"
    
    Set tags = doc.getElementsByTagName("H1")
    For Each tag In tags
        Debug.Print tag.innerText
        Debug.Print tag.outerHTML
    Next
    Set tags = doc.body.getElementsByTagName("A")
    For Each tag In tags
        Debug.Print tag.innerText, tag.getAttribute("href")
    Next
    
    
Done:
        doc.Close
        tp.Close
        Set doc = Nothing
        Set tp = Nothing
End Sub
'}}}

```

### Result ###

test1()

```
iso-2022-jp
社団法人日本ネットワークインフォメーションセンター

<H1><A href="/ja/"><SPAN class=hide>社団法人日本ネットワークインフォメーションセンター</SPAN></A></H1>
メインコンテンツへスキップ  http://www.nic.ad.jp/#primaryContent
TOP           http://www.nic.ad.jp/ja/
ENGLISH       http://www.nic.ad.jp/en/
SITEMAP       http://www.nic.ad.jp/ja/sitemap.html
RSS           http://www.nic.ad.jp/ja/index.xml
社団法人日本ネットワークインフォメーションセンター      http://www.nic.ad.jp/ja/
WHOISとは     http://www.nic.ad.jp/ja/whois/index.html
JPNIC WHOIS Gateway         http://www.nic.ad.jp/ja/whois/ja-gateway.html
Q&A           http://www.nic.ad.jp/ja/question/ip5.html
              http://www.nic.ad.jp/ja/ip/application.html
              http://www.nic.ad.jp/ja/profile/beginner.html
              http://www.nic.ad.jp/ja/member/guide/
              http://www.nic.ad.jp/member/
              http://www.nic.ad.jp/#notice1
              http://www.nic.ad.jp/#event1
              http://www.nic.ad.jp/ja/member/list/index.html
              http://www.nic.ad.jp/ja/member/list/6.html
              http://www.nic.ad.jp/ja/member/list/131.html
              http://www.nic.ad.jp/ja/member/list/494.html
              http://www.venus.gr.jp/opf-jp/events/showcase4/
              http://www.nic.ad.jp/ja/mailmagazine/
              http://www.nic.ad.jp/ja/ip/ipv4pool/
JPNICとは     http://www.nic.ad.jp/ja/profile/index.html
活動理念      http://www.nic.ad.jp/ja/profile/view.html
組織概要      http://www.nic.ad.jp/ja/profile/about.html
定款・細則    http://www.nic.ad.jp/ja/profile/rule.html
情報公開      http://www.nic.ad.jp/ja/profile/disclose/index.html
プレスリリース              http://www.nic.ad.jp/ja/pressrelease/index.html
会員リスト    http://www.nic.ad.jp/ja/member/list/index.html
入会案内      http://www.nic.ad.jp/ja/member/guide/index.html
関連団体へのリンク          http://www.nic.ad.jp/ja/profile/link.html
Q&A           http://www.nic.ad.jp/ja/question/nic.html
IPアドレス    http://www.nic.ad.jp/ja/ip/index.html
IPアドレス管理の基礎知識    http://www.nic.ad.jp/ja/ip/admin-basic.html
IPアドレス・AS番号が欲しい時は            http://www.nic.ad.jp/ja/ip/whereto/
IPアドレス登録管理業務について            http://www.nic.ad.jp/ja/ip/about-regist-admin.html
ドキュメント一覧            http://www.nic.ad.jp/ja/ip/doc/
IPアドレス関連のミーティング              http://www.nic.ad.jp/ja/ip/event/
IPアドレストピックス        http://www.nic.ad.jp/ja/ip/topics/
統計・各種リスト            http://www.nic.ad.jp/ja/stat/ip/
JPIRR         http://www.nic.ad.jp/ja/irr/
逆引きネームサーバの適切な設定について    http://www.nic.ad.jp/ja/dns/lame/
Q&A           http://www.nic.ad.jp/ja/question/ip.html
インターネットガバナンス    http://www.nic.ad.jp/ja/governance/index.html
歴史と背景    http://www.nic.ad.jp/ja/newsletter/No26/020.html
ICANN情報     http://www.nic.ad.jp/ja/icann/index.html
JPNICからの発信             http://www.nic.ad.jp/ja/opinion/index.html
国際関係組織一覧            http://www.nic.ad.jp/ja/intl/org/org.html
インターネットの基礎        http://www.nic.ad.jp/ja/basics/index.html
インターネットのしくみ      http://www.nic.ad.jp/ja/basics/beginners/index.html
用語集        http://www.nic.ad.jp/ja/tech/glossary.html
1分解説       http://www.nic.ad.jp/ja/basics/terms/index.html
10分講座      http://www.nic.ad.jp/ja/newsletter/10minute.html
歴史の一幕    http://www.nic.ad.jp/ja/newsletter/history.html
インターネットの技術        http://www.nic.ad.jp/ja/tech/index.html
JPIRR         http://www.nic.ad.jp/ja/irr/index.html
VoIP/SIP相互接続検証タスクフォース        http://www.nic.ad.jp/ja/voip-sip-tf/index.html
ENUM          http://www.nic.ad.jp/ja/enum/index.html
DNS運用健全化タスクフォース http://www.nic.ad.jp/ja/dnsqc/index.html
IETFとRFC     http://www.nic.ad.jp/ja/tech/rfc-jp.html
ドメイン名の国際化          http://www.nic.ad.jp/ja/tech/idn.html
セキュリティ関連情報        http://www.nic.ad.jp/ja/security/
リンク集      http://www.nic.ad.jp/ja/tech/links.html
JPNIC認証局   http://www.nic.ad.jp/ja/research/ca/
JPNIC認証局とは             http://www.nic.ad.jp/ja/research/ca/about-jpnic-ca.html
資源一括登録システム        http://www.nic.ad.jp/ja/research/ca/about-web-transaction.html
経路情報の登録認可機構      http://www.nic.ad.jp/ja/research/ca/routereg-outline/
Q&A           http://www.nic.ad.jp/ja/question/ca.html
fingerprintのページ         https://serv.nic.ad.jp/capub/fingerprint.html
ドメイン名    http://www.nic.ad.jp/ja/dom/
ドメイン名とは              http://www.nic.ad.jp/ja/dom/basics.html
国際化ドメイン名            http://www.nic.ad.jp/ja/dom/idn.html
JPNICが行っているドメイン名関連事業紹介   http://www.nic.ad.jp/ja/dom/intro.html
ドメイン名トピックス        http://www.nic.ad.jp/ja/dom/topics.html
gTLD          http://www.nic.ad.jp/ja/dom/gtld.html
ドメイン名紛争処理方針(DRP) http://www.nic.ad.jp/ja/drp/index.html
データエスクロー            http://www.nic.ad.jp/ja/dom/escrow/
統計          http://www.nic.ad.jp/ja/stat/dom/index.html
イベント      http://www.nic.ad.jp/ja/dom/mtg.html
Q&A           http://www.nic.ad.jp/ja/question/domain.html
意見募集      http://www.nic.ad.jp/ja/dom/opinion/
JPNICライブラリ             http://www.nic.ad.jp/ja/library.html
ドキュメントライブラリ      http://www.nic.ad.jp/ja/doc/index.html
イベント・講演会資料        http://www.nic.ad.jp/ja/materials/index.html
ニュースレター              http://www.nic.ad.jp/ja/newsletter/index.html
メールマガジン              http://www.nic.ad.jp/ja/mailmagazine/index.html
執筆記事      http://www.nic.ad.jp/ja/pub/index.html
調査報告書    http://www.nic.ad.jp/ja/research/index.html
会議資料      http://www.nic.ad.jp/ja/profile/mtg/index.html
JPNICからの提言             http://www.nic.ad.jp/ja/opinion/
インターネットの歴史・統計  http://www.nic.ad.jp/ja/history.html
JPNICの歩み   http://www.nic.ad.jp/ja/history/index.html
歴史の一幕    http://www.nic.ad.jp/ja/newsletter/history.html
JPドメイン名の歩み          http://www.nic.ad.jp/ja/dom/jpdom.html
統計          http://www.nic.ad.jp/ja/stat/index.html
各種メーリングリストへの参加、退会        http://www.nic.ad.jp/ja/profile/ml.html
第29回ICANN報告会開催のご案内             http://www.nic.ad.jp/ja/topics/2011/20110106-01.html
IPレジストリシステムおよびJPIRRシステム1月の定期メンテナンスに伴うサービス一時停止のお知らせ      http://www.nic.ad.jp/ja/topics/2010/20101228-01.html
IPv4アドレス在庫枯渇予測に関するAPNIC理事会の声明について             http://www.nic.ad.jp/ja/topics/2010/20101215-02.html
ICANNトピックス：ICANN理事会(2010年12月10日開催)決議概要              http://www.nic.ad.jp/ja/topics/2010/20101215-01.html
IPアドレス等料金体系改定の見送りについて  http://www.nic.ad.jp/ja/topics/2010/20101213-01.html
第19回JPNICオープンポリシーミーティング参加登録締め切り間近のお知らせ http://www.nic.ad.jp/ja/topics/2010/20101210-01.html
トピックス一覧へ            http://www.nic.ad.jp/ja/topics/index.html
ICANNアナウンスメント一覧   http://www.nic.ad.jp/ja/icann/announcements.html
第42回総会    http://www.nic.ad.jp/ja/materials/general-meeting/20101210/
JPNICが管理するIPアドレス・AS番号・IRRサービスに関する統計            http://www.nic.ad.jp/ja/stat/ip/20101208.html
アジア太平洋地域の国別IPv4アドレス、IPv6アドレス、AS番号配分状況      http://www.nic.ad.jp/ja/stat/ip/asia-pacific.html
地域インターネットレジストリ(RIR)ごとのIPv4アドレス、IPv6アドレス、AS番号配分状況   http://www.nic.ad.jp/ja/stat/ip/world.html
Web更新情報一覧へ           http://www.nic.ad.jp/ja/changelog/index.html
vol.811       http://www.nic.ad.jp/ja/mailmagazine/backnumber/2011/vol811.html
vol.810       http://www.nic.ad.jp/ja/mailmagazine/backnumber/2010/vol810.html
メールマガジン一覧へ        http://www.nic.ad.jp/ja/mailmagazine/backnumber/index.html
重要なお知らせ              
IPアドレス等料金体系改定の見送りについて  http://www.nic.ad.jp/ja/topics/2010/20101213-01.html
IPv6アドレスの分配方法の簡略化について    http://www.nic.ad.jp/ja/topics/2010/20100726-02.html
whois.jpサービスのJPNIC／JPRS共同運営およびJPNIC WHOIS転送の終了について            http://www.nic.ad.jp/ja/topics/2010/20100701-01.html
イベント情報  
JPNICオープンポリシミーティングショーケース4            http://www.venus.gr.jp/opf-jp/events/showcase4/
JANOG27       http://www.janog.gr.jp/meeting/janog27/
第29回ICANN報告会           http://www.nic.ad.jp/ja/topics/2011/20110106-01.html
HOSTING PRO 2011            http://hosting-pro.jp/
イベントカレンダーへ        http://www.nic.ad.jp/ja/event/calendar.html
              http://etjp.jp/
              http://icann.nic.ad.jp/
              http://igtf.jp/
              http://www.kokatsu.jp/
              http://www.v6pc.jp/
              http://www.jdna.jp/
              http://thinkquest.jp/
              http://www.nic.ad.jp/ja/voip-sip-tf/index.html
お問い合わせ先              http://www.nic.ad.jp/ja/profile/info.html
著作権／リンク              http://www.nic.ad.jp/ja/copyright.html
JPNIC個人情報保護方針       http://www.nic.ad.jp/ja/privacy.html
Q&A           http://www.nic.ad.jp/ja/question/index.html
 IPv6 Enabled http://www.ipv6forum.com/ipv6_enabled/approval_list.php
```


# More Code #

  * using event procedures.
> > イベントプロシジャを使ってみる。
  * to enable user customized procedure, we choose a `OnReadyStateChange` as a fixed name sub on standard module, called from the event procedure.
> > 処理内容をユーザーが調整しやすいように、`OnReadyStateChange` を名前固定の、標準モジュールにあるコールバック関数とした。これがイベントプロシジャから呼ばれる。

```

'workbook
'  name;htllo_http_get_html.xlsm

'require
'  ;{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B} 4 0 Microsoft HTML Object Library



'module
'  name;Module1
'{{{
Option Explicit

Dim Que As Collection

Sub test1()
    Const Url As String = "http://www.kantei.go.jp"
    
    Dim doc As HelloHttpGetHtml
    
    Set Que = New Collection
    Set doc = New HelloHttpGetHtml
    doc.Key = Now()
    Que.Add doc, doc.Key
    
    doc.send Url
    Debug.Print doc.readyState
End Sub

' コールバック
Public Sub OnReadyStateChange(doc As HelloHttpGetHtml)
    Dim tags As MSHTML.IHTMLElementCollection
    Dim tag As MSHTML.IHTMLElement
    Dim i As Long
    
    Debug.Print doc.readyState
    If Not doc.IsComplete Then Exit Sub
    If Not doc.IsNotBlank Then Exit Sub
    
    Debug.Print doc.Charset
    
    Set tags = doc.Document.getElementsByTagName("H1")
    For Each tag In tags
        Debug.Print tag.innerText
        Debug.Print tag.outerHTML
    Next
    Set tags = doc.Document.body.getElementsByTagName("A")
    For Each tag In tags
        Debug.Print tag.innerText, tag.getAttribute("href")
    Next
    
    For i = Que.Count To 1 Step -1
        If Que(i).Key = doc.Key Then
            Que.Remove i
            Set doc = Nothing
            Exit For
        End If
    Next
End Sub

'}}}

'class
'  name;HelloHttpGetHtml
'{{{
Option Explicit

Private tp As MSHTML.IHTMLDocument4
Private WithEvents doc As MSHTML.HTMLDocument
Private MyKey As String

Private Sub Class_Initialize()
    Set tp = New MSHTML.HTMLDocument
End Sub

Private Sub Class_Terminate()
    If doc.readyState = "loading" Or doc.readyState = "interactive" Then doc.execCommand "Stop"
    Set doc = Nothing
    Set tp = Nothing
End Sub

Public Property Get Key() As String
    Key = MyKey
End Property

Public Property Let Key(NewKey As String)
    MyKey = NewKey
End Property

Public Property Get readyState() As String
    readyState = doc.readyState
End Property

Public Property Get IsComplete() As Boolean
    IsComplete = (doc.readyState = "complete")
End Property

Public Property Get IsNotBlank() As Boolean
    IsNotBlank = (LenB(doc.DocumentElement.outerHTML) > 0)
End Property

Public Property Get Document() As MSHTML.HTMLDocument
    Set Document = doc
End Property

Public Property Get Charset() As String
    Charset = doc.Charset
End Property

Public Property Let Charset(NewCharset As String)
    doc.Charset = NewCharset
    doc.execCommand "Refresh"
End Property

Public Sub send(Url As String)
    Set doc = tp.createDocumentFromUrl(Url, vbNullString)
End Sub

Private Sub doc_onreadystatechange()
    OnReadyStateChange Me
End Sub
'}}}

```

### Result ###

```
loading
interactive
complete
utf-8


<H1><A href="/"><IMG alt="首相官邸 Prime Minister of Japan and His Cabinet" src="/jp/n-common/images/logo.gif" width=406 height=42></A></H1>
              http://www.kantei.go.jp/
RSS配信について             http://www.kantei.go.jp/rss.html
サイトマップ  http://www.kantei.go.jp/sitemap.html
English       http://www.kantei.go.jp/foreign/index-e.html
              http://www.kantei.go.jp/jp/singi/index.html
              http://www.kantei.go.jp/jp/kan/meibo/index.html
              http://www.kantei.go.jp/jp/rekidainaikaku.html
              http://www.kantei.go.jp/jp/iken.html
              http://www.kantei.go.jp/jp/link/server_j.html
RSS           http://www.kantei.go.jp/index-j2.rdf
              http://www.kantei.go.jp/jp/kan/actions/index.html
              http://www.kantei.go.jp/jp/kan/statement/
菅内閣総理大臣年頭記者会見  http://www.kantei.go.jp/jp/kan/statement/201101/04nentou.html
菅内閣総理大臣　平成二十三年　年頭所感    http://www.kantei.go.jp/jp/kan/statement/201101/01nentou.html
硫黄島戦没者追悼式　追悼の辞              http://www.kantei.go.jp/jp/kan/statement/201012/14tuitounoji.html
ノーベル賞授賞式に当たって 総理メッセージ               http://www.kantei.go.jp/jp/kan/statement/201012/11message.html
日本・ボリビア共同声明      http://www.kantei.go.jp/jp/kan/statement/201012/08nichibolivia.html
菅内閣総理大臣記者会見      http://www.kantei.go.jp/jp/kan/statement/201012/06kaiken.html
日本・バングラデシュ共同声明　国際社会と南アジアの平和と繁栄にむけての強固なパートナーシップの拡大              http://www.kantei.go.jp/jp/kan/statement/201011/29nichibangladesh.html
拉致問題の解決に向けて      http://www.kantei.go.jp/jp/kan/statement/201011/29siji.html
議会開設百二十年記念式典における内閣総理大臣祝辞        http://www.kantei.go.jp/jp/kan/statement/201011/29syukuji.html
第５回 新成長戦略実現会議 菅総理指示      http://www.kantei.go.jp/jp/kan/statement/201011/25siji.html
「戦略的パートナーシップ」構築に向けた日本・モンゴル共同声明          http://www.kantei.go.jp/jp/kan/statement/201011/19nichimongolia.html
              http://www.kantei.go.jp/jp/tyoukanpress/index.html
RSS           http://www.kantei.go.jp/tyoukan.rdf
テキスト版    http://www.kantei.go.jp/jp/tyoukanpress/201101/5_a.html
動画版        http://nettv.gov-online.go.jp/prg/prg4232.html
これまでの発表一覧          http://www.kantei.go.jp/jp/tyoukanpress/index.html
内閣官房長官談話            http://www.kantei.go.jp//jp/tyokan/kan/2010/1217danwa.html
談話等一覧    http://www.kantei.go.jp/jp/tyokan/kan/index.html
              http://www.kantei.go.jp/jp/kakugikettei/index.html
アクション・プラン ～出先機関の原則廃止に向けて～       http://www.kantei.go.jp/jp/kakugikettei/2010/1228action_plan.pdf
平成23年度の経済見通しと経済財政運営の基本的態度　～新成長戦略実現に向けたステップ３ へ～（閣議了解）           http://www.kantei.go.jp/jp/kakugikettei/2010/1222mitoshi.pdf
男女共同参画基本計画の変更について（閣議決定）          http://www.kantei.go.jp/jp/kakugikettei/2010/1217dai3danjo_kihonkeikkaku.pdf
平成22年12月28日 定例閣議案件             http://www.kantei.go.jp/jp/kakugi/2010/kakugi-2010122801.html
閣議案件      http://www.kantei.go.jp/jp/kakugi/index.html
              http://www.kantei.go.jp/jp/yosan23/
平成23年度の経済見通しと経済財政運営の基本的態度　～新成長戦略実現に向けたステップ３ へ～（閣議了解）[PDF]      http://www.kantei.go.jp/jp/kakugikettei/2010/1222mitoshi.pdf
平成２３年度以降に係る防衛計画の大綱（閣議決定）[PDF]   http://www.kantei.go.jp/jp/kakugikettei/2010/1217boueitaikou.pdf
中期防衛力整備計画（平成２３年度～平成２７年度）（閣議決定）[PDF]     http://www.kantei.go.jp/jp/kakugikettei/2010/1217tyuukiboueiryokukeikaku.pdf
平成２３年度予算編成の基本方針（閣議決定）[PDF]         http://www.kantei.go.jp/jp/kakugikettei/2010/h23yosan_kihonhoushin.pdf
平成２３年度税制改正大綱（閣議決定）[PDF] http://www.kantei.go.jp/jp/kakugikettei/2010/h23zeiseitaikou.pdf
「雇用戦略対話」合意 ～『雇用戦略・基本方針２０１１』について～[PDF]  http://www.kantei.go.jp/jp/kakugikettei/2010/101215goui.pdf
社会保障改革の推進について（閣議決定）[PDF]             http://www.kantei.go.jp/jp/kakugikettei/2010/1214suishin_syakaihosyou.pdf
日本ＡＰＥＣ首脳会談の概要と成果          http://www.kantei.go.jp/jp/apecjapan2010/
基本方針[PDF] http://www.kantei.go.jp/jp/kakugikettei/2010/0917kihonhousin.pdf
              http://www.npu.go.jp/
              http://www.cao.go.jp/gyouseisasshin/index.html
              http://www.challenge25.go.jp/
              http://www.cao.go.jp/sasshin/kokumin_koe/uketsuke.html
新成長戦略実現に向けた３段構えの経済対策  http://www.kantei.go.jp/jp/keizaitaisaku2010/
新成長戦略～「元気な日本」復活のシナリオ  http://www.kantei.go.jp/jp/sinseichousenryaku/
元気な日本復活特別枠に関する評価会議      http://www.kantei.go.jp/jp/singi/genki/article/
新成長戦略実現会議          http://www.npu.go.jp/policy/policy04/archive02.html
「新しい公共」              http://www5.cao.go.jp/npc/index.html
拉致問題対策本部            http://www.rachi.go.jp/
硫黄島からの遺骨帰還のための特命チーム    http://www.kantei.go.jp/jp/singi/ioutou/
自律的労使関係制度の措置に向けての意見募集(国家公務員制度改革推進本部事務局)        http://www.gyoukaku.go.jp/koumuin/index.html
児童虐待、いじめ、ひきこもり、不登校等についての相談・通報窓口（内閣府）            http://www8.cao.go.jp/youth/soudan/index.html
住宅エコポイントの概要について（国土交通省）            http://www.mlit.go.jp/jutakukentiku/house/jutakukentiku_house_tk4_000017.html
              http://kanfullblog.kantei.go.jp/
              http://kanfullblog.kantei.go.jp/2010/12/20101228.html
来年度政府原案決定　～総理が明かす三つの攻防            http://kanfullblog.kantei.go.jp/2010/12/20101228.html
先を見すえて－反転攻勢の年が明けた        http://kanfullblog.kantei.go.jp/2011/01/20110105.html
着実に進むインフラの海外展開－　日本の技術・経験で海外のビジネスチャンスをつかむ　－              http://kanfullblog.kantei.go.jp/2011/01/20110104.html
先を見すえて－100歳の教え―――もう一つの年頭所感       http://kanfullblog.kantei.go.jp/2011/01/20110101.html
              http://www.mmz.kantei.go.jp/jp/blog/kan/index.html
              http://nettv.gov-online.go.jp/prg/prg4228.html
平成22年12月20日～12月26日　ＨＴＬＶ－１特命チーム、日本プロスポーツ大賞総理大臣賞状授与式ほか    http://nettv.gov-online.go.jp/prg/prg4228.html
              http://nettv.gov-online.go.jp/prg/prg4231.html
菅内閣総理大臣年頭記者会見-平成23年1月4日 http://nettv.gov-online.go.jp/prg/prg4231.html
地球温暖化問題に関する閣僚委員会-平成22年12月28日       http://nettv.gov-online.go.jp/prg/prg4229.html
野依科学技術・学術審議会会長等による表敬-平成22年12月27日             http://nettv.gov-online.go.jp/prg/prg4225.html
地域主権戦略会議-平成22年12月27日         http://nettv.gov-online.go.jp/prg/prg4223.html
              http://www.kantei.go.jp/jp/kan/yotei/
              http://nettv.gov-online.go.jp/
徳光＆木佐の知りたいニッポン！～くらしを彩る新しい『伝統的工芸品』    http://nettv.gov-online.go.jp/prg/prg4203.html
              http://nettv.gov-online.go.jp/prg/prg4203.html
家電エコポイント制度見直しのお知らせ      http://nettv.gov-online.go.jp/prg/prg3900.html
              http://nettv.gov-online.go.jp/prg/prg3900.html
              http://www.gov-online.go.jp/
消費者庁では、メール配信サービスにより、子どもの思わぬ事故を防ぐための注意点や豆知識を毎週お届けしています。    http://www.gov-online.go.jp/closeup/20101227.html
              http://www.gov-online.go.jp/closeup/20101227.html
リンク、著作権等について    http://www.kantei.go.jp/jp/terms.html
プライバシーポリシー        http://www.kantei.go.jp/jp/policy/privacy_policy.html
リンク集      http://www.kantei.go.jp/jp/link/server_j.html
資料集        http://www.kantei.go.jp/jp/siryou.html
官報・白書    http://www.kantei.go.jp/jp/kanpo_hakusyo.html
```