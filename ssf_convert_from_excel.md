

# ssf writer #

introduce tools to convert an Excel book to an SSF text

## SSF への保存 ##

エクセルブックを SSF 形式のテキストに変換するツールを紹介する

# ssf writer primitive #

  * a simple tool, comes with only a single VBA module.
  * minimized and essential functions, includes VBA macro, references, cell formulas, cell formats and cell names. put them into a text.
  * [more>>](ssf_writer_primitive.md)

## SSF ライター 初号機 ##
  * １つの VBA 標準モジュールだけで動作する単純なツール。
  * VBA マクロ、参照設定、セル数式、セル書式、範囲名という、最低限必要なものをテキストに変換できる威力は大きい。
  * [もっと詳しく>>](ssf_writer_primitive.md)

# ssf rw primitive #

  * mixed [ssf\_reader\_primitive](ssf_reader_primitive.md) and [ssf\_writer\_primitive](ssf_writer_primitive.md) into a single book.
  * has tool bar as a user interface.
  * [more>>](ssf_rw_primitive.md)

## SSF読み書き 壱号機 ##
  * [ssf\_reader\_primitive](ssf_reader_primitive.md) と [ssf\_writer\_primitive](ssf_writer_primitive.md) を１つのブックにまとめた。
  * ツールバーで操作する。
  * [もっと詳しく>>](ssf_rw_primitive.md)

# ssf rw v2 #

  1. read and write file with a dialog.
  1. support many encodings arround utf-8. while the primitive support only the localized ANSI.
  1. write SSF for the selected part.
  1. write SSF for each module.
  1. handle colors, font and more formats.
  * [more>>](ssf_rw_v2.md)

## SSF読み書き ２号機 ##
  1. ダイアログ付きのファイル入出力。
  1. ANSI 以外の多数の文字コードにも対応。 UTF-8 を標準。
  1. 部分的な SSF 出力。
  1. モジュール単位での SSF 出力。
  1. 色や文字サイズなど、書式の強化。
  * [もっと詳しく>>](ssf_rw_v2.md)