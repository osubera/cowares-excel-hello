

# Definitions of SSF #

we only define a framework here, because the ssf depends office applications and operating systems.

### ssf ###

  1. a ssf is a group of ssf blocks.
  1. a void block is found between any ssf blocks.

### line ###

  1. a parser do a line at once.
  1. escaped lines are handled as a single line on parsing.
  1. escaped lines override every rule below.

### ssf block ###

  1. a ssf line is defined explicitly and independently. maybe by the first character of a line.
  1. a ssf block is a continuous ssf lines
  1. a ssf block has a pair of a ssf key and a ssf value.
  1. the first line of a ssf block contains a ssf key, the second line and later contain a ssf value.
  1. a ssf key is for telling how to parse it, not for identifing unique block.

### void block ###

  1. a void line is every line but ssf lines.
  1. a void block is a continuous void lines.

## SSF の定義 ##

ssf は、表計算アプリやOSに依存して変化してよいものなので、定義は枠組みの決定に限定する。

### ssf ###

  1. ssf ブロックの集まりを ssf という。
  1. ssf ブロックの間には、必ず無効ブロックが入る。

### 行 ###

  1. 解釈は行単位で行う。
  1. 複数行を解釈上の単一行とみなすための、行エスケープ機構を用意する。
  1. 行エスケープはすべてのルールの中で最優先される。

### ssf ブロック ###

  1. 行の先頭文字など、該当行だけで明確な定義により、 ssf 有効行を決める。
  1. ssf 有効行の連続した固まりを ssf ブロックと呼ぶ。
  1. ssf ブロックは ssf キーと ssf 値の一対のデータを表す。
  1. ssf ブロック内の先頭行が ssf キーを持ち、２行目以降が ssf 値を持つ。
  1. ssf キーは、解釈ルールを示すものであり、重複を妨げるキーではない。

### 無効ブロック ###

  1. ssf 有効行でないものを無効行と呼ぶ。
  1. １つ以上の連続する無効行の固まりを無効ブロックと呼ぶ。
