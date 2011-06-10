//' ssf_to_html
//' convert ssf to html table
//' Copyright (C) 2011 Tomizono - kobobau.mocvba.com
//' Fortitudinous, Free, Fair, http://cowares.nobody.jp

//' usage> CScript //Nologo ssf_to_html.js /t Title /e:Charset FILE
// arguments do not work because of a bug of WScript.Arguments, maybe

var Env;
var CellStream;
var LastBlockName;
var BlockName;
var BlockVoid;
var Bch = "'";
var Ech = '\r\n';
var EchDefault = Ech;
var EscapeBegin = '{{{';
var EscapeEnd = '}}}';
var MagicBegin = 'ssf-begin';
var MagicEnd = 'ssf-end';
var Delimiter = ';';

function Main(File, Charset, Caption) {
  if (Charset == undefined) { Charset = 'utf-8'; }
  if (Caption == undefined) { Caption = ''; }
  if (File == undefined) { File = ''; }
  WScript.Echo(FileReader(File, Charset, Caption));
}

function FileReader(FileName, Charset, Caption) {
  var adTypeText = 2;
  
  if (FileName == '') { return ''; }
  InitializeEnv();
  
  var Stream = new ActiveXObject('ADODB.Stream');
  Stream.Open();
  Stream.Type = adTypeText;
  Stream.Charset = Charset;
  Stream.LoadFromFile(FileName);
  ReadSsf(Stream);
  Stream.Close();
  delete Stream;
  
  if (Caption != '') { CellStream.Caption = Caption; }
  var HtmlText = CellStream.ToHtml();
  TerminateEnv();
  return HtmlText;
}



function CellsReadFrom(BlockName, Block) {
  CellStream.SetBlockName(BlockName);
  CellStream.ReadBlock(Block);
}

function ReadFrom(BlockName, Block) {
  switch (BlockName) {
  case 'cells-text':
  case 'cells-formula':
  case 'cells-color':
  case 'cells-background-color':
    CellsReadFrom(BlockName, Block);
    break;
  }
}

function InitializeEnv() {
  Env = new Array();
  CellStream = new HtmlCellStream;
}

function TerminateEnv() {
  CellStream.Terminate();
  delete CellStream;
  delete Env;
}

function ClearEnv() {
  delete Env;
  Env = new Array();
}

function ReadSsf(Stream) {
  var Finder = new StreamParser;
  Finder.Stream = Stream;
  ParseSsf(Finder);
  Finder.Terminate();
  delete Finder;
}

function ParseSsf(Finder) {
  do {
    if (!BeginTheMagic(Finder)) { return; }
    EvalAfter(Finder);
    EndTheMagic();
  } while (RequireTheMagic() && !Finder.AtEndOfStream());
}

function RequireTheMagic() {
  return false;
}

function DoFlush() {
  if (LastBlockName != '') {
    ReadFrom(LastBlockName, Env);
    ClearEnv();
  }
}

function EndTheMagic() {
  DoFlush();
}

function BeginTheMagic(Finder) {
  if (!RequireTheMagic()) { return true; }
  
  var At = Finder.FindString(1, MagicBegin);
  if (At <= 1) { return false; }
  
  Finder.Text = Finder.Text.substring(At - 2);
  
  Bch = Finder.Text.charAt(0);
  var EndBegins = MagicBegin.length + 2;
  var i = Finder.FindString(EndBegins, Bch);
  if (i <= EndBegins) { return false; }
  
  Ech = Finder.Text.substring(EndBegins - 1, i - 2);
  EchDefault = Ech;
  
  return true;
}

function SetSpecialChars() {
  EscapeOff();
}

function EscapeOn() {
  Ech = EchDefault + Bch + '}}}' + EchDefault;
}

function EscapeOff() {
  Ech = EchDefault;
}

function EvalAfter(Finder) {
  SetSpecialChars();
  
  BlockVoid = false;
  BlockName = '';
  var EchNextBlock = Ech + Bch;
  
  while (!Finder.AtEndOfStream()) {
    BeforeTag = EvalComma(Finder, Ech);
    BeforeTag = EvalEscape(BeforeTag, Finder);
    EvalBefore(BeforeTag);
    
    if (BlockName == '') {
      Finder.Text = Ech + Finder.Text;
      BeforeTag = EvalComma(Finder, EchNextBlock);
      if (Finder.Text != '') { Finder.Text = Bch + Finder.Text; }
    } else if (BlockName == MagicEnd) {
      break;
    }
  }
}

function EvalComma(Finder, ech) {
  var At = Finder.FindString(1, ech);
  if (At == 0) {
    BeforeTag = Finder.Text;
    Finder.Text = ''
  } else {
    BeforeTag = Finder.Text.substr(0, At - 1);
    Finder.Text = Finder.Text.substring(At - 1 + ech.length);
  }
  return BeforeTag;
}

function EvalEscape(BeforeTag, Finder) {
  if (BeforeTag != Bch + EscapeBegin) { return BeforeTag; }
  
  EscapeOn();
  BeforeTag = Bch + Delimiter + EvalComma(Finder,Ech);
  EscapeOff();
  return BeforeTag;
}

function EvalBefore(BeforeTag) {
  var Key; var Value;
  
  if (BeforeTag == '') {
    BlockName = '';
  } else if (BeforeTag.charAt(0) != Bch) {
    BlockName = '';
  } else if (BlockName == '') {
    DoFlush();
    BlockName = BeforeTag.substring(1);
    Key = '';
    Value = '';
    BlockVoid = false;
    LastBlockName = BlockName;
  } else if (!BlockVoid) {
    var At = BeforeTag.indexOf(Delimiter);
    if (At == -1) {
      Key = BeforeTag;
      Value = '';
    } else {
      Key = BeforeTag.substr(0, At);
      Value = BeforeTag.substring(At + 1);
    }
    Key = Key.substring(1).replace('\t', '').replace(/^ +/, '').replace(/ +$/,'');
    
    if (Key == 'void') {
      BlockVoid = true;
    } else {
      Env.push(new Array(Key, Value));
    }
  }
}



// class AddressUtilsHelper
//' Excel Columns are 26 decimal, but each digit begins at 1
//' our inside Key is [Row Number],[Col Number]

function AddressUtilsHelper() {
}

AddressUtilsHelper.prototype = {
  //' map 1..26 into A..Z
  N2A: function(Number) {
    var Achar = 65;    //' A
    return String.fromCharCode(Number + Achar - 1);
  },
  //' map A..Z into 1..26
  A2N: function(Alphabet) {
    var Achar = 65;    //' A
    return Alphabet.toUpperCase().charCodeAt(0) - Achar + 1;
  },
  //' Column String to number
  Col2Num: function(ColString) {
    var Num = 0;
    for (var i = 0; i < ColString.length; i++ ) {
      Num = 26 * Num + this.A2N(ColString.charAt(i));
    }
    return Num;
  },
  //' number to Column String
  Num2Col: function(Number) {
    var Col = '';
    while (Number > 0) {
      var x = ((Number - 1) % 26) + 1;
      Col = this.N2A(x) + Col;
      Number = (Number - x) / 26;
    }
    return Col;
  },
  //' A1 to array
  A1RowCol: function(A1) {
    //' extract A1 format
    if (A1.match(/([a-zA-Z]+)\$?([0-9]*)/)) {
      var Col = this.Col2Num(RegExp.$1);
      var Row = parseInt(RegExp.$2);
      if (isNaN(Row)) { Row = 0; }
      return new Array(Row, Col);
    } else {
      return new Array(0, 0);
    }
  },
  //' array to A1
  RowColA1: function(RowCol) {
    if (!(RowCol instanceof Array)) {
      return '';
    } else if (RowCol.length < 2) {
      return '';
    } else if (RowCol[0] == 0) {
      return this.Num2Col(RowCol[1]);
    } else {
      return this.Num2Col(RowCol[1]) + RowCol[0];
    }
  },
  //' R1C1 to array
  R1C1RowCol: function(R1C1) {
    //' extract R1C1 absolute format
    if (R1C1.match(/[rR]([0-9]+)[cC]([0-9]+)/)) {
      var Row = parseInt(RegExp.$1);
      var Col = parseInt(RegExp.$2);
      return new Array(Row, Col);
    } else {
      return new Array(0, 0);
    }
  },
  //' array to R1C1
  RowColR1C1: function(RowCol) {
    if (!(RowCol instanceof Array)) {
      return '';
    } else if (RowCol.length < 2) {
      return '';
    } else {
      return 'R' + RowCol[0] + 'C' + RowCol[1];
    }
  },
  //' Key to array
  KeyRowCol: function(Key) {
    var x = Key.split(',');
    if (x.length < 2) {
      return new Array(0, 0);
    } else {
      return new Array(parseInt(x[0]), parseInt(x[1]));
    }
  },
  //' array to Key
  RowColKey: function(RowCol) {
    if (!(RowCol instanceof Array)) {
      return '';
    } else if (RowCol.length < 2) {
      return '';
    } else {
      return RowCol[0] + ',' + RowCol[1];
    }
  },
  //' extract Range size
  RangeSize: function(Address) {
    var x = new Array(2);
    var y = this.RangeStartEnd(Address);
    for (var i = 0; i < 2; i++) {
      x[i] = this.R1C1RowCol(y[i]);
      if (x[i][1] == 0) { x[i] = this.A1RowCol(y[i]); }
    }
    var Row1 = x[0][0];
    var Col1 = x[0][1];
    var Row2 = x[1][0];
    var Col2 = x[1][1];
    
    if ((Col1 == 0) || (Col2 == 0)) {
      return new Array(0);
    } else {
      return new Array(Col2 - Col1 + 1, Row1, Row2, Col1, Col2);
    }
  },
  //' extract Range
  RangeStartEnd: function(Address) {
    var StartAt; var EndAt;
    var x = Address.split(':');
    if (x.length == 1) {
      StartAt = Address;
      EndAt = StartAt;
    } else {
      StartAt = x[0];
      EndAt = x[1];
    }
    return new Array(StartAt, EndAt);
  },
  Terminate: function() {
  }
};

var AddressUtils = new AddressUtilsHelper;

// class StreamParser

function StreamParser() {
  this.Text = '';
  this.Stream;
}

StreamParser.prototype = {
  EOS: function() {
    return this.Stream.EOS;
  },
  AtEndOfStream: function() {
    return this.EOS() && (this.Text == '')
  },
  MoreText: function() {
    if (this.EOS()) { return ''; }
    
    var BuffSize = 8192;
    var out = this.Stream.ReadText(BuffSize);
    this.Text += out;
    return out;
  },
  FindString: function(StartAt, Search) {
    var Require = StartAt + Search.length - 1;
    while (this.Text.length < Require) {
        var more = this.MoreText();
        if (more == '') { break; }
    }
    
    var out = this.Text.indexOf(Search, StartAt - 1);
    while (out == -1) {
        var At = this.Text.length - Search.length + 2;
        var more = this.MoreText();
        if (more == '') { break; }
        
        out = this.Text.indexOf(Search, At - 1);
    }
    
    return out + 1;
  },
  Terminate: function() {
    delete this.Text;
  }
};

// class HtmlCellStream

function HtmlCellStream() {
  this.Text = new MatrixCells;
  this.Color = new MatrixCells;
  this.BackgroundColor = new MatrixCells;
  this.CurrentMatrix;
  this.Caption = '&nbsp;';
  this.LocalKey;
}

HtmlCellStream.prototype = {
  SetBlockName: function(NewName) {
    this.LocalKey = NewName;
    
    switch (this.LocalKey) {
    case 'cells-text':
    case 'cells-formula':
      this.CurrentMatrix = this.Text;
      break;
    case 'cells-color':
      this.CurrentMatrix = this.Color;
      break;
    case 'cells-background-color':
      this.CurrentMatrix = this.BackgroundColor;
      break;
    }
  },
  ReadBlock: function(Block) {
    this.CurrentMatrix.RepeatCell(1);
    
    for (var i in Block) {
      var Key = Block[i][0];
      var Value = Block[i][1];
      this.ReadSsfLine(Key, Value);
    }
  },
  ReadSsfLine: function(Key, Value) {
    switch (Key) {
    case 'address':
      this.CurrentMatrix.SetRange(Value);
      break;
    case 'repeat':
      this.CurrentMatrix.RepeatCell(parseInt(Value));
      break;
    case 'skip':
      this.CurrentMatrix.SkipCell(parseInt(Value));
      break;
    case '':
      this.CurrentMatrix.SetCell(Value);
      break;
    }
  },
  ToHtml: function() {
    var ExcelTABLE = 'border:medium ridge #66cc99;background-color:white;border-collapse:collapse;';
    var ExcelTH = 'border:thin outset white;background-color:#c7c7c7;color:black;padding-left:5px;padding-right:5px;';
    var ExcelTD = 'border:thin solid #c7c7c7;background-color:white;color:black;';
    var ExcelCAPTION = 'border:none;background-color:#66cc99;color:#7e434e;';
    
    var HasText = (this.Text.ColumnsCount > 0);
    if (!HasText) { return ''; }
    
    var HasColor = (this.Color.ColumnsCount > 0);
    var HasBackgroundColor = (this.BackgroundColor.ColumnsCount > 0);
    
    this.Text.DefaultValue = '&nbsp;';
    var TextData = this.Text.GetArray();
    var ColumnsHeader = this.Text.GetColumnsHeader();
    var RowsHeader = this.Text.GetRowsHeader();
    
    var out = new StringStream;
    
    out.WriteLine('<table style="' + ExcelTABLE + '">');
    out.WriteLine(' <tr>');
    out.WriteLine('  <td colspan="' + String(ColumnsHeader.length + 1) + '" style="' + ExcelCAPTION + '">' + this.Caption + '</td>');
    out.WriteLine(' </tr>');
    out.WriteLine(' <tr>');
    out.WriteLine('  <th style="' + ExcelTH + '">&nbsp;</th>');
    for (var C = 0; C < ColumnsHeader.length; C++) {
      out.WriteLine('  <th style="' + ExcelTH + '">' + ColumnsHeader[C] + '</th>');
    }
    out.WriteLine(' </tr>');
    for (var R = 0; R < RowsHeader.length; R++) {
      out.WriteLine(' <tr>');
      out.WriteLine('  <th style="' + ExcelTH + '">' + RowsHeader[R] + '</th>');
      for (var C = 0; C < ColumnsHeader.length; C++) {
        var CellKey = AddressUtils.RowColKey(new Array(R + this.Text.Row1, C + this.Text.Col1));
        var CellColor = this.Color.GetData(CellKey);
        var CellBackgroundColor = this.BackgroundColor.GetData(CellKey);
        var CellStyle = ExcelTD;
        if (CellColor != '') { CellStyle = CellStyle + 'color:' + CellColor + ';'; }
        if (CellBackgroundColor != '') { CellStyle = CellStyle + 'background-color:' + CellBackgroundColor + ';'; }
        out.WriteText('  <td');
        if (CellStyle != '') { out.WriteText(' style="' + CellStyle + '"'); }
        out.WriteLine('>' + TextData[R][C] + '</td>');
      }
      out.WriteLine(' </tr>');
    }
    out.WriteLine('</table>');
    
    var outHtml = out.Text;
    delete out;
    return outHtml;
  },
  Terminate: function() {
    delete this.Text;
    delete this.Color;
    delete this.BackgroundColor;
    delete this.CurrentMatrix;
    delete this.Caption;
    delete this.LocalKey;
  }
};

// class MatrixCells

function MatrixCells() {
  this.RawData = {};
  this.DefaultValue = '';
  this.Clear();
}

MatrixCells.prototype = {
  PopArray: function() {
    var out = this.GetArray();
    this.Clear();
    return out;
  },
  GetArray: function() {
    var Cs = this.ColumnsCount;
    if (Cs == 0) { Cs = 1; }
    if (this.R < this.Row2) { this.R = this.Row2; }
    var Rs = this.R - this.Row1 + 1;
    
    var AllRows = new Array(Rs);
    for (var i = 0; i < Rs; i++) {
      var EachRow = new Array(Cs);
      for (var j = 0; j < Cs; j++) {
        EachRow[j] = this.GetData(AddressUtils.RowColKey(new Array(i + this.Row1, j + this.Col1)))
      }
      AllRows[i] = EachRow;
    }
    return AllRows;
  },
  GetColumnsHeader: function() {
    var Cs = this.ColumnsCount;
    if (Cs == 0) { Cs = 1; }
    
    var Header = new Array(Cs);
    for (var i = 0; i < Cs; i++) {
      Header[i] = AddressUtils.RowColA1(new Array(0, i + this.Col1));
    }
    return Header;
  },
  GetRowsHeader: function() {
    if (this.R < this.Row2) { this.R = this.Row2; }
    var Rs = this.R - this.Row1 + 1;
    
    var Header = new Array(Rs);
    for (var i = 0; i < Rs; i++) {
      Header[i] = i + this.Row1;
    }
    return Header;
  },
  SetRange: function(Address) {
    var R1; var R2; var C1; var C2;
    var ret = AddressUtils.RangeSize(Address);
    if (ret[0] == 0) { return; }
    R1 = ret[1];
    R2 = ret[2];
    C1 = ret[3];
    C2 = ret[4];
    
    if (this.ColumnsCount == 0) {
      this.Col1 = C1;
      this.Col2 = C2;
      this.Row1 = R1;
      this.Row2 = R2;
    } else {
      if (this.Col1 > C1) { this.Col1 = C1; }
      if (this.Col2 < C2) { this.Col2 = C2; }
      if (this.Row1 > R1) { this.Row1 = R1; }
      if (this.Row2 < R2) { this.Row2 = R2; }
    }
    
    this.ColumnsCount = this.Col2 - this.Col1 + 1;
    this.R = R1;
    this.C = C1;
  },
  SetCell: function(Value) {
    while (this.RepeatCount > 0) {
      this.SetData(AddressUtils.RowColKey(new Array(this.R, this.C)), Value);
      this.NextCell();
      this.RepeatCount--;
    }
    this.RepeatCount = 1;
  },
  RepeatCell: function(Count) {
    this.RepeatCount = Count;
  },
  SkipCell: function(Count) {
    while (Count > 0) {
      this.NextCell();
      Count--;
    }
  },
  NextCell: function() {
    if (this.C == this.Col2) {
      this.C = this.Col1;
      this.R++;
    } else {
      this.C++;
    }
  },
  PopData: function(Key) {
    var out = this.RawData[Key.toUpperCase()];
    if (out == undefined) {
      out = this.DefaultValue;
    } else {
      delete this.RawData[Key.toUpperCase()];
    }
    return out;
  },
  GetData: function(Key) {
    var out = this.RawData[Key.toUpperCase()];
    if (out == undefined) {
      out = this.DefaultValue;
    }
    return out;
  },
  SetData: function(Key, Value) {
    this.RawData[Key.toUpperCase()] = Value;
  },
  Clear: function() {
    for (var Key in this.RawData) { delete this.RawData[Key]; }
    this.R = 1;
    this.C = 1;
    this.RepeatCount = 1;
    this.Row1 = 1;
    this.Col1 = 1;
    this.Row2 = 0;
    this.Col2 = 0;
    this.ColumnsCount = 0;
  },
  Terminate: function() {
    delete this.RawData;
    delete this.DefaultValue;
    delete this.ColumnsCount;
    delete this.Row1;
    delete this.Row2;
    delete this.Col1;
    delete this.Col2;
    delete this.R;
    delete this.C;
    delete this.RepeatCount;
  }
};

// class StringStream

function StringStream() {
  this.Text = '';
  this.EOS = true;
}

StringStream.prototype = {
  WriteText: function(Data) {
    this.Text += Data;
  },
  WriteLine: function(Data) {
    this.WriteText(Data + '\n');
  },
  ReadText: function(Size) {
    // ignore Size
    var out = this.Text;
    this.Text = '';
    this.EOS = true;
    return out;
  },
  Terminate: function() {
    delete this.Text;
    delete this.EOS;
  }
};

// test
function Test() {
  // test StringStream
  if (false) { (function () {
    var x = new StringStream;
    x.Text = 'abc\ndef';
    x.WriteText('cat');
    x.WriteLine('dog');
    x.WriteLine('cow');
    WScript.echo(x.Text);
    x.Text = 'reading';
    x.EOS = false;
    while (!x.EOS) {
      WScript.echo(x.ReadText(12345));
    }
    x.Terminate();
  })(); }
  
  // test MatrixCells
  if (false) { (function () {
    var x = new MatrixCells;
    for (var k in x.RawData) { WScript.echo(k + '=>' + x.RawData[k]); }
    x.SetData('abc', 'def');
    x.SetData('gh','ij');
    for (var k in x.RawData) { WScript.echo(k + '=>' + x.RawData[k]); }
    x.DefaultValue = 'not found';
    WScript.echo(x.GetData('aBc'));
    WScript.echo(x.GetData('aBcd'));
    for (var k in x.RawData) { WScript.echo(k + '=>' + x.RawData[k]); }
    WScript.echo(x.PopData('gh'));
    for (var k in x.RawData) { WScript.echo(k + '=>' + x.RawData[k]); }
    x.Clear();
    with (x) {
      WScript.echo([R,C,RepeatCount,ColumnsCount,Row1,Row2,Col1,Col2].join(' '));
    }
    x.SetRange('C2:E6');
    with (x) {
      WScript.echo([R,C,RepeatCount,ColumnsCount,Row1,Row2,Col1,Col2].join(' '));
    }
    WScript.echo(x.GetColumnsHeader().toString());
    WScript.echo(x.GetRowsHeader().toString());
    for (var i = 101; i < 108; i++) {
      x.SetCell(i);
    }
    WScript.echo(x.GetRowsHeader().toString());
    WScript.echo(x.GetArray().toString());
    WScript.echo(x.PopArray().toString());
    WScript.echo(x.GetArray().toString());
    x.Terminate();
  })(); }
  
  // test HtmlCellStream
  if (false) { (function () {
    var x = new HtmlCellStream;
    x.SetBlockName('cells-text');
    x.ReadBlock([['address','B3:D4'],['',1],['',2],['',3],['',4],['',5]]);
    WScript.echo(x.Text.ColumnsCount);
    //with (x.CurrentMatrix) {
    with (x.Text) {
      WScript.echo(GetArray().toString());
    }
    WScript.echo(x.ToHtml());
    x.Terminate();
  })(); }
  
  // test StreamParser
  if (false) { (function () {
    var x = new StreamParser;
    var adTypeText = 2;
    var Stream = new ActiveXObject('ADODB.Stream');
    Stream.Open();
    Stream.Type = adTypeText;
    Stream.Charset = 'utf-8';
    Stream.LoadFromFile('C:\\tmp\\test.txt');
    
    x.Stream = Stream;
    WScript.echo(x.AtEndOfStream());
    //WScript.echo(typeof x.AtEndOfStream());
    WScript.echo(x.FindString(1,'ssf-begin'));
    //WScript.echo(x.Text);
    WScript.echo(x.FindString(1,'cells-formula'));
    WScript.echo(x.Text.substring(120,133));
    
    Stream.Close();
    delete Stream;
    x.Terminate();
  })(); }
  
  if (false) { (function () {
    var x = new StreamParser;
    var adTypeText = 2;
    var Stream = new StringStream;
    Stream.Text = "'cells-formula\r\n'address;A1:B2\r\n';123";
    Stream.EOS = false;
    
    x.Stream = Stream;
    WScript.echo(x.AtEndOfStream());
    WScript.echo(x.Text);
    WScript.echo(x.FindString(1,'cells-formula'));
    
    Stream.Terminate();
    delete Stream;
    x.Terminate();
  })(); }
  
  // test AddressUtils
  if (false) { (function () {
    WScript.echo(AddressUtils.N2A(5));
    WScript.echo(AddressUtils.A2N('C'));
    WScript.echo(AddressUtils.Col2Num('AA'));
    WScript.echo(AddressUtils.Num2Col(256));
    WScript.echo(AddressUtils.A1RowCol('F12').toString());
    WScript.echo(AddressUtils.A1RowCol('G').toString());
    WScript.echo(AddressUtils.RowColA1([12, 6]));
    WScript.echo(AddressUtils.RowColA1([0, 256]));
    WScript.echo(AddressUtils.R1C1RowCol('R23C45').toString());
    WScript.echo(AddressUtils.R1C1RowCol('A1').toString());
    WScript.echo(AddressUtils.RowColR1C1([12, 6]));
    WScript.echo(AddressUtils.KeyRowCol('23,45').toString());
    WScript.echo(AddressUtils.RowColKey([12, 6]));
    WScript.echo(AddressUtils.RangeSize('B3:F12').toString());
    WScript.echo(AddressUtils.RangeStartEnd('B3:F12').toString());
  })(); }
}

//Test();


var rc = 0;
Main('C:\\tmp\\test.txt', 'utf-8', 'Worksheet Title');
try {
  //Main('C:\\tmp\\test.txt', 'utf-8', 'Worksheet Title');
} catch (e) {
  rc = e.Number;
  WScript.Echo('Error: ' + e.Description);
} finally {
  WScript.Quit(rc);
}


