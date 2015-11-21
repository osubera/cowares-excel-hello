

# Introduction #

  1. a primitive tool to convert a Word Document into an SSF text

## 概要 ##
  1. ワードドキュメントを SSF テキストに変換する、原始的なツール

# Details #

  * this reader is a very simple tool for a quick start.
  * it supports VBA code modules only.
  * it works in a single VBA module.
  * [ssf\_reader\_primitive\_word](ssf_reader_primitive_word.md) is also available as a simple reader.

## 説明 ##
  * このSSF読み込み装置は、 迅速な初動に適した、非常に簡単なものだ。
  * このツールは、 VBA コードだけを扱う。
  * このツールは、たった一つの VBA モジュールで動いている。
  * 対応する読み込み装置として [ssf\_reader\_primitive\_word](ssf_reader_primitive_word.md) がある。

# Downloads #

  * [downloads](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=ssf+writer)

# How to use #

  1. add a vba module into your blank word document.
  1. paste the below Code on the module.
  1. add references for
    * Microsoft Scripting Runtime
    * Microsoft Visual Basic for Applications Extensibility 5.3
  1. change the default filename at the topmost of the module code.
    * DefaultFileName = "C:\Users\Public\Documents\ssf.txt"
    * it gives a dialog to choose a file.
  1. save it as a word template (dot file), as an ssf\_writer\_primitive.dot.
  1. activate the addin ssf\_writer\_primitive.dot.
  1. open a word document to convert.
  1. run the macro, "Writer".
  1. a text file will be written.
  1. [Trust access to the VBA project object model](vba#how_to_allow_access_to_a_Visual_Basic_project.md) is required.

## 使い方 ##
  1. 新規のワード文書にVBAのモジュールを挿入する。
  1. そのモジュールに、下記の Code を貼り付ける。
  1. 次の参照設定を追加する。
    * Microsoft Scripting Runtime
    * Microsoft Visual Basic for Applications Extensibility 5.3
  1. モジュールコードの先頭近くにあるファイル名を変更する。
    * DefaultFileName = "C:\Users\Public\Documents\ssf.txt"
    * ファイルダイアログが出て、ファイルを選べる。
  1. この文書をワードテンプレート（dotファイル）、 ssf\_writer\_primitive.dot として保存する。
  1. アドインの ssf\_writer\_primitive.dot を有効にする。
  1. 変換したいワード文書を開く。
  1. マクロ "Writer" を実行する。
  1. テキストファイルが保存される。
  1. [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する](vba#Visual_Basic_%E3%83%97%E3%83%AD%E3%82%B8%E3%82%A7%E3%82%AF%E3%83%88%E3%81%B8%E3%81%AE%E3%82%A2%E3%82%AF%E3%82%BB%E3%82%B9%E3%82%92%E4%BF%A1%E9%A0%BC%E3%81%95%E3%81%9B.md) が必要。

# Code #

```
'ssf-begin
';

'document
'  name;ssf_writer_primitive.doc

'require
'  ;{420B2830-E718-11CF-893D-00A0C9054228} 1 0 Microsoft Scripting Runtime
'  ;{0002E157-0000-0000-C000-000000000046} 5 3 Microsoft Visual Basic for Applications Extensibility 5.3

'module
'  name;SsfWriterPrimitive
'{{{
Option Explicit

Function GetFileName() As String
    Const DefaultFileName = "C:\Users\Public\Documents\ssf.txt"
    Dim D As Dialog
    
    Set D = Dialogs(wdDialogFileSaveAs)
    D.Name = DefaultFileName
    If Not D.Display Then Err.Raise 53  ' File Not Found
    
    GetFileName = Application.WordBasic.FileNameInfo(D.Name, 1)
End Function

Public Sub Writer()
    Dim fs As Scripting.FileSystemObject
    Dim Stream As Scripting.TextStream
    Set fs = New Scripting.FileSystemObject
    Set Stream = fs.OpenTextFile(GetFileName, ForWriting, True, TristateFalse)
    
    WriteTo Stream
    
    Stream.Close
    Set Stream = Nothing
    Set fs = Nothing
End Sub

Function WriteTo(Stream As Scripting.TextStream) As String
    Stream.Write DumpDocument(ActiveDocument)
End Function

Function DumpDocument(Book As Document) As String
    Dim Result As String
    
    Result = "'ssf-begin" & vbCrLf & "';" & vbCrLf & vbCrLf
    Result = Result & "'document" & vbCrLf
    Result = Result & "'  name;" & Book.Name
    Result = Result & vbCrLf    ' end of name line
    Result = Result & vbCrLf    ' end of document block
    Result = Result & DumpProjectRequires(Book)
    ' put this at the last: http://code.google.com/p/cowares-excel-hello/wiki/hello_thisworkbook#Case_2
    Result = Result & DumpVbaCodes(Book)
    Result = Result & "'ssf-end" & vbCrLf & vbCrLf
    
    DumpDocument = Result
End Function

Function DumpProjectRequires(Book As Document) As String
    Dim Result As String
    Dim Project As VBProject
    Dim NumberOfReferences As Long
    Dim i As Long
    
    DumpProjectRequires = ""
    Set Project = Book.VBProject
    NumberOfReferences = Project.References.Count
    If NumberOfReferences = 0 Then Exit Function
    ' it doesn't work, because we have at least 4 references.
    
    Result = "'require" & vbCrLf
    For i = 1 To NumberOfReferences
        ' avoid to print 5 standard references
        ' VBA (builtin), Word (builtin), stdole and Office
        ' normal.dot has a blank name and guid
        If Not Project.References(i).BuiltIn Then
            If LCase(Project.References(i).Name) <> "stdole" _
                    And LCase(Project.References(i).Name) <> "office" _
                    And Project.References(i).GUID <> "" Then
                ' machine needs Guid, Major and Minor.  human needs Description
                Result = Result & "'  ;" & Project.References(i).GUID & " "
                Result = Result & Project.References(i).Major & " "
                Result = Result & Project.References(i).Minor & " "
                Result = Result & Project.References(i).Description & vbCrLf
            End If
        End If
    Next
    
    Result = Result & vbCrLf
    
    DumpProjectRequires = Result
End Function

Function DumpVbaCodes(Book As Document) As String
    Dim Result As String
    Dim Module As VBComponent
    Dim BookModule As VBComponent
    
    Result = ""
    ' let ThisDocument go to the last
    For Each Module In Book.VBProject.VBComponents
        If Module.Name = "ThisDocument" Then
            Set BookModule = Module
        Else
            Result = Result & DumpVbaCodeModule(Module.CodeModule)
        End If
    Next
    If Not BookModule Is Nothing Then
        Result = Result & DumpVbaCodeModule(BookModule.CodeModule)
    End If
    
    DumpVbaCodes = Result
End Function

Function DumpVbaCodeModule(TheCode As CodeModule) As String
    Dim Result As String
    Dim ModuleType As String
    Dim NumberOfLines As Long
    Dim Source As String
    
    Result = ""
    Select Case TheCode.Parent.Type
    Case vbext_ct_StdModule
        ModuleType = "module"   ' Module
    Case vbext_ct_ClassModule
        ModuleType = "class"    ' Class
    Case vbext_ct_MSForm
        ModuleType = "form"     ' not for Excel 2000
    Case vbext_ct_ActiveXDesigner
        ModuleType = "activex"
    Case vbext_ct_Document
        ModuleType = "code"     ' Word Objects
    Case Else
        ModuleType = "unknown-type-" & TheCode.Parent.Type
    End Select
    NumberOfLines = TheCode.CountOfLines
    If NumberOfLines > TheCode.CountOfDeclarationLines Then
        ' avoid to print a blank code, that contains "Option Explicit" only
        Result = Result & "'" & ModuleType & vbCrLf
        Result = Result & "'  name;" & TheCode.Parent.Name & vbCrLf
        Result = Result & "'{{{" & vbCrLf
        
        Source = TheCode.Lines(1, NumberOfLines)
        ' need at least one linefeed on the end
        If Right(Source, 2) <> vbCrLf Then
            Source = Source & vbCrLf
        End If
        ' must disable escaping signs in the source itself
        Source = Replace(Source, vbCrLf & "'{{{" & vbCrLf, vbCrLf & "'#{{{" & vbCrLf)
        Source = Replace(Source, vbCrLf & "'}}}" & vbCrLf, vbCrLf & "'#}}}" & vbCrLf)
        
        Result = Result & Source
        Result = Result & "'}}}" & vbCrLf
    End If
    Result = Result & vbCrLf
    
    DumpVbaCodeModule = Result
End Function
'}}}


'ssf-end

```

