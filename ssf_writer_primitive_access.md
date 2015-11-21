

# Introduction #

  1. a primitive tool to convert an Access Mdb into an SSF text

## 概要 ##
  1. アクセスMDBファイルを SSF テキストに変換する、原始的なツール

# Details #

  * this reader is a very simple tool for a quick start.
  * it supports VBA code modules only.
  * it works in a single VBA module.
  * [ssf\_reader\_primitive\_access](ssf_reader_primitive_access.md) is also available as a simple reader.

## 説明 ##
  * このSSF読み込み装置は、 迅速な初動に適した、非常に簡単なものだ。
  * このツールは、 VBA コードだけを扱う。
  * このツールは、たった一つの VBA モジュールで動いている。
  * 対応する読み込み装置として [ssf\_reader\_primitive\_access](ssf_reader_primitive_access.md) がある。

# Downloads #

  * [downloads](http://code.google.com/p/cowares-excel-hello/downloads/list?can=2&q=ssf+writer)

# How to use #

  1. prepare a blank access mdb as ssf\_writer\_primitive.mdb.
  1. add a vba module into it.
  1. paste the below Code on the module.
  1. add references for
    * Microsoft Scripting Runtime
    * Microsoft Visual Basic for Applications Extensibility 5.3
  1. change filenames at the topmost of the module code.
```
Function GetMdbFileName() As String
    'GetMdbFileName = "C:\Users\Public\Documents\ssf.mdb"
    GetMdbFileName = "C:\tmp\ssf.mdb"
End Function

Function GetSsfFileName() As String
    'GetSsfFileName = "C:\Users\Public\Documents\ssf.txt"
    GetSsfFileName = "C:\tmp\ssf.txt"
End Function
```
    * the mdb file name is a mdb file to read.
    * the ssf file name is an ssf text to write.
    * mdb file must exist before run.
    * ssf file will be overwritten if exists.
  1. save it.
  1. run the macro, "Writer".
  1. a text file will be written.
  1. [Trust access to the VBA project object model](vba#how_to_allow_access_to_a_Visual_Basic_project.md) is required.

## 使い方 ##
  1. 新規のアクセスMDBとして、 ssf\_writer\_primitive.mdb を用意する。
  1. そのMDBにVBAのモジュールを挿入する。
  1. そのモジュールに、下記の Code を貼り付ける。
  1. 次の参照設定を追加する。
    * Microsoft Scripting Runtime
    * Microsoft Visual Basic for Applications Extensibility 5.3
  1. モジュールコードの先頭近くにあるファイル名を変更する。
```
Function GetMdbFileName() As String
    'GetMdbFileName = "C:\Users\Public\Documents\ssf.mdb"
    GetMdbFileName = "C:\tmp\ssf.mdb"
End Function

Function GetSsfFileName() As String
    'GetSsfFileName = "C:\Users\Public\Documents\ssf.txt"
    GetSsfFileName = "C:\tmp\ssf.txt"
End Function
```
    * MDB ファイル名は、読み込む対象を指定。
    * SSF ファイル名は、書き込み先を指定。
    * MDB ファイルは、実行前に用意しておく。
    * SSF ファイルは上書きされる。
  1. 保存する。
  1. マクロ "Writer" を実行する。
  1. テキストファイルが保存される。
  1. [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する](vba#Visual_Basic_%E3%83%97%E3%83%AD%E3%82%B8%E3%82%A7%E3%82%AF%E3%83%88%E3%81%B8%E3%81%AE%E3%82%A2%E3%82%AF%E3%82%BB%E3%82%B9%E3%82%92%E4%BF%A1%E9%A0%BC%E3%81%95%E3%81%9B.md) が必要。

# Code #

```
'ssf-begin
';

'mdb
'  name;ssf_writer_primitive

'require
'  ;{00000201-0000-0010-8000-00AA006D2EA4} 2 1 Microsoft ActiveX Data Objects 2.1 Library
'  ;{420B2830-E718-11CF-893D-00A0C9054228} 1 0 Microsoft Scripting Runtime
'  ;{0002E157-0000-0000-C000-000000000046} 5 3 Microsoft Visual Basic for Applications Extensibility 5.3

'module
'  name;SsfWriterPrimitive
'{{{
Option Compare Database
Option Explicit

Function GetMdbFileName() As String
    'GetMdbFileName = "C:\Users\Public\Documents\ssf.mdb"
    GetMdbFileName = "C:\tmp\ssf.mdb"
End Function

Function GetSsfFileName() As String
    'GetSsfFileName = "C:\Users\Public\Documents\ssf.txt"
    GetSsfFileName = "C:\tmp\ssf.txt"
End Function

Public Sub Writer()
    Dim fs As Scripting.FileSystemObject
    Dim Stream As Scripting.TextStream
    Dim App As Access.Application
    
    Set fs = New Scripting.FileSystemObject
    Set Stream = fs.OpenTextFile(GetSsfFileName, ForWriting, True, TristateFalse)
    Set App = GetObject(GetMdbFileName)
    App.Visible = True
    App.DoCmd.RunCommand acCmdDebugWindow
    
    WriteTo Stream, App
    
    App.Quit acQuitSaveNone
    Stream.Close
    
    Set App = Nothing
    Set Stream = Nothing
    Set fs = Nothing
End Sub

Function WriteTo(Stream As Scripting.TextStream, App As Access.Application) As String
    Stream.Write DumpMdb(App)
End Function

Function DumpMdb(App As Access.Application) As String
    Dim Result As String
    
    Result = "'ssf-begin" & vbCrLf & "';" & vbCrLf & vbCrLf
    Result = Result & "'mdb" & vbCrLf
    Result = Result & "'  name;" & App.VBE.ActiveVBProject.Name
    Result = Result & vbCrLf    ' end of name line
    Result = Result & vbCrLf    ' end of document block
    Result = Result & DumpProjectRequires(App)
    ' put this at the last: http://code.google.com/p/cowares-excel-hello/wiki/hello_thisworkbook#Case_2
    Result = Result & DumpVbaCodes(App)
    Result = Result & "'ssf-end" & vbCrLf & vbCrLf
    
    DumpMdb = Result
End Function

Function DumpProjectRequires(App As Access.Application) As String
    Dim Result As String
    Dim Project As VBProject
    Dim NumberOfReferences As Long
    Dim i As Long
    
    DumpProjectRequires = ""
    Set Project = App.VBE.ActiveVBProject
    NumberOfReferences = Project.References.Count
    If NumberOfReferences = 0 Then Exit Function
    ' it doesn't work, because we have at least 4 references.
    
    Result = "'require" & vbCrLf
    For i = 1 To NumberOfReferences
        ' avoid to print 3 standard references
        ' VBA (builtin), Access (builtin) and stdole
        If Not Project.References(i).BuiltIn Then
            If LCase(Project.References(i).Name) <> "stdole" Then
                ' machine needs Guid, Major and Minor.  human needs Description
                Result = Result & "'  ;" & Project.References(i).Guid & " "
                Result = Result & Project.References(i).Major & " "
                Result = Result & Project.References(i).Minor & " "
                Result = Result & Project.References(i).Description & vbCrLf
            End If
        End If
    Next
    
    Result = Result & vbCrLf
    
    DumpProjectRequires = Result
End Function

Function DumpVbaCodes(App As Access.Application) As String
    Dim Result As String
    Dim Module As VBComponent
    
    Result = ""
    ' let ThisDocument go to the last
    For Each Module In App.VBE.ActiveVBProject.VBComponents
        Result = Result & DumpVbaCodeModule(Module.CodeModule)
    Next
    
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
        ModuleType = "code"     ' UserForm? Objects
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

