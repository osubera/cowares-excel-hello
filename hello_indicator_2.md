# Introduction #

  * how to use a notepad as an active indicator informing the progress in vba

## 概要 ##
  * VBAでメモ帳を使って進行状況をライブ表示する

# Details #

  * another solution of [hello\_indicator](hello_indicator.md).
  * we use a notepad this time.

## 説明 ##
  * [hello\_indicator](hello_indicator.md) の別解。
  * 今度はメモ帳を使う。

# How to use #

  1. run a macro `test_NotepadIndicator`
    1. a notepad appears to show the progress, and disappears on done.
    1. close the notepad window to stop the macro interructively.
  1. run a macro `test_NotepadLogger`
    1. a notepad appears to show the progress, and disappears on done.
    1. close the notepad window to stop the macro interructively.
    1. this is similar as the former, except appending messages like a logger.

## 使い方 ##
  1. マクロ `test_NotepadIndicator` を実行する
    1. メモ帳が現れて、進行状況を表示し、終了すると消える。
    1. マクロを途中で中断するときは、メモ帳ウィンドウを閉じる。
  1. マクロ `test_NotepadLogger` を実行する
    1. メモ帳が現れて、進行状況を表示し、終了すると消える。
    1. マクロを途中で中断するときは、メモ帳ウィンドウを閉じる。
    1. 先のとほぼ同じだが、こちらはログのようにメッセージを追記する。

# Code #

```
'workbook
'  name;hello_indicator_2.xls


'module
'  name;DoNotepad
'{{{
Option Explicit
 
Private Const GW_CHILD = 5
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_CLOSE = &H10
Private Const EM_REPLACESEL = &HC2
Private Const EM_SETSEL = &HB1
Private Const EM_SETMODIFY = &HB9
Private Const HWND_BOTTOM = 1
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SW_RESTORE = 9
 
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 
' hWndで指定したメモ帳の未保存フラグをクリアする。
' clear the save me flag
Public Function SetSavedNotepad(hWnd As Long) As Long
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    SendMessage i, EM_SETMODIFY, 0, 0
    SetSavedNotepad = i
End Function
 
' hWndで指定したメモ帳を閉じる。
' close the notepad
Public Sub CloseNotepad(hWnd As Long)
    SetSavedNotepad hWnd
    SendMessage hWnd, WM_CLOSE, 0, 0
End Sub
 
' メモ帳を新しく起動し、hWndを返す。
' kick up a new notepad process, return the hWnd
Public Function OpenNotepad(Optional iWindowState As Long = vbNormalFocus, _
            Optional NameMe As String = "") As Long
    Dim hWnd As Long
    Dim ProcID As Long
    Dim i As Long
    Dim TitleText As String
    Dim ExePath As String
    
    On Error GoTo Err1
    
    TitleText = " - notepad - meets VBA"
    'TitleText = "無題 - ﾒﾓ帳"
    ExePath = "notepad.exe"
    
    ProcID = Shell(ExePath, iWindowState)
    If ProcID = 0 Then GoTo Err1
    
    hWnd = GetWindowByProcessId(ProcID)
    If hWnd = 0 Then GoTo Err1
    
    TitleText = IIf(NameMe = "", ProcID, NameMe) & TitleText
    i = SetWindowText(hWnd, TitleText)
    'MoveWindow hWnd, 0, 50, 300, 200, 1
    ' Z-order も変えるなら SetWindowPos
    'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    OpenNotepad = hWnd
    Exit Function
Err1:
    MsgBox "メモ帳の起動に失敗しました。", , "failed to start a notepad"
    OpenNotepad = 0
End Function
 
' hWndで指定したメモ帳の内容を、指定した文字で置き換える。
' repalce text at the notepad
Public Function WriteNotepad(hWnd As Long, strTextAll As String) As Boolean
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    WriteNotepad = _
        (0 <> SendMessageStr(i, WM_SETTEXT, 0, strTextAll))
End Function
 
' hWndで指定したメモ帳に、指定した文字を追加する。改行つき。
' push text into the notepad with a linefeed
' iPos=0: 現在のカーソル位置 / at a cursor position
'     -1: 先頭              / at the first
'      1: 最後              / at the last
Public Function WriteLineNotepad(hWnd As Long, strText As String, Optional iPos As Long = 0) As Boolean
    WriteLineNotepad = WriteTextNotepad(hWnd, strText & vbNewLine, iPos)
End Function
 
' hWndで指定したメモ帳に、指定した文字を追加する。改行無し。
' push text into the notepad without a linefeed
' iPos=0: 現在のカーソル位置 / at a cursor position
'     -1: 先頭              / at the first
'      1: 最後              / at the last
Public Function WriteTextNotepad(hWnd As Long, strText As String, Optional iPos As Long = 0) As Boolean
    Dim i As Long
    i = GetWindow(hWnd, GW_CHILD)
    Select Case iPos
    Case -1
        SendMessage i, EM_SETSEL, 0, 0
    Case 1
        SendMessage i, EM_SETSEL, 0, -1     ' 全部選択
        SendMessage i, EM_SETSEL, -1, 0     ' 選択解除 (カーソルが選択領域の最後に移動)
    End Select
    WriteTextNotepad = _
        (0 <> SendMessageStr(i, EM_REPLACESEL, 0, strText))
End Function
 
' hWndで指定したメモ帳の内容を、文字として取得する。
' get text from the notepad
Public Function ReadNotepad(hWnd As Long) As String
    Dim i As Long
    Dim j As Long
    Dim x As String
    i = GetWindow(hWnd, GW_CHILD)
    j = 1 + SendMessage(i, WM_GETTEXTLENGTH, 0, 0)
    x = String(j, Chr(0))
    SendMessageStr i, WM_GETTEXT, j, x
    ReadNotepad = x
End Function

' hWnd から ProcessID を取得する。
Public Function GetWindowProcessId(hWnd As Long) As Long
    Dim ProcID As Long
    Dim ThreadID As Long
    ThreadID = GetWindowThreadProcessId(hWnd, ProcID)
    GetWindowProcessId = ProcID
End Function

' ProcessID から hWnd を取得する。(メモ帳)
Public Function GetWindowByProcessId(ProcessId As Long, _
        Optional TaskName As String = "Notepad", _
        Optional TitleText As String = vbNullString) As Long
    Dim ProcID As Long
    Dim ThreadID As Long
    Dim hWnd As Long

    hWnd = 0
    Do
        hWnd = FindWindowEx(0, hWnd, TaskName, TitleText)
        If hWnd = 0 Then Exit Do
        ThreadID = GetWindowThreadProcessId(hWnd, ProcID)
    Loop Until ProcessId = ProcID
    
    GetWindowByProcessId = hWnd
End Function

' hWndで指定したメモ帳を、ユーザーに視覚で通知する。
Public Function ShowNotepad(hWnd As Long) As Boolean
    Dim Result As Long
    Result = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    ShowWindow hWnd, SW_RESTORE
    ShowNotepad = (Result <> 0)
End Function
'}}}

'module
'  name;Module1
'{{{
Option Explicit

Private Const HWND_TOPMOST = -1
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Sub test_NotepadIndicator()
    Dim St As Long
    Dim i As Long
    Dim Msg As String
    Const iStart = 0
    Const iEnd = 100
    
    St = OpenNotepad(NameMe:="hello indicator 2 アプリケーション実行中")
    StyleUp St
    
    For i = iStart To iEnd
        HeavyTask
        
        Msg = "実行中 " & i & vbCrLf & "残り  " & iEnd - i & vbCrLf & "Close this window to Stop"
        If Not WriteNotepad(St, Msg) Then Exit For
        
        DoEvents
    Next
    
    CloseNotepad St
    MsgBox "Stop at " & i
End Sub

Public Sub test_NotepadLogger()
    Dim St As Long
    Dim i As Long
    Dim Msg As String
    Const iStart = 0
    Const iEnd = 100
    
    St = OpenNotepad(NameMe:="hello indicator 2 アプリケーション実行中")
    StyleUp St
    
    For i = iStart To iEnd
        HeavyTask
        
        Msg = "実行中 " & i & vbCrLf & "残り  " & iEnd - i & vbCrLf & "Close this window to Stop"
        If Not WriteLineNotepad(St, Msg) Then Exit For
        
        SetSavedNotepad St
        DoEvents
    Next
    
    CloseNotepad St
End Sub

Private Sub HeavyTask()
    Application.Wait Now() + Rnd(2) / 24 / 60 / 60
End Sub

Private Sub StyleUp(hWnd As Long)
    ' locate at the center of the application
    Dim NewWidth As Long
    Dim NewHeight As Long
    Dim NewLeft As Long
    Dim NewTop As Long
    
    NewWidth = 300
    NewHeight = 160
    LocateCenter NewLeft, NewTop, NewWidth, NewHeight
    SetWindowPos hWnd, HWND_TOPMOST, NewLeft, NewTop, NewWidth, NewHeight, 0
End Sub

Private Sub LocateCenter(ByRef Left As Long, ByRef Top As Long, ByRef Width As Long, ByRef Height As Long)
    On Error GoTo AccessWillFail
    Dim dpi As Variant
    Const ppi = 72  ' Excel uses this point per inch for screen unit
    
    With Application
        Left = .Left + .Width / 2 - Width / 2
        Top = .Top + .Height / 2 - Height / 2
    End With
    If Left < 0 Then Left = 0
    If Top < 0 Then Top = 0
    
    dpi = ScreenDPI
    Left = Left * dpi(0) / ppi
    Top = Top * dpi(1) / ppi
    Width = Width * dpi(0) / ppi
    Height = Height * dpi(1) / ppi
    Exit Sub
    
AccessWillFail:
    Left = 100
    Top = 100
End Sub

Private Function ScreenDPI() As Variant
    Dim DC As Long
    DC = GetDC(HWND_DESKTOP)
    ScreenDPI = Array(GetDeviceCaps(DC, LOGPIXELSX), GetDeviceCaps(DC, LOGPIXELSY))
    ReleaseDC HWND_DESKTOP, DC
End Function
'}}}



```