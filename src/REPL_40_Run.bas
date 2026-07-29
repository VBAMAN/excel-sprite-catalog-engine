Attribute VB_Name = "REPL_40_Run"
Option Explicit

'=========================================================
' REPL_40_Run
'
' Replace Engine
'
' 1. Screen Area をスプライトコードへ変換
' 2. Dummy Area にカタログを生成
'=========================================================

Public Sub REPL_Run(ByVal Ws As Worksheet)

    Application.EnableEvents = False

    On Error GoTo ExitProc

    '--------------------------------------
    ' Replace
    '--------------------------------------
    REPL_ReplaceRange Ws, REPL_SCREEN_AREA

    '--------------------------------------
    ' Catalog
    '--------------------------------------
    REPL_CreateCatalog Ws, REPL_SCREEN_AREA, REPL_DUMMY_AREA

ExitProc:

    Application.EnableEvents = True

End Sub
