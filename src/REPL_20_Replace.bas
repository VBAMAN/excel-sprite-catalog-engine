Attribute VB_Name = "REPL_20_Replace"
Option Explicit

'=========================================================
' REPL_20_Replace
'
' 指定範囲をスプライトコードへ変換する
'
' 0  → 00
' 1  → 01
' ...
' 9  → 09
'
' A  → 0A
' ...
' Z  → 0Z
'
' X  → 0X
' /  → 0S
' -  → 0M
'
' "10" はそのまま
'=========================================================

Public Sub REPL_ReplaceRange( _
    ByVal Ws As Worksheet, _
    ByVal TargetRange As String)

    Dim c As Range

    Application.ScreenUpdating = False

    For Each c In Ws.Range(TargetRange)

        If Len(c.Value) > 0 Then
            c.Value = REPL_Normalize(CStr(c.Value))
        End If

    Next

    Application.ScreenUpdating = True

End Sub


'=========================================================
' 文字列を2文字コードへ変換
'=========================================================

Private Function REPL_Normalize(ByVal Text As String) As String

    Dim s As String

    s = Trim$(UCase$(Text))

    Select Case s

        '-------------------------
        ' 特殊記号
        '-------------------------
        Case "X"
            REPL_Normalize = REPL_CODE_STRIKE

        Case "/"
            REPL_Normalize = REPL_CODE_SPARE

        Case "-"
            REPL_Normalize = REPL_CODE_MINUS

        '-------------------------
        ' フレーム番号10
        '-------------------------
        Case "10"
            REPL_Normalize = "10"

        '-------------------------
        ' 数字
        '-------------------------
        Case "0" To "9"
            REPL_Normalize = REPL_PREFIX_NUMBER & s

        '-------------------------
        ' 英字
        '-------------------------
        Case "A" To "Z"
            REPL_Normalize = REPL_PREFIX_ALPHA & s

        '-------------------------
        ' その他
        '-------------------------
        Case Else
            REPL_Normalize = s

    End Select

End Function

