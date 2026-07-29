Attribute VB_Name = "REPL_30_Catalog"
Option Explicit

'=========================================================
' REPL_30_Catalog
'
' REPL_SCREEN_AREA に存在するユニークな文字列を
' REPL_DUMMY_AREA へ一覧出力する
'
' ※ 空白セルは無視
' ※ "10" は 1つの文字列として扱う
'=========================================================

Public Sub REPL_CreateCatalog( _
    ByVal Ws As Worksheet, _
    ByVal SourceRange As String, _
    ByVal TargetRange As String)

    Dim Dict As Object
    Dim c As Range
    Dim Dummy As Range
    Dim OutCell As Range
    Dim Key As Variant

    Set Dict = CreateObject("Scripting.Dictionary")

    '-----------------------------
    ' Screen Area から収集
    '-----------------------------
    For Each c In Ws.Range(REPL_SCREEN_AREA)

        If Trim$(c.Value) <> "" Then

            If Not Dict.Exists(CStr(c.Value)) Then
                Dict.Add CStr(c.Value), True
            End If

        End If

    Next

    '-----------------------------
    ' Dummy Area をクリア
    '-----------------------------
    Set Dummy = Ws.Range(REPL_DUMMY_AREA)

    Dummy.ClearContents

    '-----------------------------
    ' 一覧を書き込み
    '-----------------------------
    Set OutCell = Dummy.Cells(1, 1)

    For Each Key In Dict.Keys

        OutCell.Value = Key

        Set OutCell = OutCell.Offset(0, 1)

        ' 行末なら次の行へ
        If OutCell.Column > Dummy.Columns(Dummy.Columns.Count).Column Then

            Set OutCell = Dummy.Cells(OutCell.Row - Dummy.Row + 2, 1)

        End If

        ' Dummy Area を超えたら終了
        If OutCell.Row > Dummy.Rows(Dummy.Rows.Count).Row Then Exit For

    Next

End Sub
