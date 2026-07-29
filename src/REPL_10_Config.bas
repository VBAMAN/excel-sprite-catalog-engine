Attribute VB_Name = "REPL_10_Config"
Option Explicit

'=========================================================
' REPL_10_Config
'
' Replace Engine Configuration
'
' 指定範囲の文字列をスプライトコードへ変換し、
' スプライトエンジンへ渡すための設定
'=========================================================

'---------------------------------------------------------
' Replace Working Area
'
' ※ SPRTエンジンと重なる範囲を設定する
'---------------------------------------------------------

Public Const REPL_DUMMY_AREA  As String = "AX5:BE12"
Public Const REPL_SCREEN_AREA As String = "G5:AT34"

'---------------------------------------------------------
' Replace Prefix
'---------------------------------------------------------

Public Const REPL_PREFIX_NUMBER As String = "0"   ' 0～9 → 00～09
Public Const REPL_PREFIX_ALPHA  As String = "0"   ' A～Z → 0A～0Z

'---------------------------------------------------------
' Replace Special Code
'---------------------------------------------------------

Public Const REPL_CODE_SPACE   As String = ""     ' 空白
Public Const REPL_CODE_MINUS   As String = "0M"   ' -
Public Const REPL_CODE_SPARE   As String = "0S"   ' /
Public Const REPL_CODE_STRIKE  As String = "0X"   ' X

'---------------------------------------------------------
' Catalog
'---------------------------------------------------------

Public Const REPL_MAX_CODE_LEN As Long = 2
