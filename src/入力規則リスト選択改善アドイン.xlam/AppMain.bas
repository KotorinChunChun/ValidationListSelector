Attribute VB_Name = "AppMain"
Rem
Rem @appname ValidationListSelector - 入力規則リスト選択改善アドイン
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/04/12 : 初回版
Rem    2020/04/   :
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "入力規則リスト選択改善アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.10"
Public Const APP_UPDATE = "2020/04/12"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/20200412-validation-list-selector-addin"

Public instValidationListSelector As ValidationListSelector

'--------------------------------------------------
'アドイン実行時
Sub AddinStart()
    MsgBox "入力規則リストを使いやすくします！！！" & vbLf & _
            "" & vbLf & _
            "", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

'アドイン一時停止時
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "監視を停止しますか？", _
        "ほんとにやめちゃうの？")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "ありがと～～～", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "またあそんでね？", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

'アドイン設定表示
Sub AddinConfig(): Call SettingForm.Show: End Sub

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'アドイン完全終了
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
'--------------------------------------------------

'監視開始
'Workbook_Openから呼ばれる
Sub MonitorStart(): Set instValidationListSelector = New ValidationListSelector: End Sub

'監視停止
Sub MonitorStop(): Set instValidationListSelector = Nothing: End Sub
