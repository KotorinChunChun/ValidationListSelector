Attribute VB_Name = "AppMain"
Rem
Rem @appname ValidationListSelector - ���͋K�����X�g�I�����P�A�h�C��
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/04/12 : �����
Rem    2020/04/   :
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "���͋K�����X�g�I�����P�A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.10"
Public Const APP_UPDATE = "2020/04/12"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/20200412-validation-list-selector-addin"

Public instValidationListSelector As ValidationListSelector

'--------------------------------------------------
'�A�h�C�����s��
Sub AddinStart()
    MsgBox "���͋K�����X�g���g���₷�����܂��I�I�I" & vbLf & _
            "" & vbLf & _
            "", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

'�A�h�C���ꎞ��~��
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "�Ď����~���܂����H", _
        "�ق�Ƃɂ�߂��Ⴄ�́H")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "���肪�Ɓ`�`�`", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "�܂�������łˁH", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

'�A�h�C���ݒ�\��
Sub AddinConfig(): Call SettingForm.Show: End Sub

'�A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'�A�h�C�����S�I��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
'--------------------------------------------------

'�Ď��J�n
'Workbook_Open����Ă΂��
Sub MonitorStart(): Set instValidationListSelector = New ValidationListSelector: End Sub

'�Ď���~
Sub MonitorStop(): Set instValidationListSelector = Nothing: End Sub
