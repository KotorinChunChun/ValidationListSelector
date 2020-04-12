VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListSelectorForm 
   Caption         =   "���X�g - "
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "ListSelectorForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ListSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem
Rem @module
Rem   ListSelectorForm
Rem
Rem @description
Rem   �C���N�������^���T�[�`�Ή��̃}���`�Z���N�g�Ή����X�g�I���t�H�[��
Rem
Rem @note
Rem�@ �K��OpenForm���\�b�h����N�����Ă��������BShow���\�b�h�̓_���ł��B
Rem
Rem @author
Rem   @KotorinChunChun
Rem
Rem @update
Rem   2020/04/12
Rem
Option Explicit

Private ListItems As Variant
Private IsCanceled As Boolean

Private Sub UserForm_Initialize()
    ListBox1.Enabled = False
    Me.Caption = "�J���҂ցFOpenForm���\�b�h���g���ċN�����Ă�������"
    Me.Hide
End Sub

Rem �������w�肵�ăt�H�[����\������
Rem
Rem @param arr_listitems        ���X�g�ɕ\������S�A�C�e���z��
Rem @param arr_defaultvalues    �N�����ɑI����Ԃɂ���A�C�e���z��
Rem @parem can_multiselect      �����I���̉�
Rem
Rem @return As Variant(0 to n)  �I������Ă����A�C�e���z��
Rem                             �L�����Z�����FNull
Public Function OpenForm(arr_listitems, arr_defaultvalues, can_multiselect As Boolean) As Variant
    ListItems = arr_listitems
    If Not IsArray(arr_listitems) Then arr_listitems = ToArray(arr_listitems)
    If Not IsArray(arr_defaultvalues) Then arr_defaultvalues = ToArray(arr_defaultvalues)
    
    '�ė��p���ꂽ�ꍇ�̂��ߏ�����
    IsCanceled = False
    TextBox1.Text = ""
    
    '���X�g�{�b�N�X�I�����[�h
    ListBox1.MultiSelect = IIf(can_multiselect, fmMultiSelectMulti, fmMultiSelectSingle)
    
    '���X�g�{�b�N�X�A�C�e���o�^
    Call Update_listbox1
    
    '���X�g�{�b�N�X�����I���A�C�e��
    If Not IsMissing(arr_defaultvalues) Then
        ListBox_SetSelectedItems ListBox1, arr_defaultvalues
    End If
    
    Application.Wait [Now() + "00:00:00.2"]
    ListBox1.Enabled = True
    Me.Caption = "���X�g�I��"
    Me.Show '�����[�_���t�H�[���ł͂�����VBA���~�܂�
    
    '�t�H�[�����I�����ꂽ�Ƃ�/Unload���ꂽ�Ƃ��G���[���o��
    On Error Resume Next
    OpenForm = Null
    OpenForm = Me.Result
    On Error GoTo 0
End Function

Rem �I�����ʂ̎擾
Rem
Rem @return As Variant(0 to n)  �I������Ă����A�C�e���z��
Rem                             �L�����Z�����FNull
Rem @note
Rem  ���[�h���X�Ή��̂��߂Ɍ��J���Ă��邪�A
Rem  �Ăяo������Unload�΍�̏��������G�ɂȂ邽��
Rem  �ł������OpenForm�֐��̖߂�l���g���ׂ�
Public Property Get Result() As Variant
    '�L�����Z�����ꂽ�Ƃ�
    If IsCanceled Then Result = Null: Exit Property
    
    Result = ListBox_GetSelectedItems(ListBox1)
End Property

'���X�g�{�b�N�X���e�X�V
Sub Update_listbox1()
    Call ListBox_SetItems(ListBox1, ListItems, TextBox1.Text, OptionButton2.Value)
End Sub

Sub CloseForm(success As Boolean)
    IsCanceled = Not success
    Me.Hide
End Sub

'--------------------------------------------------

'OK�{�^���������ĂȂ����ǂ�
Private Sub Btn_OK_Click()
    Call CloseForm(True)
End Sub

'�L�����Z���{�^���������ĂȂ����ǂ�
Private Sub Btn_Cancel_Click()
    Call CloseForm(False)
End Sub

'���X�g�_�u���N���b�N���m��̍��}
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CloseForm(True)
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyEscape Then Call CloseForm(False)
End Sub
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.Value = vbKeyEscape Then Call CloseForm(False)
    If KeyCode.Value = vbKeyReturn Then Call CloseForm(True)
End Sub

Private Sub OptionButton1_Click(): Update_listbox1: End Sub
Private Sub OptionButton2_Click(): Update_listbox1: End Sub
Private Sub TextBox1_Change(): Update_listbox1: End Sub

