VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListSelectorForm 
   Caption         =   "リスト - "
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "ListSelectorForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
Rem   インクリメンタルサーチ対応のマルチセレクト対応リスト選択フォーム
Rem
Rem @note
Rem　 必ずOpenFormメソッドから起動してください。Showメソッドはダメです。
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
    Me.Caption = "開発者へ：OpenFormメソッドを使って起動してください"
    Me.Hide
End Sub

Rem 引数を指定してフォームを表示する
Rem
Rem @param arr_listitems        リストに表示する全アイテム配列
Rem @param arr_defaultvalues    起動時に選択状態にするアイテム配列
Rem @parem can_multiselect      複数選択の可否
Rem
Rem @return As Variant(0 to n)  選択されていたアイテム配列
Rem                             キャンセル時：Null
Public Function OpenForm(arr_listitems, arr_defaultvalues, can_multiselect As Boolean) As Variant
    ListItems = arr_listitems
    If Not IsArray(arr_listitems) Then arr_listitems = ToArray(arr_listitems)
    If Not IsArray(arr_defaultvalues) Then arr_defaultvalues = ToArray(arr_defaultvalues)
    
    '再利用された場合のため初期化
    IsCanceled = False
    TextBox1.Text = ""
    
    'リストボックス選択モード
    ListBox1.MultiSelect = IIf(can_multiselect, fmMultiSelectMulti, fmMultiSelectSingle)
    
    'リストボックスアイテム登録
    Call Update_listbox1
    
    'リストボックス初期選択アイテム
    If Not IsMissing(arr_defaultvalues) Then
        ListBox_SetSelectedItems ListBox1, arr_defaultvalues
    End If
    
    Application.Wait [Now() + "00:00:00.2"]
    ListBox1.Enabled = True
    Me.Caption = "リスト選択"
    Me.Show '←モーダルフォームではここでVBAが止まる
    
    'フォームが終了されたとき/Unloadされたときエラーが出る
    On Error Resume Next
    OpenForm = Null
    OpenForm = Me.Result
    On Error GoTo 0
End Function

Rem 選択結果の取得
Rem
Rem @return As Variant(0 to n)  選択されていたアイテム配列
Rem                             キャンセル時：Null
Rem @note
Rem  モードレス対応のために公開しているが、
Rem  呼び出し元でUnload対策の処理が複雑になるため
Rem  できる限りOpenForm関数の戻り値を使うべき
Public Property Get Result() As Variant
    'キャンセルされたとき
    If IsCanceled Then Result = Null: Exit Property
    
    Result = ListBox_GetSelectedItems(ListBox1)
End Property

'リストボックス内容更新
Sub Update_listbox1()
    Call ListBox_SetItems(ListBox1, ListItems, TextBox1.Text, OptionButton2.Value)
End Sub

Sub CloseForm(success As Boolean)
    IsCanceled = Not success
    Me.Hide
End Sub

'--------------------------------------------------

'OKボタン実装してないけどね
Private Sub Btn_OK_Click()
    Call CloseForm(True)
End Sub

'キャンセルボタン実装してないけどね
Private Sub Btn_Cancel_Click()
    Call CloseForm(False)
End Sub

'リストダブルクリックが確定の合図
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

