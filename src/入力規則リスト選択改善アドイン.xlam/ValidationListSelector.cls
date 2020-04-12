VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationListSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem @module
Rem   ValidationListSelector
Rem
Rem @description
Rem   入力規則の入ったセルをダブルクリックしたとき ListSelectorForm を表示する
Rem
Rem @note
Rem　 インスタンスを生成するだけで監視します。
Rem
Rem @author
Rem   @KotorinChunChun
Rem
Rem @update
Rem   2020/04/12
Rem
Option Explicit

Public WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Private Sub app_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Call OpenValidationList(Sh, Target, Cancel)
End Sub

'Private Sub app_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'    Call OpenValidationList(Sh, Target, Cancel)
'End Sub
'
'Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'    Debug.Print Sh.Name, Target.Address, Target.Parent.Parent.Name
'    Call OpenValidationList(Sh, Target, False)
'End Sub

Sub OpenValidationList(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)

    Set Target = Target(1, 1)
    
    'エラー対策
    On Error Resume Next
    If Target.Validation.Type <> XlDVType.xlValidateList Then Exit Sub

    Dim Formula: Formula = Target.Validation.Formula1
    If Formula = "" Then Exit Sub
    On Error GoTo 0
    
    '入力リストの選択項目を取得
    Dim list As Variant
    If Formula Like "=*" Then
        list = Application.Evaluate(Formula)
        list = WorksheetFunction.Transpose(list)
    Else
        list = Split(Formula, ",")
    End If
    
    '初期状態で選択する値
    Dim def As Variant
    def = Split(Target.Value, ",")
    
    'フォームの実行と結果の取得
    Dim Result
    Result = ListSelectorForm.OpenForm(list, def, False)
    
    'キャンセルされたのでなければ、結果をコンマ区切りで連結してセルに入力
    If Not IsNull(Result) Then
        Target.Value = Join(Result, ",")
    End If
    
    Cancel = True
End Sub

'----------------------------------------
'コンストラクタ
Private Sub Class_Initialize()
    Set app = Application
End Sub
