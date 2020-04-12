Attribute VB_Name = "test"
Option Explicit

Sub 入力規則からリスト内容を取得する1()
    Dim Target As Range
    Set Target = Selection
    
    Dim v As Validation
    Set v = Target.Validation
    
    If Intersect(Target, Target.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then Exit Sub
    
    If v.Type <> XlDVType.xlValidateList Then Exit Sub
    
    Dim list As Variant
    list = Application.Evaluate(v.Formula1)
End Sub

Sub 入力規則の入ったセル全部選択()
    Cells.SpecialCells(xlCellTypeAllValidation).Select
End Sub

Sub Test_NG_直接Showしてはならない()
    ListSelectorForm.Show
End Sub

Sub Test_OK_OpenFormを使用する()
    Debug.Print Join(ListSelectorForm.OpenForm(Split("a,b,c", ","), "b", False), ",")
End Sub

Sub Test_OK_マルチセレクト()
    Debug.Print Join(ListSelectorForm.OpenForm(Split("a,b,c", ","), "b", True), ",")
End Sub

'モードレス（フォームを常駐させる）ために用意しているが
'モーダル（シート操作不能状態）ではこの書き方をするメリットはない
Sub Test_OK_インスタンス実行と遅延値取得()
    Dim fm As ListSelectorForm
    Set fm = New ListSelectorForm
    
    Call fm.OpenForm(Split("a,b,c", ","), "b", True)
    
    'Alt+F4などでUnloadされた時、エラー処理が厄介になる。
    Dim dummy
    
    On Error Resume Next
    dummy = Null
    dummy = fm.Result
    On Error GoTo 0
    
    If IsNull(dummy) Then
        Debug.Print "キャンセルされました。"
    Else
        Debug.Print Join(dummy, ",") & "が選択されました"
    End If
End Sub


Sub Macro5()
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="あ,い,う,え"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub

