Attribute VB_Name = "test"
Option Explicit

Sub ���͋K�����烊�X�g���e���擾����1()
    Dim Target As Range
    Set Target = Selection
    
    Dim v As Validation
    Set v = Target.Validation
    
    If Intersect(Target, Target.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then Exit Sub
    
    If v.Type <> XlDVType.xlValidateList Then Exit Sub
    
    Dim list As Variant
    list = Application.Evaluate(v.Formula1)
End Sub

Sub ���͋K���̓������Z���S���I��()
    Cells.SpecialCells(xlCellTypeAllValidation).Select
End Sub

Sub Test_NG_����Show���Ă͂Ȃ�Ȃ�()
    ListSelectorForm.Show
End Sub

Sub Test_OK_OpenForm���g�p����()
    Debug.Print Join(ListSelectorForm.OpenForm(Split("a,b,c", ","), "b", False), ",")
End Sub

Sub Test_OK_�}���`�Z���N�g()
    Debug.Print Join(ListSelectorForm.OpenForm(Split("a,b,c", ","), "b", True), ",")
End Sub

'���[�h���X�i�t�H�[�����풓������j���߂ɗp�ӂ��Ă��邪
'���[�_���i�V�[�g����s�\��ԁj�ł͂��̏����������郁���b�g�͂Ȃ�
Sub Test_OK_�C���X�^���X���s�ƒx���l�擾()
    Dim fm As ListSelectorForm
    Set fm = New ListSelectorForm
    
    Call fm.OpenForm(Split("a,b,c", ","), "b", True)
    
    'Alt+F4�Ȃǂ�Unload���ꂽ���A�G���[���������ɂȂ�B
    Dim dummy
    
    On Error Resume Next
    dummy = Null
    dummy = fm.Result
    On Error GoTo 0
    
    If IsNull(dummy) Then
        Debug.Print "�L�����Z������܂����B"
    Else
        Debug.Print Join(dummy, ",") & "���I������܂���"
    End If
End Sub


Sub Macro5()
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="��,��,��,��"
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

