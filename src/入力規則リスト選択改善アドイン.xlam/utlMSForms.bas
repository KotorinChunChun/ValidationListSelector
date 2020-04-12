Attribute VB_Name = "utlMSForms"
Rem
Rem @module     UtlMSForms
Rem
Rem @description MSForms�̃C�P�ĂȂ��R���g���[�����A�C�C�����Ɏg�����߂̊֐��Q
Rem              ����؂�o��������
Rem
Option Explicit

'�ϐ��̔z��(For Each�ł���悤�ɂ���)
Function ToArray(v)
    ToArray = v
    If IsArray(v) Then Exit Function
    ToArray = VBA.Array(v)
End Function

'���X�g�{�b�N�X�ɕ\������A�C�e�����t�B���^����
Public Sub ListBox_SetItems(lb As MSForms.ListBox, _
                            ListItems, _
                            likeFilterText As String, _
                            isAllMatch As Boolean)
                            
    If likeFilterText = "" Then likeFilterText = "*"
    
    Dim headString As String, footString As String
    If Not isAllMatch Then
        headString = "*"
        footString = "*"
    End If
    
    '���X�g�{�b�N�X�̑I���ς݃A�C�e���ێ�
    Dim selectedItems
    selectedItems = ListBox_GetSelectedItems(lb)
    
    On Error GoTo LabelError
    '���X�g�{�b�N�X�ɃA�C�e���ǉ�
    Dim item
    lb.Clear
    For Each item In ListItems
        If item Like headString & likeFilterText & footString Then
            lb.AddItem item
        End If
    Next
    On Error GoTo 0
    
    '���X�g�{�b�N�X�̑I���ς݃A�C�e������
    ListBox_SetSelectedItems lb, selectedItems
    
    lb.Enabled = True
    Exit Sub
    
    '���̓t�H�[�}�b�g�� Like���Z�q�̃��[���ɓK�����Ȃ������ꍇ
LabelError:
    lb.Clear
    lb.AddItem "ERROR"
    lb.Enabled = False
End Sub

'���X�g�{�b�N�X�̑I���A�C�e����z��Ŏ擾
'��������F���Ȃ̂ŏd���A�C�e���͖������ɑS�Ď擾���܂��B
'��I����:�v�f0�̔z��
Public Function ListBox_GetSelectedItems(lb As MSForms.ListBox) As Variant
    ListBox_GetSelectedItems = VBA.Array()
    If lb.ListCount = 0 Then Exit Function
    
    Dim arr
    ReDim arr(0 To lb.ListCount - 1)
    Dim i As Long, nextIndex As Long
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            arr(nextIndex) = lb.list(i)
            nextIndex = nextIndex + 1
        End If
    Next
    
    If nextIndex = 0 Then ListBox_GetSelectedItems = VBA.Array(): Exit Function
    ReDim Preserve arr(0 To nextIndex - 1)
    
    ListBox_GetSelectedItems = arr
End Function

'���X�g�{�b�N�X�̎w�肵���A�C�e����I����Ԃɂ���
'��������F���Ȃ̂ŏd���A�C�e���͖������ɑS�đI�����܂��B
'����؂蕶���̓J���}�ł��B�A�C�e���ɃJ���}���܂ނƃo�O��܂��B
Public Function ListBox_SetSelectedItems(lb As MSForms.ListBox, arr_select_items) As Variant
    If Not IsArray(arr_select_items) Then Exit Function
    
    Dim i As Long
    Dim item
    Dim isMatched As Boolean
    
    For i = 0 To lb.ListCount - 1
        isMatched = False
        For Each item In ToArray(arr_select_items)
            isMatched = isMatched Or (lb.list(i) = item)
'            If isMatched Then Stop
        Next
        lb.Selected(i) = isMatched
    Next
    
End Function
