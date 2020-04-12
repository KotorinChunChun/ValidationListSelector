Attribute VB_Name = "utlMSForms"
Rem
Rem @module     UtlMSForms
Rem
Rem @description MSFormsのイケてないコントロールを、イイ感じに使うための関数群
Rem              から切り出したもの
Rem
Option Explicit

'変数の配列化(For Eachできるようにする)
Function ToArray(v)
    ToArray = v
    If IsArray(v) Then Exit Function
    ToArray = VBA.Array(v)
End Function

'リストボックスに表示するアイテムをフィルタする
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
    
    'リストボックスの選択済みアイテム保持
    Dim selectedItems
    selectedItems = ListBox_GetSelectedItems(lb)
    
    On Error GoTo LabelError
    'リストボックスにアイテム追加
    Dim item
    lb.Clear
    For Each item In ListItems
        If item Like headString & likeFilterText & footString Then
            lb.AddItem item
        End If
    Next
    On Error GoTo 0
    
    'リストボックスの選択済みアイテム復元
    ListBox_SetSelectedItems lb, selectedItems
    
    lb.Enabled = True
    Exit Sub
    
    '入力フォーマットが Like演算子のルールに適合しなかった場合
LabelError:
    lb.Clear
    lb.AddItem "ERROR"
    lb.Enabled = False
End Sub

'リストボックスの選択アイテムを配列で取得
'※文字列認識なので重複アイテムは無条件に全て取得します。
'非選択時:要素0の配列
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

'リストボックスの指定したアイテムを選択状態にする
'※文字列認識なので重複アイテムは無条件に全て選択します。
'※区切り文字はカンマです。アイテムにカンマが含むとバグります。
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
