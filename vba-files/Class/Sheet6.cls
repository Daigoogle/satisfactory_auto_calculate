VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function GetItemMenuList() As ItemMenu()
    Dim i As Byte, j As Byte
    Dim v_in As Variant
    Dim MenuList() As ItemMenu
    
    ' A2:M200の範囲をバリアント配列に取得
    v_in = Range("A2:M200")
    
    ' 空白行までのカウント（最大199行）
    For i = 1 To 199
        If v_in(i, 1) = "" Then
            Exit For
        End If
        ' 素材/副産物がnullの場合0を入力
        For j = 2 To 11
            If v_in(i, j) = "" Then
                v_in(i, j) = 0
            End If
        Next j
    Next i
    
    ' 有効なデータ行数に合わせて配列をリサイズ
    ReDim MenuList(1 To i - 1)
    
    
    
    ' データを新しい配列にコピー（.Value は不要）
    For i = 1 To UBound(MenuList, 1)
        MenuList(i).ItemNo = v_in(i, 1)
        For j = 1 To 4
            MenuList(i).BaseNo(j) = v_in(i, j + 1)
        Next j
        MenuList(i).SidelineNo = v_in(i, 6)
        For j = 1 To 4
            MenuList(i).BaseNum(j) = v_in(i, j + 6)
        Next j
        MenuList(i).SidelineNum = v_in(i, 11)
        MenuList(i).OutputNum = v_in(i, 12)
        MenuList(i).Rate = v_in(i, 13)
    Next i
        
    GetItemMenuList = MenuList()
End Function
