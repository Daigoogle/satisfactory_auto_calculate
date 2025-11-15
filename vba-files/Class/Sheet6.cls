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

Function GetItemMenuList() As ItemMenu()
    Dim i As Byte
    Dim v_in As Variant
    Dim MenuList() As Variant
    
    ' A2:M200の範囲をバリアント配列に取得
    v_in = Range("A2:M200")
    
    ' 空白行までのカウント（最大199行）
    For i = 1 To 199
        If v_in(i, 1) = "" Then
            Exit For
        End If
    Next i
    
    ' 有効なデータ行数に合わせて配列をリサイズ
    ReDim MenuList(1 To i - 1, 1 To 13)
    
    ' データを新しい配列にコピー（.Value は不要）
    For i = 1 To UBound(MenuList, 1)
        MenuList(i).ItemNo = v_in(i, 1)
        MenuList(i).Base1No = v_in(i, 2)
        MenuList(i).Base2No = v_in(i, 3)
        MenuList(i).Base3No = v_in(i, 4)
        MenuList(i).Base4No = v_in(i, 5)
        MenuList(i).SidelineNo = v_in(i, 6)
        MenuList(i).Base1Num = v_in(i, 7)
        MenuList(i).Base2Num = v_in(i, 8)
        MenuList(i).Base3Num = v_in(i, 9)
        MenuList(i).Base4Num = v_in(i, 10)
        MenuList(i).SidelineNum = v_in(i, 11)
        MenuList(i).OutputNum = v_in(i, 12)
        MenuList(i).Rate = v_in(i, 13)
    Next i
    
    GetItemMenuList = MenuList()
End Function


