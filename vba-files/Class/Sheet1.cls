VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function GetMaterialList() As Material()
    Dim i As Byte
    Dim v_in As Variant
    Dim out() As Material
    
    ' C3:D11の範囲をバリアント配列に取得
    v_in = Range("C3:D11")
    
    ' 空白行までのカウント（最大9行）
    For i = 1 To 9
        ' VBA では Or を使用（|| は不可）
        If v_in(i, 1) = "" Or v_in(i, 2) = "" Then
            Exit For
        End If
    Next i
    
    ' 有効なデータ行数に合わせて配列をリサイズ
    ReDim out(1 To i - 1)
    
    ' データを新しい配列にコピー（.Value は不要）
    For i = 1 To UBound(out, 1)
        out(i).Name = v_in(i, 1)
        out(i).Num = v_in(i, 2)
    Next i
    
    GetMaterialList = out()
End Function
