VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 文字列を検索して行番号を返す関数
Public Function StrToNo(str As Variant) As Integer
    Dim i As Byte
    Dim n As Integer
    Dim v_in As Variant

    ' 初期化
    n = 0
    ' E2:E200の範囲をバリアント配列に取得
    v_in = Range("E2:E200")

    ' 文字列を検索（1から200まで）
    For i = 1 To 200
        If v_in(i, 1) = str Then
            n = i
            Exit For
        End If
    Next i
    
    StrToNo = n
End Function

' 行番号を指定して文字列を返す関数
Public Function NoToStr(n As Integer) As String
    ' E列のn行目の値を取得
    NoToStr = Cells(n + 1, 5).value
End Function

' 名前のリストを渡す関数
Public Function GetNameList() As String()
    Dim i As Byte
    Dim v_in As Variant
    Dim out() As String
    
    ' E2:E200の範囲をバリアント配列に取得
    v_in = Range("E2:E200")
    
    ' 空白行までのカウント（最大199行）
    For i = 1 To 199
        If v_in(i, 1) = "" Then
            Exit For
        End If
    Next i
    ' 有効なデータ行数に合わせて配列をリサイズ
    ReDim out(1 To i - 1)
    
    ' データを新しい配列にコピー（.Value は不要）
    For i = 1 To UBound(out, 1)
        out(i) = v_in(i, 1)
    Next i
    
    GetNameList = out()
End Function
