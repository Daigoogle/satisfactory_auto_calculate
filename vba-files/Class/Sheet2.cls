VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim recipieList() As ItemMenu
Dim itemName() As String ' アイテム名

' セル内容リセット
Function CellsReset()
    Range("A4:U3000").ClearContents
End Function

' 工程計算用関数
Function BackToTheOrigin(sozaiID As Integer) As Integer
    Dim i As Integer, ansSum As Integer
    ansSum = 0
    For i = 1 To 4
        If recipieList(sozaiID).BaseNo(i) = 0 Then
            Exit For
        End If
        ' 再帰呼び出し
        ansSum = ansSum + 1
        ansSum = ansSum + BackToTheOrigin(recipieList(sozaiID).BaseNo(i))
    Next i
    BackToTheOrigin = ansSum
End Function

' 工程１ステップ算出関数
Function OneStep(wantMat As MaterialPack, structurePosition As LongLong) As Variant
    Dim i As Byte
    Dim magnification As Single
    Dim out(1 To 21) As Variant

    magnification = RoundUp(wantMat.Num / recipieList(wantMat.NameID).OutputNum, 3)
    
    out(1) = structurePosition ' 工程位置
    out(2) = wantMat.NameID ' 製品名（一旦ID）
    out(3) = wantMat.Num ' 要求生産量
    out(14) = recipieList(wantMat.NameID).OutputNum ' 1機械1回あたりの生産数
    out(20) = recipieList(wantMat.NameID).Rate ' レート
    For i = 1 To 4
        If recipieList(wantMat.NameID).BaseNo(i) > 0 Then
            out(2 + i * 2) = recipieList(wantMat.NameID).BaseNo(i) ' 素材名（一旦ID）
            out(3 + i * 2) = recipieList(wantMat.NameID).BaseNum(i) * magnification  ' 素材必要数
            out(14 + i) = recipieList(wantMat.NameID).BaseNum(i) ' 素材単価
        End If
    Next i
    If recipieList(wantMat.NameID).SidelineNo > 0 Then
        out(12) = recipieList(wantMat.NameID).SidelineNo ' 副産物名（一旦ID）
        out(13) = recipieList(wantMat.NameID).SidelineNum * magnification ' 副産物
    End If
    OneStep = out
End Function

' 副産物種類特定
Function GetNonZeroUnique(arr() As Variant) As Variant
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i)(12) <> 0 Then
            If Not dic.Exists(arr(i)(12)) Then
                dic.Add arr(i)(12), arr(i)(12)
            End If
        End If
    Next i

    GetNonZeroUnique = dic.Items
End Function

' 工程＆副産物出力メイン関数
Public Sub Update_Cells()
    Dim i As Integer, j As Integer, k As Integer, nRows As Integer
    Dim structurePosition As LongLong
    Dim sidelineFlag As Boolean
    Dim parts As MaterialPack
    Dim sideline As Variant
    Dim wholeProcess As Variant ' 工程全体格納用2次元配列

'初期化
    Sheet1.CellsReset
    Sheet2.CellsReset
    nRows = i = 0
    j = 1

' シート１より要求素材取得
    Dim wantMat() As MaterialPack
    wantMat = Sheet1.GetMaterialList()

' シート６よりレシピ一覧（ID）を取得
    recipieList = Sheet6.GetItemMenuList()

' シート４より名前リストを取得
    itemName = Sheet4.GetNameList()

' 工程＆必要素材算出
    ' 工程出力に必要な行数を算出
    nRows = UBound(wantMat) - LBound(wantMat) + 1
    For i = LBound(wantMat) To UBound(wantMat)
        nRows = nRows + BackToTheOrigin(wantMat(i).NameID)
    Next i
    ' 出力用配列をリサイズ
    ReDim wholeProcess(1 To nRows)
    ' 要求素材をセット
    For i = 1 To UBound(wantMat)
        wholeProcess(i) = OneStep(wantMat(i), CLngLng(i))
    Next i
    ' 再帰的に素材工程をセット
    Do While i > j ' i:記入行 j:参照行
        For k = 1 To 4
            If wholeProcess(j)(k * 2 + 2) = 0 Then
                Exit For
            End If
            parts.NameID = wholeProcess(j)(k * 2 + 2)
            parts.Num = wholeProcess(j)(k * 2 + 3)
            wholeProcess(i) = OneStep(parts, (wholeProcess(j)(1) * 10) + CLngLng(k))
            i = i + 1
        Next k
        j = j + 1
    Loop

' 副産物集計
    Dim sidelinelist As Variant
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    For i = LBound(wholeProcess) To UBound(wholeProcess)
        If wholeProcess(i)(12) <> 0 Then
            If Not dic.Exists(wholeProcess(i)(12)) Then
                dic.Add wholeProcess(i)(12), wholeProcess(i)(12)
                sidelineFlag = True
            End If
        End If
    Next i
    
    If sidelineFlag Then
        sidelinelist = dic.Items
        ReDim sideline(1 To UBound(sidelinelist) + 1, 2)
        For i = 1 To UBound(sidelinelist) + 1
            sideline(i, 1) = sidelinelist(i - 1)
            sideline(i, 2) = 0
            For j = LBound(wholeProcess) To UBound(wholeProcess)
                If wholeProcess(j)(12) = sidelinelist(i - 1) Then
                    sideline(i, 2) = sideline(i, 2) + wholeProcess(j)(13)
                End If
            Next j
        Next i
    End If

' 名前のIDを文字列に変換
    For i = 1 To nRows
        wholeProcess(i)(2) = itemName(wholeProcess(i)(2))
        For k = 1 To 4
            If wholeProcess(i)(k * 2 + 2) = "" Then
                Exit For
            End If
            wholeProcess(i)(k * 2 + 2) = itemName(wholeProcess(i)(k * 2 + 2))
        Next k
        If wholeProcess(i)(12) <> "" Then
            wholeProcess(i)(12) = itemName(wholeProcess(i)(12))
        End If
    Next i

' 工程＆副産物出力
    If sidelineFlag Then
        Sheet1.PutSidelines (sideline)
    End If
    Dim putsdatas As Variant
    ReDim putsdatas(1 To nRows, 1 To 20)
    For i = 1 To nRows
        For j = 1 To 20
            putsdatas(i, j) = wholeProcess(i)(j)
        Next j
    Next i
    
    Range("A4").Resize(UBound(putsdatas, 1), UBound(putsdatas, 2)) = putsdatas
    
End Sub

