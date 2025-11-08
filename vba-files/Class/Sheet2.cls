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

Public Sub Update_Cells()
    Dim i As Byte, j As Integer, nRows As Integer
    Dim SidelineNames() As String, SidelineNum() As Single ' 余剰素材・副産物
    Dim BaseElemNames() As String, BaseElemNum() As Single ' 必要基底素材
    Dim WantPartNames(9) As String, WantPartNum(9) As Single  ' 要求素材
    i = j = 0
    '初期化
    
    Range("A4:A1000").Value = ""
    
    'シート１に書いた素材の入力
    'Sheets ("合計必要素材").Range()
    For i = 0 To 9
        WantParts(i) = Sheets("合計必要素材").Cells(i + 3, 3)
        If WantParts(i) <> "" Then
            'MsgBox (WantParts(i))
            Cells(4 + j, 1) = j + 1
            Cells(4 + j, 2) = Sheets("合計必要素材").Cells(i + 3, 4)
            Cells(4 + j, 3) = WantParts(i)
            j = j + 1
        End If
    Next i
    nRows = 0
    '各種必要素材を記述していく！
    Do While j > nRows
        For i = 1 To 4
            If Cells(nRows + 4, i + 3).Value = "" Then
                Exit For
            End If
            Cells(4 + j, 1).Value = Cells(nRows + 4, 1) * 10 + i
            j = j + 1
        Next i
        nRows = nRows + 1
    Loop
    
End Sub

Function IsBaseMaterial() As Boolean
    
End Function

