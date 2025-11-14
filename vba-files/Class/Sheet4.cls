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

Function HideIf(value As Variant, condition As Variant) As Variant
    Dim cond As String
    Dim num As Double
    Dim op As String
    Dim testString As String

    '条件を文字列化
    cond = CStr(condition)
    
    '------------------------------------
    ' 1. 数値の直接比較（例: 0, 10, -1）
    '------------------------------------
    If IsNumeric(cond) Then
        If IsNumeric(value) Then
            If CDbl(value) = CDbl(cond) Then
                HideIf = ""
                Exit Function
            End If
        End If
        
        '一致しない → 値を返して終了
        HideIf = value
        Exit Function
    End If
    
    '------------------------------------
    ' 2. 比較式（<=0, >=10 等）
    '------------------------------------
    'パターン判定: >, <, >=, <=
    If cond Like "[<>]*" Then
        '演算子判別
        If Left(cond, 2) = ">=" Or Left(cond, 2) = "<=" Then
            op = Left(cond, 2)
            num = CDbl(Mid(cond, 3))
        Else
            op = Left(cond, 1)
            num = CDbl(Mid(cond, 2))
        End If
        
        If IsNumeric(value) Then
            Select Case op
                Case "="
                    If CDbl(value) = num Then HideIf = "" Else HideIf = value
                Case ">"
                    If CDbl(value) > num Then HideIf = "" Else HideIf = value
                Case "<"
                    If CDbl(value) < num Then HideIf = "" Else HideIf = value
                Case ">="
                    If CDbl(value) >= num Then HideIf = "" Else HideIf = value
                Case "<="
                    If CDbl(value) <= num Then HideIf = "" Else HideIf = value
            End Select
            Exit Function
        End If
    End If

    '------------------------------------
    ' 3. 文字列一致 or ワイルドカード一致
    '------------------------------------
    testString = CStr(value)
    
    '完全一致
    If testString = cond Then
        HideIf = ""
        Exit Function
    End If
    
    'ワイルドカード一致 (*abc*, abc?, etc)
    If testString Like cond Then
        HideIf = ""
        Exit Function
    End If

    '非該当 → そのまま返す
    HideIf = value
End Function
