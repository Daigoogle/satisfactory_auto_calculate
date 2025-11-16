Attribute VB_Name = "Module1"
Option Explicit

' アイテムメニューの型定義
Type ItemMenu
    ItemNo As Integer
    BaseNo(1 To 4) As Integer
    SidelineNo As Integer
    BaseNum(1 To 4) As Integer
    SidelineNum As Integer
    OutputNum As Integer
    Rate As Single
End Type

' 素材名と製造数のセット
Type MaterialPack
    NameID As Integer
    Num As Single
End Type

Function HideIf(value As Variant, condition As String, Optional hideReturn As Variant = "") As Variant
    Dim condGroup() As String
    Dim cond As Variant
    Dim evalResult As Boolean
    Dim part As Variant
    
    '--------------------------------------
    ' エラー判定
    '--------------------------------------
    If UCase(condition) = "ISERROR" Then
        If IsError(value) Then HideIf = hideReturn Else HideIf = value
        Exit Function
    End If
    
    If UCase(condition) = "ISBLANK" Then
        If Trim(CStr(value)) = "" Then HideIf = hideReturn Else HideIf = value
        Exit Function
    End If

    '--------------------------------------
    ' OR 条件で分割
    ' 例: ">10 OR <0"
    '--------------------------------------
    condGroup = Split(condition, " OR ")
    
    For Each cond In condGroup
        
        '--- AND 条件があるか？ ---
        Dim andParts() As String
        Dim allTrue As Boolean: allTrue = True
        
        andParts = Split(cond, " AND ")
        
        For Each part In andParts
            part = Trim(part)
            
            If Not EvaluateCondition(value, part) Then
                allTrue = False
                Exit For
            End If
        Next part
        
        ' OR 条件のどれかが true なら非表示
        If allTrue Then
            HideIf = hideReturn
            Exit Function
        End If
    Next cond
 
    '--------------------------------------
    ' 条件に一致しなければ値を返す
    '--------------------------------------
    HideIf = value
End Function

'===========================================
' 条件の個別評価ロジック
'===========================================
Private Function EvaluateCondition(value As Variant, cond As Variant) As Boolean
    Dim op As String
    Dim Num As Double

    '--------------------------------------
    ' 正規表現
    '--------------------------------------
    If UCase(Left(cond, 6)) = "REGEX:" Then
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = Mid(cond, 7)
        re.IgnoreCase = True
        re.Global = False
        EvaluateCondition = re.Test(CStr(value))
        Exit Function
    End If
    
    '--------------------------------------
    ' 文字列ワイルドカード一致
    '--------------------------------------
    If InStr(cond, "*") > 0 Or InStr(cond, "?") > 0 Then
        EvaluateCondition = (CStr(value) Like cond)
        Exit Function
    End If
    
    '--------------------------------------
    ' 文字列完全一致
    '--------------------------------------
    If Not IsNumeric(cond) And _
       (Left(cond, 1) <> "<" And Left(cond, 1) <> ">" And Left(cond, 1) <> "=") Then
        EvaluateCondition = (CStr(value) = cond)
        Exit Function
    End If
    
    '--------------------------------------
    ' 比較式解析（> < >= <= = <>）
    '--------------------------------------
    If cond Like "[<>]*" Or cond Like "=[0-9]*" Then
        
        '演算子抽出
        If Left(cond, 2) = ">=" Or Left(cond, 2) = "<=" Or Left(cond, 2) = "<>" Then
            op = Left(cond, 2)
            Num = CDbl(Mid(cond, 3))
        Else
            op = Left(cond, 1)
            Num = CDbl(Mid(cond, 2))
        End If
        
        If Not IsNumeric(value) Then
            EvaluateCondition = False
            Exit Function
        End If
        
        Select Case op
            Case "=":  EvaluateCondition = (CDbl(value) = Num)
            Case ">":  EvaluateCondition = (CDbl(value) > Num)
            Case "<":  EvaluateCondition = (CDbl(value) < Num)
            Case ">=": EvaluateCondition = (CDbl(value) >= Num)
            Case "<=": EvaluateCondition = (CDbl(value) <= Num)
            Case "<>": EvaluateCondition = (CDbl(value) <> Num)
        End Select
        
        Exit Function
    End If
    
    '--------------------------------------
    ' 最後の fallback
    '--------------------------------------
    EvaluateCondition = False
End Function

Function RoundUp(ByVal value As Single, Optional ByVal digits As Long = 0) As Single
    Dim factor As Single
    factor = 10 ^ digits
    
    If value >= 0 Then
        RoundUp = -Int(-value * factor) / factor
    Else
        RoundUp = Int(value * factor) / factor
    End If
End Function
