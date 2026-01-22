Attribute VB_Name = "Module_HalfProductLogic"
Option Explicit

Sub 半製品_工程判定(工程名 As String)

    ' まず全部いったんリセット
    Range("B24").Locked = True   ' 枚数
    Range("B24").Interior.ColorIndex = 0

    Range("B25").Locked = False  ' ロット（基本は任意、あとで工程ごとに制御）
    Range("B25").Interior.ColorIndex = 0

    Range("B29").Locked = False  ' 平均重量
    Range("B29").Interior.ColorIndex = 0

    Range("C23").Value = ""      ' 単位クリア

    ' ▼ 単位設定
    Select Case 工程名
        Case "成型", "LS撹拌", "検品"
            Range("C23").Value = "㎏"
        Case "プリント", "個包装", "メレンゲデコ"
            Range("C23").Value = "粒"
        Case "包装糖", "カット・検品"
            Range("C23").Value = "玉"
        Case "LS個包装"
            Range("C23").Value = "個"
        Case "箱折"
            Range("C23").Value = "箱"
        Case "パッキン裁断", "緩衝材裁断"
            Range("C23").Value = "枚"
        Case Else
            ' 想定外工程は空欄のまま
    End Select

    ' ▼ 枚数必須：成型のみ
    If 工程名 = "成型" Then
        Range("B24").Locked = False
        Range("B24").Interior.Color = RGB(255, 255, 150) ' 濃いめの黄色＝必須
    End If

    ' ▼ ロット不要：成型・LS撹拌・包装糖
    If 工程名 = "成型" Or 工程名 = "LS撹拌" Or 工程名 = "包装糖" Then
        Range("B25").Locked = True
        Range("B25").Value = ""
        Range("B25").Interior.ColorIndex = 0
    Else
        ' それ以外は必須
        Range("B25").Locked = False
        Range("B25").Interior.Color = RGB(255, 255, 150)
    End If

    ' ▼ 平均重量必須：検品・プリント・個包装 かつ TＤＳの「検品時計量」＝1
    Dim 必須 As Boolean
    必須 = False

    If 工程名 = "検品" Or 工程名 = "プリント" Or 工程名 = "個包装" Then
        If 半製品_検品時計量フラグ(Range("B9").Value) = 1 Then
            必須 = True
        End If
    End If

    If 必須 Then
        Range("B29").Locked = False
        Range("B29").Interior.Color = RGB(255, 255, 150)
    Else
        Range("B29").Locked = False
        Range("B29").Interior.ColorIndex = 0
    End If

End Sub
Function 半製品_検品時計量フラグ(半製品名 As String) As Long

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Range

    Set ws = Sheets("半製品マスタ")   ' TＤＳを読み込んでいるシート
    Set tbl = ws.ListObjects(1)       ' テーブル名が別なら修正

    半製品_検品時計量フラグ = 0

    For Each r In tbl.ListColumns("種類").DataBodyRange
        If r.Value = 半製品名 Then
            ' 「検品時計量」列（全角名）を参照
            半製品_検品時計量フラグ = r.Offset(0, tbl.ListColumns("検品時計量").Index - r.Column).Value
            Exit Function
        End If
    Next r

End Function
