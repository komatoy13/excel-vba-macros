Attribute VB_Name = "Module_Time"
Option Explicit

Function 半製品_標準時間取得(工程名 As String, 半製品名 As String) As Double

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Range

    Set ws = Sheets("半製品マスタ")   ' TＤＳを読み込んでいるシート
    Set tbl = ws.ListObjects(1)        ' テーブル名が違う場合は修正

    半製品_標準時間取得 = 0

    For Each r In tbl.ListColumns("種類").DataBodyRange
        If r.Value = 半製品名 Then

            ' 検品系は「検品速度」
            If 工程名 = "検品" Or 工程名 = "カット・検品" Then
                半製品_標準時間取得 = r.Offset(0, tbl.ListColumns("検品速度").Index - r.Column).Value
            Else
                半製品_標準時間取得 = r.Offset(0, tbl.ListColumns("成型速度").Index - r.Column).Value
            End If

            Exit Function
        End If
    Next r

End Function
Sub 主作業時間計算()

    Dim 種類 As String
    Dim 工程 As String
    Dim 名称 As String
    Dim 数量 As Double
    Dim 標準時間 As Double
    Dim 理論時間 As Double
    Dim 実績時間 As Double
    Dim 差分 As Double

    種類 = Range("B3").Value
    工程 = Range("B21").Value
    名称 = Range("B9").Value
    数量 = Val(Range("B23").Value)
    実績時間 = Val(Range("B27").Value)

    ' ▼ 商品は対象外
    If 種類 <> "半製品" Then Exit Sub
    If 名称 = "" Or 工程 = "" Then Exit Sub

    ' ▼ 標準時間取得
    標準時間 = 半製品_標準時間取得(工程, 名称)
    If 標準時間 = 0 Then Exit Sub

    ' ▼ 理論時間
    理論時間 = 標準時間 * 数量

    ' ▼ 差分
    差分 = 実績時間 - 理論時間
    Range("C27").Value = 差分   ' 内部保持

    ' ▼ 異常値チェック（赤色表示）
    Range("B27").Interior.ColorIndex = 0

    If 理論時間 > 0 Then

        ' 遅すぎる（2倍以上）
        If 実績時間 >= 理論時間 * 2 Then
            Range("B27").Interior.color = RGB(255, 150, 150)
        End If

        ' 早すぎる（1/3以下）
        If 実績時間 <= 理論時間 / 3 Then
            Range("B27").Interior.color = RGB(255, 150, 150)
        End If

    End If

End Sub
