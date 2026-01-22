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

Private Sub Worksheet_Change(ByVal Target As Range)

    ' ▼ 作業区分（商品/半製品）変更時：B3
    If Not Intersect(Target, Range("B3")) Is Nothing Then
        If Range("B3").Value = "商品" Then
            Call 商品モード初期化
        ElseIf Range("B3").Value = "半製品" Then
            Call 半製品モード初期化
        End If
    End If

    ' ▼ 半製品工程プルダウン変更時：B21
    If Not Intersect(Target, Range("B21")) Is Nothing Then
        If Range("B3").Value = "半製品" Then
            Call 半製品_工程判定(Range("B21").Value)
        End If
    End If

    ' ▼ 数量（B23）・主作業時間（B27）変更時：主作業時間ロジック
    If Not Intersect(Target, Range("B23,B27")) Is Nothing Then
        Call 主作業時間計算
    End If

    ' ▼ 数量（B23）・ロット（B25）変更時：歩留まり計算
    If Not Intersect(Target, Range("B23,B25")) Is Nothing Then
        Call 歩留まり計算
    End If

End Sub
