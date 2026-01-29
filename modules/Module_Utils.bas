Attribute VB_Name = "Module_Utils"
Option Explicit

Sub 履歴スライド()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = Sheets("Input")

    ' ▼ 履歴欄の範囲（1件＝F?O列）
    Const START_COL As Long = 6   ' F列
    Const END_COL As Long = 15    ' O列
    Const ROW As Long = 40        ' 履歴表示行（必要なら変更）

    ' ▼ 右にスライド（古い履歴を右へ）
    For i = END_COL To START_COL + 10 Step -10
        ws.Range(ws.Cells(ROW, i), ws.Cells(ROW, i + 9)).Value = _
            ws.Range(ws.Cells(ROW, i - 10), ws.Cells(ROW, i - 1)).Value
    Next i

    ' ▼ 最新データを左端（F列?O列）に挿入
    ws.Cells(ROW, START_COL).Value = Range("B2").Value   ' 作業日
    ws.Cells(ROW, START_COL + 1).Value = IIf(Range("B3").Value = "商品", Range("B8").Value, Range("B9").Value) ' 名称
    ws.Cells(ROW, START_COL + 2).Value = IIf(Range("B3").Value = "商品", 商品工程文字列(), Range("B21").Value) ' 工程
    ws.Cells(ROW, START_COL + 3).Value = Range("B23").Value   ' 数量
    ws.Cells(ROW, START_COL + 4).Value = Range("B27").Value   ' 主作業時間
    ws.Cells(ROW, START_COL + 5).Value = Range("D23").Value   ' 歩留まり
    ws.Cells(ROW, START_COL + 6).Value = Range("B31").Value   ' 歩留り評価
    ws.Cells(ROW, START_COL + 7).Value = Range("B32").Value   ' 時間評価
    ws.Cells(ROW, START_COL + 8).Value = Range("B33").Value   ' 評価理由
    ws.Cells(ROW, START_COL + 9).Value = Now                  ' 登録日時

End Sub


