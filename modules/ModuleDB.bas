Attribute VB_Name = "ModuleDB"
Option Explicit

Sub DB登録()

    Dim wsDB As Worksheet
    Dim nextRow As Long
    Dim 種類 As String

    Set wsDB = Sheets("DB")
    nextRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).ROW + 1

    種類 = Range("B3").Value

    ' ▼ 歩留まり異常チェック（赤色 & 評価に×なし）
    If Range("D23").Interior.color = RGB(255, 150, 150) Then
        Dim 評価 As String
        評価 = Range("B31").Value & Range("B32").Value

        If InStr(評価, "×") = 0 Then
            If MsgBox("歩留まりが標準と大きく異なりますが、評価に×がありません。このまま登録しますか？", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    ' ▼ 種類
    wsDB.Cells(nextRow, 1).Value = 種類

    ' ▼ 作業日
    wsDB.Cells(nextRow, 2).Value = Range("B2").Value

    ' ▼ 名称（商品 or 半製品）
    If 種類 = "商品" Then
        wsDB.Cells(nextRow, 3).Value = Range("B8").Value
    Else
        wsDB.Cells(nextRow, 3).Value = Range("B9").Value
    End If

    ' ▼ 工程（商品は工程文字列、半製品は B21）
    If 種類 = "商品" Then
        wsDB.Cells(nextRow, 4).Value = 商品工程文字列()
    Else
        wsDB.Cells(nextRow, 4).Value = Range("B21").Value
    End If

    ' ▼ 数量
    wsDB.Cells(nextRow, 5).Value = Range("B23").Value

    ' ▼ 単位
    wsDB.Cells(nextRow, 6).Value = Range("C23").Value

    ' ▼ 枚数
    wsDB.Cells(nextRow, 7).Value = Range("B24").Value

    ' ▼ ロット
    wsDB.Cells(nextRow, 8).Value = Range("B25").Value

    ' ▼ 平均重量
    wsDB.Cells(nextRow, 9).Value = Range("B29").Value

    ' ▼ 準備時間
    wsDB.Cells(nextRow, 10).Value = Range("B26").Value

    ' ▼ 主作業時間
    wsDB.Cells(nextRow, 11).Value = Range("B27").Value

    ' ▼ 掃除時間
    wsDB.Cells(nextRow, 12).Value = Range("B28").Value

    ' ▼ 歩留まり
    wsDB.Cells(nextRow, 13).Value = Range("D23").Value

    ' ▼ 歩留り評価
    wsDB.Cells(nextRow, 14).Value = Range("B31").Value

    ' ▼ 時間評価
    wsDB.Cells(nextRow, 15).Value = Range("B32").Value

    ' ▼ 評価理由
    wsDB.Cells(nextRow, 16).Value = Range("B33").Value

    ' ▼ 工程メモ
    wsDB.Cells(nextRow, 17).Value = Range("B34").Value

    ' ▼ 中分類メモ
    wsDB.Cells(nextRow, 18).Value = Range("B35").Value

    ' ▼ 種類メモ
    wsDB.Cells(nextRow, 19).Value = Range("B36").Value

    ' ▼ 備考
    wsDB.Cells(nextRow, 20).Value = Range("B37").Value

    ' ▼ 登録日時（Now）
    wsDB.Cells(nextRow, 21).Value = Now

    MsgBox "登録が完了しました。"

End Sub
Sub 未完了登録()

    Dim wsDB As Worksheet
    Dim nextRow As Long
    Dim 種類 As String

    Set wsDB = Sheets("DB")
    nextRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).ROW + 1

    種類 = Range("B3").Value

    ' ▼ 種類
    wsDB.Cells(nextRow, 1).Value = 種類

    ' ▼ 作業日
    wsDB.Cells(nextRow, 2).Value = Range("B2").Value

    ' ▼ 名称
    If 種類 = "商品" Then
        wsDB.Cells(nextRow, 3).Value = Range("B8").Value
    Else
        wsDB.Cells(nextRow, 3).Value = Range("B9").Value
    End If

    ' ▼ 工程
    If 種類 = "商品" Then
        wsDB.Cells(nextRow, 4).Value = 商品工程文字列()
    Else
        wsDB.Cells(nextRow, 4).Value = Range("B21").Value
    End If

    ' ▼ 数量
    wsDB.Cells(nextRow, 5).Value = Range("B23").Value

    ' ▼ 単位
    wsDB.Cells(nextRow, 6).Value = Range("C23").Value

    ' ▼ 枚数
    wsDB.Cells(nextRow, 7).Value = Range("B24").Value

    ' ▼ ロット
    wsDB.Cells(nextRow, 8).Value = Range("B25").Value

    ' ▼ 平均重量
    wsDB.Cells(nextRow, 9).Value = Range("B29").Value

    ' ▼ 準備時間
    wsDB.Cells(nextRow, 10).Value = Range("B26").Value

    ' ▼ 主作業時間
    wsDB.Cells(nextRow, 11).Value = Range("B27").Value

    ' ▼ 掃除時間
    wsDB.Cells(nextRow, 12).Value = Range("B28").Value

    ' ▼ 歩留まり
    wsDB.Cells(nextRow, 13).Value = Range("D23").Value

    ' ▼ 歩留り評価
    wsDB.Cells(nextRow, 14).Value = Range("B31").Value

    ' ▼ 時間評価
    wsDB.Cells(nextRow, 15).Value = Range("B32").Value

    ' ▼ 評価理由
    wsDB.Cells(nextRow, 16).Value = Range("B33").Value

    ' ▼ 工程メモ
    wsDB.Cells(nextRow, 17).Value = Range("B34").Value

    ' ▼ 中分類メモ
    wsDB.Cells(nextRow, 18).Value = Range("B35").Value

    ' ▼ 種類メモ
    wsDB.Cells(nextRow, 19).Value = Range("B36").Value

    ' ▼ 備考
    wsDB.Cells(nextRow, 20).Value = Range("B37").Value

    ' ▼ 登録日時
    wsDB.Cells(nextRow, 21).Value = Now

    ' ▼ 未完了フラグ（22列目）
    wsDB.Cells(nextRow, 22).Value = "未完"

    MsgBox "未完了として保存しました。"

End Sub
Sub 未完了一覧生成()

    Dim wsList As Worksheet
    Dim wsDB As Worksheet
    Dim lastRow As Long
    Dim listRow As Long
    Dim i As Long
    Dim btn As Button

    Set wsList = Sheets("未完了一覧")
    Set wsDB = Sheets("DB")

    ' ▼ 一覧シート初期化
    wsList.Cells.Clear

    ' ▼ 見出し
    wsList.Range("A1").Value = "DB行"
    wsList.Range("B1").Value = "作業日"
    wsList.Range("C1").Value = "種類"
    wsList.Range("D1").Value = "名称"
    wsList.Range("E1").Value = "工程"
    wsList.Range("F1").Value = "数量"
    wsList.Range("G1").Value = "ロット"
    wsList.Range("H1").Value = "呼び出し"

    listRow = 2
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).ROW

    ' ▼ DB から未完了だけ抽出
    For i = 2 To lastRow
        If wsDB.Cells(i, 22).Value = "未完" Then

            wsList.Cells(listRow, 1).Value = i                      ' DB行番号
            wsList.Cells(listRow, 2).Value = wsDB.Cells(i, 2).Value ' 作業日
            wsList.Cells(listRow, 3).Value = wsDB.Cells(i, 1).Value ' 種類
            wsList.Cells(listRow, 4).Value = wsDB.Cells(i, 3).Value ' 名称
            wsList.Cells(listRow, 5).Value = wsDB.Cells(i, 4).Value ' 工程
            wsList.Cells(listRow, 6).Value = wsDB.Cells(i, 5).Value ' 数量
            wsList.Cells(listRow, 7).Value = wsDB.Cells(i, 8).Value ' ロット

            ' ▼ 呼び出しボタン作成
            Set btn = wsList.Buttons.Add( _
                wsList.Cells(listRow, 8).Left, _
                wsList.Cells(listRow, 8).Top, _
                wsList.Cells(listRow, 8).Width, _
                wsList.Cells(listRow, 8).Height)

            btn.Caption = "呼び出し"
            btn.OnAction = "未完了呼び出し"

            listRow = listRow + 1
        End If
    Next i

    MsgBox "未完了一覧を更新しました。"

End Sub
Sub 未完了呼び出し()

    Dim wsList As Worksheet
    Dim wsDB As Worksheet
    Dim wsInput As Worksheet
    Dim btn As Button
    Dim rowNum As Long
    Dim dbRow As Long
    Dim 種類 As String

    Set wsList = Sheets("未完了一覧")
    Set wsDB = Sheets("DB")
    Set wsInput = Sheets("Input")

    ' ▼ どのボタンが押されたか取得
    Set btn = wsList.Buttons(Application.Caller)
    rowNum = btn.TopLeftCell.ROW

    ' ▼ DBの行番号
    dbRow = wsList.Cells(rowNum, 1).Value

    ' ▼ 種類（商品/半製品）
    種類 = wsDB.Cells(dbRow, 1).Value
    wsInput.Range("B3").Value = 種類

    ' ▼ 作業日
    wsInput.Range("B2").Value = wsDB.Cells(dbRow, 2).Value

    ' ▼ 名称
    If 種類 = "商品" Then
        wsInput.Range("B8").Value = wsDB.Cells(dbRow, 3).Value
    Else
        wsInput.Range("B9").Value = wsDB.Cells(dbRow, 3).Value
    End If

    ' ▼ 工程
    If 種類 = "商品" Then
        Call 商品工程復元(wsDB.Cells(dbRow, 4).Value)
    Else
        wsInput.Range("B21").Value = wsDB.Cells(dbRow, 4).Value
    End If

    ' ▼ 数量・単位・枚数・ロット・平均重量
    wsInput.Range("B23").Value = wsDB.Cells(dbRow, 5).Value
    wsInput.Range("C23").Value = wsDB.Cells(dbRow, 6).Value
    wsInput.Range("B24").Value = wsDB.Cells(dbRow, 7).Value
    wsInput.Range("B25").Value = wsDB.Cells(dbRow, 8).Value
    wsInput.Range("B29").Value = wsDB.Cells(dbRow, 9).Value

    ' ▼ 時間系
    wsInput.Range("B26").Value = wsDB.Cells(dbRow, 10).Value
    wsInput.Range("B27").Value = wsDB.Cells(dbRow, 11).Value
    wsInput.Range("B28").Value = wsDB.Cells(dbRow, 12).Value

    ' ▼ 歩留まり
    wsInput.Range("D23").Value = wsDB.Cells(dbRow, 13).Value

    ' ▼ 評価
    wsInput.Range("B31").Value = wsDB.Cells(dbRow, 14).Value
    wsInput.Range("B32").Value = wsDB.Cells(dbRow, 15).Value

    ' ▼ メモ類
    wsInput.Range("B33").Value = wsDB.Cells(dbRow, 16).Value
    wsInput.Range("B34").Value = wsDB.Cells(dbRow, 17).Value
    wsInput.Range("B35").Value = wsDB.Cells(dbRow, 18).Value
    wsInput.Range("B36").Value = wsDB.Cells(dbRow, 19).Value
    wsInput.Range("B37").Value = wsDB.Cells(dbRow, 20).Value

    ' ▼ 工程判定を再実行（半製品のみ）
    If 種類 = "半製品" Then
        Call 半製品_工程判定(wsInput.Range("B21").Value)
    End If

    MsgBox "未完了データを呼び出しました。"

End Sub

