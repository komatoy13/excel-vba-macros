Attribute VB_Name = "Module_Yield"
Option Explicit

Function 前工程名取得(現名称 As String, 現工程 As String) As String

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Range

    Set ws = Sheets("名称マッピング")
    Set tbl = ws.ListObjects("名称マッピング")

    前工程名取得 = ""

    For Each r In tbl.ListColumns("現工程の名称").DataBodyRange
        If r.Value = 現名称 And r.Offset(0, 2).Value = 現工程 Then
            前工程名取得 = r.Offset(0, 1).Value
            Exit Function
        End If
    Next r

End Function
Function 前工程名_ユーザー選択(現名称 As String, 現工程 As String) As String

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Range
    Dim 候補 As String
    Dim 選択 As String

    Set ws = Sheets("半製品マスタ")
    Set tbl = ws.ListObjects(1)

    候補 = ""

    For Each r In tbl.ListColumns("種類").DataBodyRange
        候補 = 候補 & r.Value & vbCrLf
    Next r

    選択 = InputBox("前工程が登録されていません。" & vbCrLf & _
                     "前工程の半製品名を選んでください。" & vbCrLf & vbCrLf & 候補, _
                     "前工程の選択")

    If 選択 = "" Then
        前工程名_ユーザー選択 = ""
        Exit Function
    End If

    Call 名称マッピング追加(現名称, 選択, 現工程)

    前工程名_ユーザー選択 = 選択

End Function
Sub 名称マッピング追加(現名称 As String, 前名称 As String, 工程 As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As listRow

    Set ws = Sheets("名称マッピング")
    Set tbl = ws.ListObjects("名称マッピング")

    Set newRow = tbl.ListRows.Add

    newRow.Range(1, 1).Value = 現名称
    newRow.Range(1, 2).Value = 前名称
    newRow.Range(1, 3).Value = 工程
    newRow.Range(1, 4).Value = Date

End Sub
Function 前工程数量取得(前名称 As String, 前作業日 As Date) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = Sheets("DB")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).ROW

    前工程数量取得 = 0

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = 前名称 And ws.Cells(i, 2).Value = 前作業日 Then
            前工程数量取得 = ws.Cells(i, 6).Value
            Exit Function
        End If
    Next i

End Function
Sub 歩留まり計算()

    Dim 現名称 As String
    Dim 現工程 As String
    Dim 現数量 As Double
    Dim ロット As Date

    Dim 前名称 As String
    Dim 前数量 As Double
    Dim 歩留 As Double

    現名称 = Range("B9").Value
    現工程 = Range("B21").Value
    現数量 = Val(Range("B23").Value)
    ロット = Range("B25").Value

    If 現名称 = "" Or 現工程 = "" Or ロット = 0 Then Exit Sub

    ' ▼ 名称マッピング表から前工程を探す
    前名称 = 前工程名取得(現名称, 現工程)

    ' ▼ 見つからなければユーザーに選ばせる
    If 前名称 = "" Then
        前名称 = 前工程名_ユーザー選択(現名称, 現工程)
        If 前名称 = "" Then Exit Sub
    End If

    ' ▼ DBから前工程数量を取得
    前数量 = 前工程数量取得(前名称, ロット)

    If 前数量 = 0 Then
        MsgBox "前工程の数量がDBに見つかりません。"
        Exit Sub
    End If

    ' ▼ 歩留まり計算
    歩留 = 現数量 / 前数量

    Range("D23").Value = 歩留

    ' ▼ 異常値は赤く
    If 歩留 < 0.7 Or 歩留 > 1.1 Then
        Range("D23").Interior.Color = RGB(255, 150, 150)
    Else
        Range("D23").Interior.ColorIndex = 0
    End If

End Sub
