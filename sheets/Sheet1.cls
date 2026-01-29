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

    If Target.Address <> "$B$4" Then Exit Sub
    If Target.Value = "" Then Exit Sub

    Application.EnableEvents = False

    Dim code As String
    code = Trim(Target.Value)
    '===========================
    ' B3（作業区分）を自動判定
    '===========================
    If Len(code) = 4 And InStr(code, "-") = 0 Then
        Me.Range("B3").Value = "製品"
    Else
        Me.Range("B3").Value = "半製品"
    End If
    
    '===========================
    ' B5 のグレーアウト制御
    '===========================
    Call SetB5Appearance
    '===========================
    ' 製品 or 半製品の処理へ
    '===========================
    If Me.Range("B3").Value = "製品" Then
        Call FillProduct(code)
    Else
        Call FillHalfProduct(code)
    End If

    Application.EnableEvents = True
    Call SetDropdowns
End Sub
Private Sub SetDropdowns()

    Dim ws As Worksheet
    Set ws = Me

    ' 大分類（B6）
    If ws.Range("B3").Value = "製品" Then
        ws.Range("B6").Validation.Delete
        ws.Range("B6").Validation.Add Type:=xlValidateList, _
            Formula1:="=T商品名[中分類]"
    Else
        ws.Range("B6").Validation.Delete
        ws.Range("B6").Validation.Add Type:=xlValidateList, _
            Formula1:="=ＴＤＳ[形状]"
    End If

    ' 中分類（B7）
    If ws.Range("B3").Value = "製品" Then
        ws.Range("B7").Validation.Delete
        ws.Range("B7").Validation.Add Type:=xlValidateList, _
            Formula1:="=T商品名[小分類]"
    Else
        ws.Range("B7").Validation.Delete
        ws.Range("B7").Validation.Add Type:=xlValidateList, _
            Formula1:="=ＴＤＳ[色]"
    End If

    ' 詳分類（B8）→ 基本空欄なのでプルダウン無し
    ws.Range("B8").Validation.Delete

End Sub
Private Sub SetB5Appearance()

    With Me.Range("B5")
        If Me.Range("B3").Value = "製品" Then
            .Value = ""
            .Interior.color = RGB(230, 230, 230)
            .Font.color = RGB(150, 150, 150)
            .Locked = True
        Else
            .Interior.color = RGB(255, 255, 255)
            .Font.color = RGB(0, 0, 0)
            .Locked = True ' ← 半製品でも入力不可
        End If
    End With

    ' 名称（B9）も常に入力不可
    With Me.Range("B9")
        .Interior.color = RGB(240, 240, 240)
        .Font.color = RGB(0, 0, 0)
        .Locked = True
    End With

End Sub


'===========================================
' 製品：T商品名テーブルから取得（列名で参照）
'===========================================
Private Sub FillProduct(ByVal code As String)

    Dim lo As ListObject
    Dim r As listRow
    Dim found As Boolean

    Set lo = Sheets("T商品名").ListObjects("T商品名")

    found = False
    For Each r In lo.ListRows
        If r.Range(lo.ListColumns("コード").Index).Value = code Then
            found = True

            Me.Range("B5").Value = "" ' 製品は半製品工程を使わない
            Me.Range("B6").Value = r.Range(lo.ListColumns("中分類").Index).Value
            Me.Range("B7").Value = r.Range(lo.ListColumns("小分類").Index).Value
            Me.Range("B8").Value = ""
            Me.Range("B9").Value = r.Range(lo.ListColumns("商品名").Index).Value

            Exit For
        End If
    Next r

    If Not found Then
        MsgBox "製品コード """ & code & """ は T商品名 に存在しません。", vbExclamation
    End If

End Sub


'===========================================
' 半製品：ＴＤＳテーブルから取得（列名で参照）
'===========================================
Private Sub FillHalfProduct(ByVal code As String)

    Dim lo As ListObject
    Dim r As listRow
    Dim found As Boolean

    Set lo = Sheets("半製品マスタ").ListObjects("ＴＤＳ")

    '-----------------------------------------
    ' ① まず通常パターン（記号列と完全一致）を探す
    '-----------------------------------------
    found = False
    For Each r In lo.ListRows
        If r.Range(lo.ListColumns("記号").Index).Value = code Then

            Call WriteHalfProduct(r, r.Range(lo.ListColumns("工程").Index).Value)
            found = True
            Exit Sub

        End If
    Next r

    '-----------------------------------------
    ' ② 見つからなければイレギュラー（f/c/h/k）
    '-----------------------------------------
    If Len(code) = 9 And Mid(code, 2, 1) = "-" Then

        Dim engCode As String
        Dim baseCode As String
        Dim engName As String

        engCode = Left(code, 1)
        baseCode = Mid(code, 3)

        Select Case engCode
            Case "f": engName = "成型"
            Case "c": engName = "検品"
            Case "h": engName = "包装糖"
            Case "k": engName = "カット・検品"
            Case Else: engName = "不明"
        End Select

        For Each r In lo.ListRows
            If r.Range(lo.ListColumns("記号").Index).Value = baseCode Then

                Call WriteHalfProduct(r, engName)
                found = True
                Exit Sub

            End If
        Next r
    End If

    If Not found Then
        MsgBox "半製品コード """ & code & """ は ＴＤＳ に存在しません。", vbExclamation
    End If

End Sub


'===========================================
' 半製品の書き込み処理（共通）
'===========================================
Private Sub WriteHalfProduct(ByVal r As listRow, ByVal engName As String)

    Dim shape As String
    Dim color As String
    Dim typeName As String

    shape = r.Range(r.Parent.ListColumns("形状").Index).Value
    color = r.Range(r.Parent.ListColumns("色").Index).Value
    typeName = r.Range(r.Parent.ListColumns("種類").Index).Value

    Dim midClass As String
    Dim detailClass As String

    If engName = "プリント" Then
        midClass = "プリント"
        detailClass = color
    Else
        midClass = color
        detailClass = ""
    End If

    Me.Range("B5").Value = engName
    Me.Range("B6").Value = shape
    Me.Range("B7").Value = midClass
    Me.Range("B8").Value = detailClass
    Me.Range("B9").Value = typeName

End Sub
