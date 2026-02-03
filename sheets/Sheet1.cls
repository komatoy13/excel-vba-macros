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

'==========================
' メイン：Worksheet_Change
'==========================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit

    Call UnprotectInputSheet
    Application.EnableEvents = False

    ' 常に作業区分・工程プルダウンは維持
    Call SetWorkTypeDropdown
    Call SetProcessDropdown

    Select Case Target.Address

        ' 作業区分
        Case "$B$3"
            Call ClearForWorkTypeChange
            Call SetDaiDropdown(Me.Range("B5").Value)
            Call SetShoDropdown(Me.Range("B5").Value, Me.Range("B6").Value)
            Call SetUnit
            Call SetInputFields
            Call SetLockFields
            Call UpdateNameBySelection

        ' 工程
        Case "$B$5"
            Call SetDaiDropdown(Me.Range("B5").Value)
            Call SetShoDropdown(Me.Range("B5").Value, Me.Range("B6").Value)
            Call SetUnit
            Call SetInputFields
            Call SetLockFields
            Call UpdateNameBySelection

        ' 大分類
        Case "$B$6"
            Call SetShoDropdown(Me.Range("B5").Value, Me.Range("B6").Value)
            Call SetInputFields
            Call SetLockFields
            Call UpdateNameBySelection

        ' 中分類
        Case "$B$7"
            Call SetInputFields
            Call SetLockFields
            Call UpdateNameBySelection

        ' 詳分類
        Case "$B$8"
            Call SetInputFields
            Call SetLockFields
            Call UpdateNameBySelection

        ' 作業コード
        Case "$B$4"
            Call HandleWorkCodeChange(Target)

    End Select

SafeExit:
    Application.EnableEvents = True
    Call ProtectInputSheet

End Sub

'==========================
' 作業コード変更時
'==========================
Private Sub HandleWorkCodeChange(ByVal Target As Range)

    Dim code As String

    If Target.Value = "" Then
        Me.Range("B3").Value = ""
        Me.Range("B5").Value = ""
        Me.Range("B6").Value = ""
        Me.Range("B7").Value = ""
        Me.Range("B8").Value = ""
        Me.Range("B9").Value = ""
        Call SetLockFields
        Exit Sub
    End If

    code = Trim(Target.Value)

    If IsProductCode(code) Then
        Me.Range("B3").Value = "製品"
        Call FillProduct(code)
    Else
        Me.Range("B3").Value = "半製品"
        Call FillHalfProduct(code)
    End If

    Call SetDaiDropdown(Me.Range("B5").Value)
    Call SetShoDropdown(Me.Range("B5").Value, Me.Range("B6").Value)
    Call SetUnit
    Call SetInputFields
    Call SetLockFields

End Sub

'==========================
' 作業区分プルダウン
'==========================
Private Sub SetWorkTypeDropdown()
    With Me.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="製品,半製品"
    End With
End Sub

'==========================
' 工程プルダウン
'==========================
Private Sub SetProcessDropdown()

    Dim listStr As String
    listStr = "成型,検品,個包装,プリント,包装糖,カット・検品,メレンゲ,着色,デコ,LS,LS個包装,箱折,緩衝材裁断,パッキン裁断"

    With Me.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=listStr
    End With

End Sub

'==========================
' 作業区分変更時のクリア
'==========================
Private Sub ClearForWorkTypeChange()
    Me.Range("B4").Value = ""
    Me.Range("B5").Value = ""
    Me.Range("B6").Value = ""
    Me.Range("B7").Value = ""
    Me.Range("B8").Value = ""
    Me.Range("B9").Value = ""
End Sub

'==========================
' 大分類プルダウン（B6）
'==========================
Private Sub SetDaiDropdown(process As String)

    Dim workType As String
    workType = Me.Range("B3").Value

    ' 例外工程：パッキン裁断・緩衝材裁断 → B6 グレーアウト
    If process = "パッキン裁断" Or process = "緩衝材裁断" Then
        Me.Range("B6").Validation.Delete
        Call LockAndGray(Me.Range("B6"))
        Exit Sub
    End If

    ' 製品：T商品名「中分類」
    If workType = "製品" Then

        Dim loP As ListObject
        Dim rP As listRow
        Dim dictP As Object
        Set dictP = CreateObject("Scripting.Dictionary")

        Set loP = Sheets("T商品名").ListObjects("T商品名")

        For Each rP In loP.ListRows
            dictP(rP.Range(1, loP.ListColumns("中分類").Index).Value) = 1
        Next rP

        With Me.Range("B6").Validation
            .Delete
            If dictP.Count > 0 Then
                .Add Type:=xlValidateList, Formula1:=Join(dictP.Keys, ",")
            End If
        End With

        Call UnlockAndWhite(Me.Range("B6"))
        Exit Sub
    End If

    ' 半製品：半製品マスタ「形状」（工程マッピングあり）
    Dim lo As ListObject
    Dim r As listRow
    Dim dict As Object
    Dim refProcess As String
    Set dict = CreateObject("Scripting.Dictionary")

    Select Case process
        Case "検品": refProcess = "成型"
        Case "カット・検品": refProcess = "包装糖"
        Case Else: refProcess = process
    End Select

    Set lo = Sheets("半製品マスタ").ListObjects("ＴＤＳ")

    For Each r In lo.ListRows
        If r.Range(1, lo.ListColumns("工程").Index).Value = refProcess Then
            dict(r.Range(1, lo.ListColumns("形状").Index).Value) = 1
        End If
    Next r

    With Me.Range("B6").Validation
        .Delete
        If dict.Count > 0 Then
            .Add Type:=xlValidateList, Formula1:=Join(dict.Keys, ",")
        End If
    End With

    Call UnlockAndWhite(Me.Range("B6"))

End Sub

'==========================
' 中分類プルダウン（B7）＋プリント時 B8
'==========================
Private Sub SetShoDropdown(process As String, dai As String)

    Dim workType As String
    workType = Me.Range("B3").Value
    dai = Trim(dai)

    ' 製品：中分類→小分類
    If workType = "製品" Then

        Dim loP As ListObject
        Dim rP As listRow
        Dim dictP As Object
        Set dictP = CreateObject("Scripting.Dictionary")

        Set loP = Sheets("T商品名").ListObjects("T商品名")

        For Each rP In loP.ListRows
            If rP.Range(1, loP.ListColumns("中分類").Index).Value = dai Then
                dictP(rP.Range(1, loP.ListColumns("小分類").Index).Value) = 1
            End If
        Next rP

        With Me.Range("B7").Validation
            .Delete
            If dictP.Count > 0 Then
                .Add Type:=xlValidateList, Formula1:=Join(dictP.Keys, ",")
            End If
        End With

        Me.Range("B8").Validation.Delete
        Call UnlockAndWhite(Me.Range("B8"))
        Exit Sub
    End If

    ' 半製品：例外工程（パッキン裁断・緩衝材裁断）
    If process = "パッキン裁断" Or process = "緩衝材裁断" Then
        Me.Range("B7").Validation.Delete
        Call LockAndGray(Me.Range("B7"))
        Me.Range("B8").Validation.Delete
        Call UnlockAndWhite(Me.Range("B8"))
        Exit Sub
    End If

    ' 半製品：中分類を使わない工程 → B7 グレーアウト
    If process = "包装糖" Or process = "カット・検品" _
       Or process = "メレンゲ" Or process = "着色" _
       Or process = "箱折" Or process = "LS" Then

        Me.Range("B7").Validation.Delete
        Call LockAndGray(Me.Range("B7"))
        Me.Range("B8").Validation.Delete
        Call UnlockAndWhite(Me.Range("B8"))
        Exit Sub
    End If

    Dim loH As ListObject
    Dim rH As listRow
    Set loH = Sheets("半製品マスタ").ListObjects("ＴＤＳ")

    ' 半製品：プリント
    If process = "プリント" Then

        Dim dictPrint As Object
        Set dictPrint = CreateObject("Scripting.Dictionary")

        For Each rH In loH.ListRows
            If rH.Range(1, loH.ListColumns("形状").Index).Value = dai Then
                dictPrint(rH.Range(1, loH.ListColumns("プリント").Index).Value) = 1
            End If
        Next rH

        With Me.Range("B7").Validation
            .Delete
            If dictPrint.Count > 0 Then
                .Add Type:=xlValidateList, Formula1:=Join(dictPrint.Keys, ",")
            End If
        End With

        Dim dictColor As Object
        Set dictColor = CreateObject("Scripting.Dictionary")

        For Each rH In loH.ListRows
            If rH.Range(1, loH.ListColumns("形状").Index).Value = dai Then
                dictColor(rH.Range(1, loH.ListColumns("色").Index).Value) = 1
            End If
        Next rH

        With Me.Range("B8").Validation
            .Delete
            If dictColor.Count > 0 Then
                .Add Type:=xlValidateList, Formula1:=Join(dictColor.Keys, ",")
            End If
        End With

        Call UnlockAndWhite(Me.Range("B8"))
        Exit Sub
    End If

    ' 半製品：LS個包装 → B7 = 分類
    If process = "LS個包装" Then

        Dim dictClass As Object
        Set dictClass = CreateObject("Scripting.Dictionary")

        For Each rH In loH.ListRows
            If rH.Range(1, loH.ListColumns("工程").Index).Value = process _
               And rH.Range(1, loH.ListColumns("形状").Index).Value = dai Then
                dictClass(rH.Range(1, loH.ListColumns("分類").Index).Value) = 1
            End If
        Next rH

        With Me.Range("B7").Validation
            .Delete
            If dictClass.Count > 0 Then
                .Add Type:=xlValidateList, Formula1:=Join(dictClass.Keys, ",")
            End If
        End With

        Me.Range("B8").Validation.Delete
        Call UnlockAndWhite(Me.Range("B8"))
        Exit Sub
    End If

    ' 半製品：通常工程 → B7 = 色
    Dim dictN As Object
    Set dictN = CreateObject("Scripting.Dictionary")

    Dim refProcess As String
    Select Case process
        Case "検品": refProcess = "成型"
        Case Else: refProcess = process
    End Select

    For Each rH In loH.ListRows
        If rH.Range(1, loH.ListColumns("工程").Index).Value = refProcess _
           And rH.Range(1, loH.ListColumns("形状").Index).Value = dai Then
            dictN(rH.Range(1, loH.ListColumns("色").Index).Value) = 1
        End If
    Next rH

    With Me.Range("B7").Validation
        .Delete
        If dictN.Count > 0 Then
            .Add Type:=xlValidateList, Formula1:=Join(dictN.Keys, ",")
        End If
    End With

    Me.Range("B8").Validation.Delete
    Call UnlockAndWhite(Me.Range("B8"))

End Sub

'==========================
' ロック制御
'==========================
Private Sub SetLockFields()

    ' B4：常に入力可
    Call UnlockAndWhite(Me.Range("B4"))
    ' B8：常に入力可
    Call UnlockAndWhite(Me.Range("B8"))
    ' B9：常にロック（白）
    Call LockAsLabel(Me.Range("B9"))

    Dim wt As String, proc As String
    wt = Me.Range("B3").Value
    proc = Me.Range("B5").Value

    ' 製品：B5 常にグレーアウト
    If wt = "製品" Then

        Call LockAndGray(Me.Range("B5"))

        If Me.Range("B4").Value = "" Then
            ' コードなし → B6/B7 入力可
            Call UnlockAndWhite(Me.Range("B6"))
            Call UnlockAndWhite(Me.Range("B7"))
        Else
            ' コードあり → B6/B7 白ロック
            Call LockAsLabel(Me.Range("B6"))
            Call LockAsLabel(Me.Range("B7"))
        End If

        GoTo AlwaysUnlock
    End If

    ' 半製品：コードなし → 全入力可
    If Me.Range("B4").Value = "" Then
        Call UnlockAndWhite(Me.Range("B3"))
        Call UnlockAndWhite(Me.Range("B5"))
        GoTo AlwaysUnlock
    End If

    ' 半製品：コードあり
    ' 例外工程
    If proc = "パッキン裁断" Or proc = "緩衝材裁断" Then
        Call LockAsLabel(Me.Range("B3"))
        Call LockAsLabel(Me.Range("B5"))
        Call LockAndGray(Me.Range("B6"))
        Call LockAndGray(Me.Range("B7"))
        Me.Range("B9").Value = proc
        GoTo AlwaysUnlock
    End If

    ' 通常半製品
    Call LockAsLabel(Me.Range("B3"))
    Call LockAsLabel(Me.Range("B5"))

    If Me.Range("B6").Interior.color <> RGB(230, 230, 230) Then
        Call LockAsLabel(Me.Range("B6"))
    End If
    If Me.Range("B7").Interior.color <> RGB(230, 230, 230) Then
        Call LockAsLabel(Me.Range("B7"))
    End If

AlwaysUnlock:

    '==========================
    ' 常にロック解除するセル
    '==========================
    Call UnlockAndWhite(Me.Range("B2"))
    Call UnlockAndWhite(Me.Range("B17"))
    Call UnlockAndWhite(Me.Range("B20"))
    Call UnlockAndWhite(Me.Range("B21"))
    Call UnlockAndWhite(Me.Range("B22"))
    Call UnlockAndWhite(Me.Range("B25:B31"))

End Sub

'==========================
' 名称自動設定（B4 空欄時）
'==========================
Private Sub UpdateNameBySelection()

    If Me.Range("B4").Value <> "" Then Exit Sub

    If Me.Range("B3").Value = "製品" Then
        Call SetNameForProduct
    ElseIf Me.Range("B3").Value = "半製品" Then
        Call SetNameForHalf
    End If

End Sub

' 製品：中分類＋小分類 → 商品名
Private Sub SetNameForProduct()

    Dim lo As ListObject
    Dim r As listRow
    Dim dai As String, sho As String

    dai = Me.Range("B6").Value
    sho = Me.Range("B7").Value

    If dai = "" Then Exit Sub

    Set lo = Sheets("T商品名").ListObjects("T商品名")

    If sho = "" Then
        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("中分類").Index).Value = dai Then
                Me.Range("B9").Value = r.Range(1, lo.ListColumns("商品名").Index).Value
                Exit Sub
            End If
        Next r
        Exit Sub
    End If

    For Each r In lo.ListRows
        If r.Range(1, lo.ListColumns("中分類").Index).Value = dai _
           And r.Range(1, lo.ListColumns("小分類").Index).Value = sho Then
            Me.Range("B9").Value = r.Range(1, lo.ListColumns("商品名").Index).Value
            Exit Sub
        End If
    Next r

End Sub

' 半製品：工程・形状・色・分類から種類
Private Sub SetNameForHalf()

    Dim lo As ListObject
    Dim r As listRow
    Dim proc As String, dai As String, sho As String

    proc = Me.Range("B5").Value
    dai = Me.Range("B6").Value
    sho = Me.Range("B7").Value
        If proc = "" Or dai = "" Then Exit Sub

    Set lo = Sheets("半製品マスタ").ListObjects("ＴＤＳ")

    ' 中分類を使わない工程：包装糖・カット検品・メレンゲ・着色・箱折・LS
    If proc = "包装糖" Or proc = "カット・検品" _
       Or proc = "メレンゲ" Or proc = "着色" _
       Or proc = "箱折" Or proc = "LS" Then

        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("工程").Index).Value = proc _
               And r.Range(1, lo.ListColumns("形状").Index).Value = dai Then
                Me.Range("B9").Value = r.Range(1, lo.ListColumns("種類").Index).Value
                Exit Sub
            End If
        Next r

        Exit Sub
    End If

    ' LS個包装：形状＋分類
    If proc = "LS個包装" Then

        If sho = "" Then Exit Sub

        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("工程").Index).Value = proc _
               And r.Range(1, lo.ListColumns("形状").Index).Value = dai _
               And r.Range(1, lo.ListColumns("分類").Index).Value = sho Then
                Me.Range("B9").Value = r.Range(1, lo.ListColumns("種類").Index).Value
                Exit Sub
            End If
        Next r

        Exit Sub
    End If

    ' プリント：形状＋プリント＋色
    If proc = "プリント" Then

        Dim colPrint As Long, colShape As Long, colColor As Long, colKind As Long
        colPrint = lo.ListColumns("プリント").Index
        colShape = lo.ListColumns("形状").Index
        colColor = lo.ListColumns("色").Index
        colKind = lo.ListColumns("種類").Index

        Dim valPrint As String, valColor As String
        valPrint = Me.Range("B7").Value
        valColor = Me.Range("B8").Value

        If valPrint = "" Or valColor = "" Then Exit Sub

        For Each r In lo.ListRows
            If r.Range(1, colShape).Value = dai _
               And r.Range(1, colPrint).Value = valPrint _
               And r.Range(1, colColor).Value = valColor Then
                Me.Range("B9").Value = r.Range(1, colKind).Value
                Exit Sub
            End If
        Next r

        Exit Sub
    End If

    ' 個包装：形状＋色
    If proc = "個包装" Then

        If sho = "" Then Exit Sub

        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("工程").Index).Value = proc _
               And r.Range(1, lo.ListColumns("形状").Index).Value = dai _
               And r.Range(1, lo.ListColumns("色").Index).Value = sho Then
                Me.Range("B9").Value = r.Range(1, lo.ListColumns("種類").Index).Value
                Exit Sub
            End If
        Next r

        Exit Sub
    End If

    ' デコ：形状＋色
    If proc = "デコ" Then

        If sho = "" Then Exit Sub

        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("工程").Index).Value = proc _
               And r.Range(1, lo.ListColumns("形状").Index).Value = dai _
               And r.Range(1, lo.ListColumns("色").Index).Value = sho Then
                Me.Range("B9").Value = r.Range(1, lo.ListColumns("種類").Index).Value
                Exit Sub
            End If
        Next r

        Exit Sub
    End If

    ' 通常：工程＋形状＋色（検品は成型扱い）
    If sho = "" Then Exit Sub

    Dim refProcess As String
    Select Case proc
        Case "検品": refProcess = "成型"
        Case Else: refProcess = proc
    End Select

    For Each r In lo.ListRows
        If r.Range(1, lo.ListColumns("工程").Index).Value = refProcess _
           And r.Range(1, lo.ListColumns("形状").Index).Value = dai _
           And r.Range(1, lo.ListColumns("色").Index).Value = sho Then
            Me.Range("B9").Value = r.Range(1, lo.ListColumns("種類").Index).Value
            Exit Sub
        End If
    Next r

End Sub

'==========================
' 工程に応じた入力欄制御
'==========================
Private Sub SetInputFields()

    Dim wt As String, proc As String, dai As String
    wt = Me.Range("B3").Value
    proc = Me.Range("B5").Value
    dai = Trim(Me.Range("B6").Value)

    Call LockAndGray(Me.Range("B18"))
    Call LockAndGray(Me.Range("B19"))
    Call LockAndGray(Me.Range("B23"))

    ' B18：枚数 → 成型のみ
    If proc = "成型" Then
        Call UnlockAndWhite(Me.Range("B18"))
    End If

    ' B19：ロット
    If wt = "製品" Then
        Call UnlockAndWhite(Me.Range("B19"))
    Else
        Select Case proc
            Case "成型", "LS", "包装糖", "パッキン裁断", "緩衝材裁断", "箱折", "メレンゲ"
                ' ロックのまま
            Case Else
                Call UnlockAndWhite(Me.Range("B19"))
        End Select
    End If

    ' B23：平均重量
    If wt = "半製品" And (proc = "検品" Or proc = "プリント" Or proc = "個包装") Then

        Dim candidates As Collection
        Set candidates = New Collection

        If proc = "個包装" Then
            candidates.Add dai & "_個包装"
            candidates.Add dai & "個包装"
        Else
            candidates.Add dai
        End If

        Dim lo As ListObject
        Dim r As listRow
        Dim nameCol As Long, weightCol As Long
        Dim matchFound As Boolean
        Dim nm As String, c As Variant

        Set lo = Sheets("型ｄ").ListObjects("型ｄ")
        nameCol = lo.ListColumns("名称").Index
        weightCol = lo.ListColumns("検品時計量").Index
        matchFound = False

        For Each r In lo.ListRows
            nm = Trim(r.Range(1, nameCol).Value)

            For Each c In candidates
                If nm = c Then
                    If r.Range(1, weightCol).Value = 1 Then
                        matchFound = True
                        Exit For
                    End If
                End If
            Next c

            If matchFound Then Exit For
        Next r

        If matchFound Then
            Call UnlockAndWhite(Me.Range("B23"))
        End If
    End If

End Sub

'==========================
' 製品コード判定
'==========================
Private Function IsProductCode(code As String) As Boolean
    IsProductCode = (Len(code) = 4 And InStr(code, "-") = 0)
End Function

'==========================
' 製品データ入力
'==========================
Private Sub FillProduct(code As String)

    Dim lo As ListObject
    Dim r As listRow

    Set lo = Sheets("T商品名").ListObjects("T商品名")

    For Each r In lo.ListRows
        If r.Range(1, lo.ListColumns("コード").Index).Value = code Then

            Me.Range("B5").Value = ""  ' 工程なし
            Me.Range("B6").Value = r.Range(1, lo.ListColumns("中分類").Index).Value
            Me.Range("B7").Value = r.Range(1, lo.ListColumns("小分類").Index).Value
            Me.Range("B8").Value = ""
            Me.Range("B9").Value = r.Range(1, lo.ListColumns("商品名").Index).Value

            Exit Sub
        End If
    Next r

    MsgBox "製品コードが見つかりません。", vbExclamation

End Sub

'==========================
' 半製品データ入力
'==========================
Private Sub FillHalfProduct(code As String)

    Dim lo As ListObject
    Dim r As listRow
    Dim baseCode As String, engCode As String, engName As String

    Set lo = Sheets("半製品マスタ").ListObjects("ＴＤＳ")

    ' 通常パターン
    For Each r In lo.ListRows
        If r.Range(1, lo.ListColumns("記号").Index).Value = code Then
            engName = r.Range(1, lo.ListColumns("工程").Index).Value
            Call WriteHalfProduct(r, engName)
            Exit Sub
        End If
    Next r

    ' イレギュラー（1字-xxxx-yy）
    If Len(code) = 9 And Mid(code, 2, 1) = "-" Then

        engCode = Left(code, 1)
        baseCode = Mid(code, 3)

        Select Case engCode
            Case "f": engName = "成型"
            Case "c": engName = "検品"
            Case "h": engName = "包装糖"
            Case "k": engName = "カット・検品"
            Case Else: engName = ""
        End Select

        For Each r In lo.ListRows
            If r.Range(1, lo.ListColumns("記号").Index).Value = baseCode Then
                Call WriteHalfProduct(r, engName)
                Exit Sub
            End If
        Next r

    End If

    MsgBox "半製品コードが見つかりません。", vbExclamation

End Sub

'==========================
' 半製品書き込み共通
'==========================
Private Sub WriteHalfProduct(r As listRow, engName As String)

    Dim shape As String, color As String, kind As String

    shape = r.Range(1, r.Parent.ListColumns("形状").Index).Value
    color = r.Range(1, r.Parent.ListColumns("色").Index).Value
    kind = r.Range(1, r.Parent.ListColumns("種類").Index).Value

    Me.Range("B5").Value = engName
    Me.Range("B6").Value = shape

    If engName = "プリント" Then
        Me.Range("B7").Value = "プリント"
        Me.Range("B8").Value = color
    Else
        Me.Range("B7").Value = color
        Me.Range("B8").Value = ""
    End If

    Me.Range("B9").Value = kind

End Sub

'==========================
' 単位（C17）
'==========================
Private Sub SetUnit()

    Dim wt As String, proc As String, unit As String

    wt = Me.Range("B3").Value
    proc = Me.Range("B5").Value

    If wt = "製品" Then
        Me.Range("C17").Value = "個"
        Exit Sub
    End If

    Select Case proc
        Case "成型", "検品", "LS"
            unit = "㎏"
        Case "プリント", "デコ"
            unit = "粒"
        Case "包装糖", "カット・検品"
            unit = "玉"
        Case "メレンゲ", "着色"
            unit = "g"
        Case "箱折"
            unit = "箱"
        Case "パッキン裁断", "緩衝材裁断"
            unit = "枚"
        Case "個包装", "LS個包装"
            unit = "包"
        Case Else
            unit = ""
    End Select

    Me.Range("C17").Value = unit

End Sub

'==========================
' ロック系共通
'==========================
Private Sub LockAsLabel(rng As Range)
    With rng
        .Locked = True
        .Interior.color = RGB(255, 255, 255)
        .Font.color = RGB(0, 0, 0)
    End With
End Sub

Private Sub LockAndGray(rng As Range)
    With rng
        .Locked = True
        .Interior.color = RGB(230, 230, 230)
        .Font.color = RGB(150, 150, 150)
    End With
End Sub

Private Sub UnlockAndWhite(rng As Range)
    With rng
        .Locked = False
        .Interior.color = RGB(255, 255, 255)
        .Font.color = RGB(0, 0, 0)
    End With
End Sub

'==========================
' シート保護
'==========================
Private Sub ProtectInputSheet()
    With Me
        .Protect Password:="input", UserInterfaceOnly:=True
        .EnableSelection = xlUnlockedCells
    End With
End Sub

Private Sub UnprotectInputSheet()
    Me.Unprotect Password:="input"
End Sub


