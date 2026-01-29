Attribute VB_Name = "Module_ProductLogic"
Option Explicit
Function 商品工程文字列() As String

    Dim arr As Variant
    Dim result As String
    Dim i As Long

    arr = Array( _
        Array("A15", "内容ラベル貼り"), _
        Array("C15", "表ラベル貼り"), _
        Array("A16", "準備"), _
        Array("C16", "文字作り"), _
        Array("A17", "詰め作業"), _
        Array("C17", "熱処理"), _
        Array("A18", "ケースセット"), _
        Array("A19", "箱詰め"), _
        Array("C19", "梱包") _
    )

    result = ""

    For i = LBound(arr) To UBound(arr)
        If Range(arr(i)(0)).Value = True Then
            If result <> "" Then result = result & "・"
            result = result & arr(i)(1)
        End If
    Next i

    商品工程文字列 = result

End Function


