Attribute VB_Name = "Module_Init"
Option Explicit

Sub 商品モード初期化()

    ' ▼ 半製品専用項目をクリア・非活性
    Range("B9").Value = ""          ' 半製品名
    Range("B21").Value = ""         ' 半製品工程
    Range("B21").Locked = True

    Range("B24").Value = ""         ' 枚数
    Range("B24").Locked = True

    Range("B29").Value = ""         ' 平均重量
    Range("B29").Locked = True

    ' ▼ ロットは任意（商品でも入力可）
    ' → 値はクリアするがロックはしない
    Range("B25").Value = ""         ' ロット
    Range("B25").Locked = False

    ' ▼ 単位は「個」に固定
    Range("C23").Value = "個"

    ' ▼ 歩留まり評価を非活性
    Range("B31").Locked = True
    Range("B31").Value = ""

    ' ▼ 商品工程チェックボックスを活性化
    Call 商品工程チェック_有効化(True)

End Sub
Sub 半製品モード初期化()

    ' ▼ 半製品専用項目を活性化
    Range("B21").Locked = False     ' 半製品工程
    Range("B24").Locked = False     ' 枚数（工程でON/OFF）
    Range("B29").Locked = False     ' 平均重量（工程＋半製品名でON/OFF）
    Range("B25").Locked = False     ' ロット（工程でON/OFF）

    ' ▼ 商品工程チェックボックスを無効化
    Call 商品工程チェック_有効化(False)

    ' ▼ 歩留まり評価を活性化
    Range("B31").Locked = False

End Sub
Sub 商品工程チェック_有効化(flag As Boolean)

    Dim arr As Variant
    Dim i As Long

    arr = Array("A15", "A16", "A17", "A18", "A19", _
                "C15", "C16", "C17", "C19")

    For i = LBound(arr) To UBound(arr)
        With Range(arr(i))
            .Value = False
            .Locked = Not flag
        End With
    Next i

End Sub
