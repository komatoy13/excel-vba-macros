VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


' ★ GitHub ローカルリポジトリのベースパス（最後は \ で終わること）
Private Const BASE_PATH As String = _
    "C:\Users\TOYOMITSU_DOUKE.KOMAYA\OneDrive\仕事\VBA\日報記録\repo\excel-vba-macros\"

Private Const MODULES_PATH As String = BASE_PATH & "modules\"
Private Const SHEETS_PATH  As String = BASE_PATH & "sheets\"
Private Const FORMS_PATH   As String = BASE_PATH & "forms\"
Private Const CLASSES_PATH As String = BASE_PATH & "classes\"
Private Const DOCS_PATH    As String = BASE_PATH & "docs\"

' ★ 設計図シート名（必要に応じて変更）
Private Const SHEET_SYSTEM As String = "設計図"
Private Const SHEET_DBMAP  As String = "db_mapping"
Private Const SHEET_FUNC   As String = "functions"
Private Sub Workbook_Open()
    Sheets("Input").Range("B2").Value = Date
End Sub

'===========================================================
' Excel を閉じるときに自動で GitHub 用ファイルを出力
'===========================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Call ExportAllForGitHub
    On Error GoTo 0
End Sub


'===========================================================
' すべてのファイルを GitHub フォルダに出力するメイン処理
'===========================================================
Public Sub ExportAllForGitHub()

    Dim vbComp As VBIDE.VBComponent
    Dim compType As Long
    Dim fileName As String

    ' --- フォルダ存在チェック＆作成 ---
    EnsureFolder BASE_PATH
    EnsureFolder MODULES_PATH
    EnsureFolder SHEETS_PATH
    EnsureFolder FORMS_PATH
    EnsureFolder CLASSES_PATH
    EnsureFolder DOCS_PATH

    ' --- VBA コンポーネントのエクスポート ---
    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        compType = vbComp.Type

        Select Case compType

            Case vbext_ct_StdModule
                fileName = MODULES_PATH & vbComp.Name & ".bas"
                vbComp.Export fileName

            Case vbext_ct_ClassModule
                fileName = CLASSES_PATH & vbComp.Name & ".cls"
                vbComp.Export fileName

            Case vbext_ct_Document
                fileName = SHEETS_PATH & vbComp.Name & ".cls"
                vbComp.Export fileName

            Case vbext_ct_MSForm
                fileName = FORMS_PATH & vbComp.Name & ".frm"
                vbComp.Export fileName

        End Select

    Next vbComp

    ' --- 設計図シートの Markdown 出力 ---
    Call ExportSheetMarkdown(SHEET_SYSTEM, DOCS_PATH & "system_design.md")
    Call ExportSheetMarkdown(SHEET_DBMAP, DOCS_PATH & "db_mapping.md")
    Call ExportSheetMarkdown(SHEET_FUNC, DOCS_PATH & "functions.md")

    MsgBox "GitHub 用エクスポートが完了しました。" & vbCrLf & _
           "GitHub Desktop でコミットしてください。", vbInformation

End Sub


'===========================================================
' シートの A1 の内容を .md として書き出す
'===========================================================
Private Sub ExportSheetMarkdown(sheetName As String, outputPath As String)

    If SheetExists(sheetName) = False Then Exit Sub

    Dim f As Integer
    Dim text As String

    text = Sheets(sheetName).Range("A1").Value

    f = FreeFile
    Open outputPath For Output As #f
    Print #f, text
    Close #f

End Sub


'===========================================================
' フォルダがなければ作成
'===========================================================
Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub


'===========================================================
' シート存在チェック
'===========================================================
Private Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

