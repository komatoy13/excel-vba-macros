Attribute VB_Name = "Module_Export"
Option Explicit

Sub ExportModules_Toyomitsu()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim today As String
    Dim fileName As String
    Dim compType As Long
    
    ' GitHub Desktop が監視するフォルダ
    exportPath = "C:\Users\TOYOMITSU_DOUKE.KOMAYA\OneDrive\仕事\VBA\日報記録\repo\excel-vba-macros\"
    
    ' フォルダ存在チェック
    If Dir(exportPath, vbDirectory) = "" Then
        MsgBox "保存先フォルダが存在しません。" & vbCrLf & exportPath, vbExclamation
        Exit Sub
    End If
    
    ' 日付（YYYYMMDD）
    today = Format(Date, "yyyymmdd")
    
    ' すべてのコンポーネントをループ
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        compType = vbComp.Type
        
        ' 対象：標準モジュール（A）＋ ThisWorkbook/シート（D）
        If compType = vbext_ct_StdModule _
        Or compType = vbext_ct_Document Then
            
            ' ファイル名：ModuleName_YYYYMMDD.bas
            fileName = exportPath & vbComp.Name & "_" & today & ".bas"
            
            ' エクスポート
            On Error Resume Next
            vbComp.Export fileName
            If Err.Number <> 0 Then
                MsgBox "エクスポート失敗：" & vbComp.Name & vbCrLf & Err.Description, vbCritical
                Err.Clear
            End If
            On Error GoTo 0
            
        End If
    Next vbComp
    
    MsgBox "エクスポート完了しました！", vbInformation
End Sub

