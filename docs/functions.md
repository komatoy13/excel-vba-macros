# 関数一覧（functions.md / Version 1.0）

このドキュメントは、日報記録システムに含まれる  
「関数・Sub・ロジックの一覧」と「役割」「依存関係」をまとめたものです。

GitHub に保存された .bas ファイルと合わせて参照することで、  
コードの整合性を保ち、拡張時の事故を防ぐ目的で作成しています。

---

# 1. モジュール構成（現行）
modules/ Module_DB.bas Module_Input.bas Module_Product.bas Module_HalfProduct.bas Module_Time.bas Module_Yield.bas Module_Utils.bas Module_History.bas
classes/ （今は空）
forms/ （今は空）
sheets/ Sheet1.cls Sheet2.cls ThisWorkbook.cls

---

# 2. 関数一覧（モジュール別）

以下は **現時点のコード構成を前提にした一覧** です。  
今後コードが増えたら Version を更新していきます。

---

## ■ Module_DB（DB 登録・修正・未完了管理）

### ● Sub RegisterToDB()
- **役割**：Input シートの内容を DB に新規登録する  
- **呼び出し元**：登録ボタン  
- **依存**：  
  - BuildProcessString（商品工程）  
  - BuildWorkerString（作業者）  
  - EvaluateYield（歩留まり）  
  - EvaluateTime（主作業時間）  

---

### ● Sub RegisterIncomplete()
- **役割**：未完了フラグを付けて DB に登録  
- **呼び出し元**：未完了登録ボタン  
- **依存**：RegisterToDB の一部ロジック

---

### ● Sub RegisterCorrection()
- **役割**：修正モードで DB の既存行を上書き  
- **呼び出し元**：修正登録ボタン  
- **依存**：Z1（DB 行番号）

---

### ● Function FindLastRow()
- **役割**：DB の最終行を取得  
- **呼び出し元**：RegisterToDB など  
- **依存**：なし

---

## ■ Module_Input（入力制御・切替・復元）

### ● Sub SwitchMode()
- **役割**：商品／半製品の入力項目を切り替える  
- **呼び出し元**：B3 の変更イベント  
- **依存**：なし

---

### ● Sub RestoreFromHistory()
- **役割**：履歴から Input に復元  
- **呼び出し元**：修正ボタン  
- **依存**：History モジュール

---

### ● Sub RestoreFromIncomplete()
- **役割**：未完了呼び出し  
- **呼び出し元**：未完了一覧のボタン  
- **依存**：DB 行番号

---

## ■ Module_Product（商品ロジック）

### ● Function BuildProcessString()
- **役割**：商品工程チェックボックスを文字列化  
- **呼び出し元**：RegisterToDB, SaveHistory  
- **依存**：なし

---

### ● Function BuildWorkerString()
- **役割**：作業者チェックボックスを文字列化  
- **呼び出し元**：RegisterToDB, SaveHistory  
- **依存**：なし

---

## ■ Module_HalfProduct（半製品ロジック）

### ● Function EvaluateYield()
- **役割**：歩留まり計算  
- **呼び出し元**：RegisterToDB  
- **依存**：数量・枚数・平均重量

---

### ● Function EvaluateTime()
- **役割**：主作業時間の標準比較  
- **呼び出し元**：RegisterToDB  
- **依存**：標準時間テーブル（予定）

---

## ■ Module_Time（時間関連）

### ● Function FormatTime()
- **役割**：時間を mm:ss 形式に整形  
- **呼び出し元**：履歴保存など  
- **依存**：なし

---

## ■ Module_Yield（歩留まり関連）

### ● Function FormatYield()
- **役割**：歩留まりを「95%」形式に整形  
- **呼び出し元**：履歴保存  
- **依存**：なし

---

## ■ Module_Utils（共通処理）

### ● Function Nz()
- **役割**：Null/空欄を 0 に変換  
- **呼び出し元**：全体  
- **依存**：なし

---

### ● Function SafeValue()
- **役割**：エラー値を安全に数値化  
- **呼び出し元**：DB登録  
- **依存**：なし

---

## ■ Module_History（履歴スライド）

### ● Sub SaveHistory()
- **役割**：最新入力を F列に保存し、G〜O をスライド  
- **呼び出し元**：登録時  
- **依存**：  
  - BuildProcessString  
  - BuildWorkerString  
  - FormatYield  
  - FormatTime  

---

### ● Sub SlideHistory()
- **役割**：履歴を右にずらす  
- **呼び出し元**：SaveHistory  
- **依存**：なし

---

# 3. 依存関係マップ（簡易）
RegisterToDB ├─ BuildProcessString ├─ BuildWorkerString ├─ EvaluateYield（半製品） ├─ EvaluateTime（半製品） └─ SaveHistory ├─ BuildProcessString ├─ BuildWorkerString ├─ FormatYield └─ FormatTime

---

# 4. 今後追加予定の関数（メモ）

- ValidateInput（入力チェック）
- SearchDB（検索機能）
- ExportCSV（外部出力）
- ImportMaster（マスタ読込）
- StandardTimeLookup（標準時間テーブル）

---

# 5. 更新履歴

- Version 1.0：初版作成（2026-01-28）
