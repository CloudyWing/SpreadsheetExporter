# v2.2.0 (2026-01-21)

## New Features

- 新增 Sample 專案展示所有 SpreadsheetExporter 功能。
- 新增 Data Validation 功能支援。
- 新增 Freeze Panes 和 AutoFilter 功能：
  - TemplateContext 新增 FreezePanes 和 IsEnabledAutoFilter 屬性。
  - RecordSetTemplate 新增 IsFreezeHeader 和 IsEnabledAutoFilter 屬性。
  - SheeterContext 新增對應屬性以傳遞設定。
  - EPPlus 和 NPOI 匯出器實作相關功能。
- 新增 DataTableTemplate 支援 System.Data.DataTable：
  - DataTableTemplate: 主要範本類別。
  - DataTableColumn: 欄位定義類別。
  - DataTableColumnCollection: 欄位集合管理類別。
  - 支援自動從 DataTable schema 產生欄位。
- 新增 BenchmarkDotNet 效能基準測試專案（效能優化）：
  - 建立 benchmarks/SpreadsheetExporter.Benchmarks 專案，用於評估不同元件的效能表現。
  - TemplateBenchmarks: 比較 RecordSetTemplate、DataTableTemplate 和 GridTemplate 的效能。
  - ExporterBenchmarks: 比較 NPOI 和 EPPlus 的匯出效能（基本匯出和包含樣式）。
  - 測試資料量：100、1,000、10,000 列。
  - 記憶體診斷：包含 GC 回收次數和記憶體配置量分析。
- 優化 RecordSetTemplate 效能並改善 SpreadsheetManager 執行緒安全（效能優化）：
  - 修正多次列舉問題和不必要的鎖定，提升效能和執行緒安全性。
- 優化 GridTemplate 儲存格位置查找效能（效能優化）：
  - 在 CellCollection 新增 nextAvailableX 追蹤下一個可用 X 座標。
  - 每次新增儲存格後更新 nextAvailableX 為當前儲存格結束位置。
  - 減少 HashSet 查找次數，預期可減少 20-50% 位置查找時間。

## Bug Fixes

- 改善 NPOI ExcelExporter 資源管理和例外處理。
- 重構 DataColumnCollection 減少程式碼重複：
  - 提取 ApplyDefaultStyles 輔助方法處理共同初始化邏輯。
  - 簡化 13 個公開方法的實作。
  - 改善程式碼可維護性，保持完全向後相容。
- 將測試框架從 FluentAssertions 遷移到 NUnit Assert：
  - 移除 FluentAssertions 套件依賴。
  - 改用 NUnit 內建的 Assert.That() 和 Assert.Throws() 語法。
  - 修正物件比較問題。
- 優化例外處理與型別安全，並統一欄寬設定：
  - 新增 Excel 欄寬單位常數，統一欄寬設定方式。
  - 強化型別轉換、合併格衝突等例外處理與錯誤訊息。
  - 實作 CellStyle 預設字型集中管理，並將 DefaultCellStyles 改為 Lazy 初始化。
  - 修正常數拼字 (AutoFiteRowHeight -> AutoFitRowHeight) 並保留舊屬性以維持相容性。
  - 改善 Null 判斷與多型覆寫機制，並於取值失敗時列出可用屬性以提升除錯效率。
  - 補充與修正多處 XML API 文件註解。
- 新增 SpreadsheetManager 單元測試和 ExcelVerifier 輔助工具。
