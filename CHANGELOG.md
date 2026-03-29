# Changelog

## v3.0.0 (2026-03-29)

### New Features

- 實作 v3 核心模型與命名：
  - `SpreadsheetDocument` 取代舊版匯出入口。
  - `SheetDefinition` / `SheetLayout` 取代 `Sheeter` / `SheeterContext`。
  - `TemplateLayout` 取代 `TemplateContext`；`GetLayout()` 取代 `GetContext()`。
  - `ISpreadsheetRenderer` 成為新的渲染抽象。
- 新增 `SpreadsheetExporter.Renderer.ClosedXML`（取代舊版 `Excel.EPPlus` / `Excel.NPOI`）：
  - 使用 ClosedXML 輸出 `.xlsx`。
  - 支援樣式、合併儲存格、欄寬/列高、凍結窗格、自動篩選、資料驗證與工作表保護。
  - 進入點：`SpreadsheetManager.SetRenderer(() => new ExcelRenderer())`。
- 更新目標框架：
  - 核心與 ClosedXML renderer 套件改為 `net8.0` 與 `net10.0`。
- 新增 JSON 模板流程：
  - `SpreadsheetDocument.FromJson(...)` 可解析 `GridTemplate` 與 `RecordSetTemplate`。
  - 支援以 JSON 描述欄寬、頁面設定、模板與資料列。
- 改善 `RecordSetColumnCollection<T>` 階層欄位 API：
  - 新增 `GetLastColumn()` 方法，回傳最後一個欄位物件。
  - 新增 `Parent` 屬性，從子集合導覽回上層集合。
  - 新增 `RootColumns` 屬性，取得根集合。
  - 移除舊版 `AddChildToLastColumn(...)` 系列多載。
- 新增 `CellGenerator<out T>` delegate，取代 Cell 屬性上的 `Func<int, int, T>`，語義更明確。
- 更新 sample 與 benchmark 專案：
  - sample 改為展示 v3 + ClosedXML 的 Fluent API 與 JSON 基本用法。
  - benchmarks 改為量測 v3 模板建立與 ClosedXML 匯出流程。

### BREAKING CHANGES

- 移除舊版 `SpreadsheetExporter.Excel.EPPlus` 與 `SpreadsheetExporter.Excel.NPOI` solution 專案。
- `Sheeter`、`SheeterContext`、`ITemplate`、`ISpreadsheetExporter` 等 v2 API 已改為 v3 命名。
- `TemplateContext` 已重新命名為 `TemplateLayout`；所有模板的 `GetContext()` 已改為 `GetLayout()`。
- Renderer 套件改為 `CloudyWing.SpreadsheetExporter.Renderer.ClosedXML`，進入類別改為 `ExcelRenderer`。
- 核心與 renderer 套件不再支援 v2 時期的舊目標框架組合。
- `FreezePanes` 與 `IsAutoFilterEnabled` 相關設定已從 `RecordSetTemplate` 移至 `SheetDefinition` 。
- `RecordSetColumnCollection<T>` 的 `AddChildToLastColumn(...)` 系列多載已移除，請改用 `GetLastColumn().ChildColumns.Add(...)`。
- `Cell` 的四個 generator 屬性型別從 `Func<int, int, T>` 改為 `CellGenerator<T>`；現有 lambda 指派無需修改，但若有明確型別宣告需更新。

## v2.3.0 (2026-02-01)

### New Features

- 升級目標框架至 .NET 10 並重構文檔系統：
  - 將目標框架更新為 `net10.0` 和 `netstandard2.0`。
  - 遷移文檔系統從 DefaultDocumentation 至 DocFX。
  - 重構文檔結構，檔名英文化，內容保持繁體中文。
  - 優化文檔內容品質，套用 DocFX 最佳實踐（alerts、表格、改進章節層級）。
  - 更新導航結構並修正跨文件引用。

### BREAKING CHANGES

- 目標框架從 `net6.0` 升級至 `net10.0`。需要 .NET 10 SDK 才能建置專案。

> 注意：此次更新與之前版本不相容，請在升級前謹慎考慮。

## v2.2.0 (2026-01-21)

### New Features

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

### Bug Fixes

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
