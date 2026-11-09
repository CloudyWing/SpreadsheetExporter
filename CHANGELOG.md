# Changelog

## v3.2.0 (2026-11-09)

### Behavior Changes

3.1 升級至 3.2 時可能觀察到下列行為變更，完整說明見 [Migration Notes](docs/articles/migration-notes.md#v31--v32)。

- 工作表有多個 templates 且包含 `RecordSetTemplate<T>` 或 `DataTableTemplate` 時，自動篩選的預設範圍改為第一個 `RecordSetTemplate<T>` 或 `DataTableTemplate` 的範圍，不再使用工作表已使用的範圍。只有單一 template 的工作表維持原行為。
- `CreateSheet` 產生唯一工作表名稱時改為不分大小寫比對，已有 `Sales` 時再建立 `SALES` 會得到 `SALES(1)`。
- `GetLayoutDiagnostics()` 與 `ValidateLayout()` 新增工作表名稱檢查，名稱不符合 Excel 規則時回報 `SE-LAYOUT-005`，名稱重複時回報 `SE-LAYOUT-006`。
- `RecordSetTemplate<T>` 的欄位以巢狀屬性路徑取值時，路徑中途的物件為 `null`，欄位值改為 `null`，先前會拋出 `ArgumentException`。`FromJson(...)` 建立的 RecordSet 不受影響，仍會拋出 `ArgumentException`。
- 清單資料驗證的項目含雙引號時自動跳脫，先前會產生無效的公式。清單來源超過 255 個字元時，ClosedXML renderer 會拋出說明長度與上限的 `ArgumentOutOfRangeException`。

### New Features

- 新增 Stream 與非同步匯出，`SpreadsheetDocument` 提供 `Export(Stream)`、`ExportAsync` 與 `ExportFileAsync`。非同步匯出先在記憶體完成 render 再寫入，失敗時不會寫入部分內容，也不會建立檔案。
- 新增 `ISpreadsheetStreamRenderer`，renderer 可直接將內容寫入 stream，未實作時退回 `byte[]` 輸出。`ExcelRenderer` 已實作此介面。
- 新增 `Export(ExportOptions)`，以 `ValidateLayout` 與 `ValidateCapabilities` 在匯出前驗證版面與 renderer 能力。版面與能力診斷合併為單一 `SpreadsheetLayoutException`，且不會呼叫 renderer。
- 新增 `BorderLineStyle` 與 `CellStyle.BorderStyle`，可指定外框線樣式。JSON style 與 JSON Schema 同步支援 `BorderStyle`。
- 新增 RecordSet 欄寬設定，`RecordSetColumnBase<T>.Width` 與 `RecordSetColumnCollection<T>.Add(..., double width)` 可指定欄寬。工作表以 `SetColumnWidth` 設定的欄寬優先，`MergedTemplate` 內的 RecordSet 不套用。
- 新增 `DataRowContext` 與 `DataTableColumn.FieldStyleContextGenerator`，DataTable 欄位樣式可取得整列資料。同時設定時優先於 `FieldStyleGenerator`。
- 新增 `SheetDefinition.AutoFilterRange`、`SetAutoFilterRange` 與 `SheetLayout.AutoFilterRange`，可明確指定自動篩選範圍。指定的範圍無效時，`GetLayoutDiagnostics()` 回報 `SE-LAYOUT-001`。
- 補上 `TryGetSheet` 的 `[NotNullWhen(true)]` 與 `GridTemplate.CreateCell` 的 `object?` nullable 標註。

### Bug Fixes

- 修正工作表上方有其他 templates 時，自動篩選套用到錯誤列的問題。
- 修正清單資料驗證項目含雙引號時產生無效公式的問題。
- 修正 `DataTableTemplate` 更換為欄位不同的 `DataTable` 後，匯出拋出例外的問題。更換後缺少既有欄位時會重建 `Columns`，資料列引用不存在的欄位時，改拋出列出可用欄位的 `ArgumentException`。
- 修正 `JsonTemplateRegistry` 與 `SpreadsheetManager` 並行存取時的執行緒安全問題。
- 修正 `DictionaryUtils` 遇到 indexer 與以 `new` 隱藏的屬性時拋出例外的問題。

### Performance

- 移除 `FieldColumn` 無限成長的快取，並以快取的成員存取器取值。
- `ExcelRenderer` 對單一儲存格套用樣式與資料驗證時不再建立 `IXLRange`。
- 模板組合表頭與資料儲存格改用 `Concat`，避免 `Union` 的雜湊去重。
- 快取 JSON 解析使用的已知紙張大小對照表。

### 文件

- 新增 NuGet 套件專用 README。
- README 補充 Stream、非同步與 `ExportOptions` 匯出用法。
- 補充目標框架、自動篩選範圍、清單驗證限制、巢狀欄位路徑為 `null`、欄寬限制、JSON 未支援項目與執行緒安全的說明。
- 新增 `SE-RENDERER-001` 診斷代碼說明與 v3.1 升級至 v3.2 的遷移說明。
- 修正文件範例錯誤，包含 `RecordContext.RowIndex` 的說明與公式範例列號、DI 範例缺少的 `using`。

### Build / CI

- 新增 `PublicApiAnalyzers`，並以 v3.1.0 的公開 API 建立 Shipped 基準。套件啟用 PackageValidation（以 3.1.0 為基準）、SourceLink 與 snupkg 符號套件。
- 發佈流程改為推送版本標籤時，驗證 nupkg 版號與標籤一致後建立 GitHub Release，並以 Trusted Publishing 發佈至 NuGet，不再使用 API Key。
- CI 改用 devops-templates 的 reusable workflow，並新增 ubuntu 建置矩陣。
- 文件網站改以 GitHub Pages artifact 部署。
- 新增 Dependabot 設定。

## v3.1.0 (2026-05-24)

### New Features

- 新增版面診斷與 Renderer 能力宣告，匯出前可檢查版面衝突與 renderer 支援狀態。
- 支援 JSON DataValidation 設定，可從 FromJson 流程建立資料驗證規則。
- 支援 JSON DataTable template，可透過 JSON 建立 DataTable 匯出模板。
- 支援 JSON Merged template，可透過 JSON 建立合併儲存格模板。
- 新增 JSON path diagnostics，JSON 解析錯誤可回報對應來源路徑。
- 新增 named style 與 JSON style inheritance，支援共用樣式與樣式繼承。
- 支援 JSON DataTable column value，可在 DataTable 欄位中指定固定值。
- 支援 DataTable 資料驗證，可在 DataTable 模板欄位設定驗證規則。
- 新增 FromJson JSON Schema，提供 JSON 模板格式驗證依據。

### Bug Fixes

- 修正模板欄位集合狀態與 DataTable 公式衝突防護。
- 補齊 FreezePanes JSON 型別診斷。
- 修正模板資料格可同時宣告值與公式的問題。

### 文件

- 補充 JSON 解析產生 Excel 章節。

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
