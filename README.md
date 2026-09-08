# FNC Database Report Generator sample (FMX)

A free, source-included FireMonkey sample that turns a FireDAC query into a polished `TTMSFNCDataGrid` report and exports the current grid view to HTML, JSON, PDF, CSV, XLS, or—when FlexCel is enabled—XLSX.

[Download the packaged FNC Report Generator sample](https://www.tmssoftware.com/download/samples/FNCDBReportGenerator.zip)

## What the sample demonstrates

- A gear-button *Configure data source* popup (SQLite): pick a database file and choose a table.
- The last connection is remembered in an INI file and reloaded on startup.
- Live database binding through `TTMSFNCDataGridDatabaseAdapter`.
- DataGrid sorting, filtering, banding, selection, and conditional row styling.
- User-resizable columns (drag a column border, or double-click it to auto-fit) whose widths flow into the exports: XLS keeps the column widths, PDF (fit-to-page) keeps the proportions, and HTML gets a `<colgroup>` with percentage widths.
- A **View** bar to toggle column filtering, switch to advanced filtering, and turn on built-in conditional formatting (data bars, colour scales, top-N and value highlights) with a run-time rule editor.
- One-click export to five core formats, with optional FlexCel-powered XLSX.
- An `Output` folder beside the executable, keeping generated reports together with the demo build.

## Requirements

- Delphi with FireMonkey and FireDAC.
- TMS FNC Core and TMS FNC UI Pack with the DataGrid, PDF, and Excel IO units installed.
- Optional: TMS FlexCel VCL for genuine XLSX output.

The repository is source-only: executables, DCUs, generated databases, and exported reports are intentionally excluded. The PNG files are documentation screenshots and are not required by the application.

## Run it

1. Open `FNCDBReportGenerator.dproj`.
2. Select Win32 or Win64.
3. Choose `Debug` or `Release` for the core formats, or `FlexCel` to include XLSX support.
4. Build and run. On the first run there is no saved database, so click the gear button, browse to a SQLite file, and connect. From then on the app reloads the last database from `FNCDBReporterSettings.ini` on startup.

## Choose your own data

The **gear** button (top-right of the header) opens the *Configure data source* popup. The sample supports **SQLite** only:

- **Database file** — the SQLite file. Use **File...** to browse for a `.sqlite`/`.db` file; a relative path (e.g. `sample-sales.sqlite`) is resolved against the executable's folder, so the INI stays portable and always finds the file regardless of the working directory.
- **Table** — after connecting, this lists the database's tables (from `sqlite_master`) and the grid shows `select * from` the chosen one. If nothing is selected yet, the first table is used.
- **Connect / Run** opens the file and shows the table; **Close** hides the popup and **saves the database file and the selected table name** to `FNCDBReporterSettings.ini` next to the executable. The app loads that INI on startup and reconnects automatically, re-selecting the saved table; with no INI (first run, or after deleting it) it starts empty and waits for you to choose a file.

## View options

The **View** bar under the export buttons controls how the grid presents the data:

- **Filtering** — toggles `Grid.Options.Filtering.Enabled`. When off, the filter row is hidden.
- **Advanced filtering** — toggles `Grid.Options.Filtering.Advanced` for the richer, operator-based filter UI (enabled only while Filtering is on).
- **Conditional formatting** — toggles `Grid.ConditionalFormatting.Enabled` only; it shows or hides the rules that are currently loaded but never loads or clears them. Rules come from the JSON file on startup or via **Load settings**, or you add them with **Edit Rules...**.
- **Edit Rules...** — opens the grid's built-in conditional-formatting editor (`ShowConditionalFormattingEditor`) so rules can be added, removed, and tuned at run time. Edited rules are kept until the next data load.
- **Columns...** — opens a checklist of the columns in the current query. Tick or untick a column to show or hide it in the grid (`UnhideColumn` / `HideColumn`), reading the current state from `IsColumnHidden`. At least one column always stays visible, and exports honour the hidden columns.
- **Save current settings** — serializes the conditional-formatting rules to `FNCDBReporterFormatting.json` (`Grid.ConditionalFormatting.SaveToJSONStream`).
- The **filter is saved automatically when the form closes** (`OnClose`), in both forms so either mode is captured: the per-column rules to `FNCDBReporterFilter.json` (`Grid.Filter.SaveToJSONStream`) and the advanced expression + advanced flag to `FNCDBReporterFilterMode.ini` (`Grid.FilterBuilder.FilterText`).
- **Load settings** — opens a file dialog to pick a conditional-formatting `.json` file and loads it through `Grid.ConditionalFormatting.LoadFromJSONStream`. On startup (only) the app also auto-loads the default `FNCDBReporterFormatting.json` and restores the filter: it picks the mode that actually has data (advanced expression first, otherwise the per-column rules) and calls `ApplyFilter`.

`TTMSFNCDataGridExcelIO` writes the classic BIFF/XLS format, even if an `.xlsx` filename is supplied. The optional configuration therefore exports an intermediate XLS and uses the separate TMS FlexCel VCL product to save a genuine XLSX workbook. Without FlexCel, the application still builds and supports HTML, JSON, PDF, CSV, and XLS; the unavailable command is labeled `XLSX*` and explains the requirement.

FlexCel is a developer/build-machine dependency. When the application is built without runtime packages, users of the compiled executable do not need a separate FlexCel installation.

No local FlexCel path is stored in the project. Install FlexCel through your normal TMS setup and make its platform DCUs available through Delphi's library or project search path before selecting the `FlexCel` configuration.

## Architecture

- `TFDConnection` + `TFDQuery` own database access.
- `TTMSFNCDataGridDatabaseAdapter` maps the dataset to the DataGrid renderer.
- HTML and CSV use the renderer's native save methods.
- PDF uses `TTMSFNCDataGridPDFIO` with repeated headers and fit-to-page output.
- XLS uses `TTMSFNCDataGridExcelIO` and preserves cell appearance.
- XLSX is conditional: the built-in exporter creates an intermediate XLS, then FlexCel saves the real Office Open XML workbook and the intermediate file is removed.
- JSON walks the visible, filtered renderer rows so it represents the same report view.

Generated databases and reports are written to an `Output` folder beside the executable. Deploy the demo to a writable location when targeting platforms that protect application bundles.
