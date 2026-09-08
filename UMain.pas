unit UMain;

interface

uses
  System.SysUtils, System.UITypes, System.Classes, System.JSON,
  System.IOUtils, System.DateUtils, System.StrUtils, System.IniFiles,
  Data.DB,
  FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Def,
  FireDAC.Stan.ExprFuncs, FireDAC.Stan.Async, FireDAC.DApt,
  FireDAC.Comp.Client,
  FireDAC.Phys.SQLite,
  FireDAC.Phys.SQLiteDef, FireDAC.Phys.SQLiteWrapper.Stat,
  FireDAC.FMXUI.Wait,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Dialogs, FMX.StdCtrls,
  FMX.Edit, FMX.Memo, FMX.Objects, FMX.ListBox,
  FMX.TMSFNCTypes, FMX.TMSFNCUtils, FMX.TMSFNCGraphicsTypes,
  FMX.TMSFNCDataGridCell, FMX.TMSFNCDataGridData,
  FMX.TMSFNCDataGridBase, FMX.TMSFNCDataGridCore,
  FMX.TMSFNCDataGridRenderer, FMX.TMSFNCDataGrid,
  FMX.TMSFNCDataGridConditionalFormatting,
  FMX.TMSFNCDataGridDatabaseAdapter, FMX.TMSFNCPDFLib, FMX.TMSFNCPDFIO,
  FMX.TMSFNCDataGridPDFIO,
  FMX.TMSFNCDataGridExcelIO, FMX.TMSFNCGraphics, System.Rtti, FireDAC.Stan.Intf,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Pool,
  FireDAC.Phys, FireDAC.DatS, FireDAC.DApt.Intf, FMX.TMSFNCCustomComponent,
  FireDAC.Comp.DataSet, FMX.TMSFNCCustomControl, FMX.Controls.Presentation,
  FMX.Layouts
  {$IFDEF FNC_USE_FLEXCEL}
  , FlexCel.FMXSupport, FlexCel.Core, FlexCel.XlsAdapter
  {$ENDIF}
  ;

type
  TMainForm = class(TForm)
    Background: TRectangle;
    HeaderBackground: TRectangle;
    LabelBrand: TLabel;
    LabelTitle: TLabel;
    ButtonConfig: TButton;
    ExportBar: TRectangle;
    LabelExport: TLabel;
    ButtonHTML: TButton;
    ButtonJSON: TButton;
    ButtonPDF: TButton;
    ButtonCSV: TButton;
    ButtonXLS: TButton;
    ButtonXLSX: TButton;
    ViewBar: TRectangle;
    LabelView: TLabel;
    CheckFiltering: TCheckBox;
    CheckAdvancedFilter: TCheckBox;
    CheckConditionalFormat: TCheckBox;
    ButtonEditRules: TButton;
    ButtonColumns: TButton;
    ButtonSaveSettings: TButton;
    ButtonLoadSettings: TButton;
    ColumnsPanel: TRectangle;
    LabelColumns: TLabel;
    ColumnsListBox: TListBox;
    ButtonColumnsClose: TButton;
    ConnectionPanel: TRectangle;
    LabelConnTitle: TLabel;
    LabelDatabase: TLabel;
    EditDatabase: TEdit;
    ButtonConnDbBrowse: TButton;
    LabelTable: TLabel;
    ComboTable: TComboBox;
    ButtonConnect: TButton;
    ButtonConnCancel: TButton;
    Grid: TTMSFNCDataGrid;
    StatusBackground: TRectangle;
    LabelStatus: TLabel;
    LabelOutput: TLabel;
    OpenDialog: TOpenDialog;
    Connection: TFDConnection;
    Query: TFDQuery;
    DataSource: TDataSource;
    Adapter: TTMSFNCDataGridDatabaseAdapter;
    PDFIO: TTMSFNCDataGridPDFIO;
    ExcelIO: TTMSFNCDataGridExcelIO;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ExportClick(Sender: TObject);
    procedure CheckFilteringChange(Sender: TObject);
    procedure CheckAdvancedFilterChange(Sender: TObject);
    procedure CheckConditionalFormatChange(Sender: TObject);
    procedure ButtonEditRulesClick(Sender: TObject);
    procedure ButtonColumnsClick(Sender: TObject);
    procedure ButtonSaveSettingsClick(Sender: TObject);
    procedure ButtonLoadSettingsClick(Sender: TObject);
    procedure ButtonColumnsCloseClick(Sender: TObject);
    procedure ColumnsListBoxChangeCheck(Sender: TObject);
    procedure ButtonConfigClick(Sender: TObject);
    procedure ButtonConnectClick(Sender: TObject);
    procedure ButtonConnCancelClick(Sender: TObject);
    procedure ButtonConnDbBrowseClick(Sender: TObject);
    procedure ComboTableChange(Sender: TObject);
  private
    FOutputFolder: string;
    FInternalUpdate: Boolean;
    FInitialTable: string;
    procedure ConfigureOpenDialog;
    procedure ConnectAndRun;
    procedure ApplyViewOptions;
    function LoadFormattingFromFile(const AFileName: string): Boolean;
    procedure PopulateColumnsList;
    procedure PopulateTables;
    function ColumnWidthColGroup: string;
    function ResolveDatabaseFile(const APath: string): string;
    function RelativeDatabaseFile(const APath: string): string;
    procedure ExportCSV(const AFileName: string);
    procedure ExportHTML(const AFileName: string);
    procedure ExportJSON(const AFileName: string);
    procedure ExportPDF(const AFileName: string);
    procedure ExportXLS(const AFileName: string);
    procedure ExportXLSX(const AFileName: string);
    function OutputFile(const AExtension: string): string;
    procedure PrepareOutputFolder;
    function SettingsFileName: string;
    function FormattingFileName: string;
    function FilterFileName: string;
    function FilterModeFileName: string;
    procedure SaveSettings;
    procedure LoadSettings;
    procedure LoadCurrentSettings;
    procedure SaveFilter;
    procedure SetBusy(const AValue: Boolean; const AText: string = '');
    procedure SetStatus(const AText: string; const AIsError: Boolean = False);
    procedure StyleGrid;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.fmx}

type
  TReportFormat = (rfHTML, rfJSON, rfPDF, rfCSV, rfXLS, rfXLSX);

const
  COLOR_NAVY = $FF12233F;
  COLOR_RED = $FFD24E4E;

  REPORT_EXTENSIONS: array[TReportFormat] of string =
    ('html', 'json', 'pdf', 'csv', 'xls', 'xlsx');

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FInternalUpdate := True;
  try
    RegisterClasses([TTMSFNCDataGridFormatRuleItem, TTMSFNCDataGridDataFilterData]);

    PrepareOutputFolder;
    StyleGrid;
    ConfigureOpenDialog;
    {$IFNDEF FNC_USE_FLEXCEL}
    ButtonXLSX.Text := 'XLSX*';
    ButtonXLSX.Hint := 'XLSX export requires the optional FlexCel build configuration.';
    {$ENDIF}
    LoadSettings;   { restore the saved database file from the INI }
    if Trim(EditDatabase.Text) <> '' then
    begin
      ConnectAndRun;
      { Auto-restore conditional formatting / filter when their files exist. }
      if TFile.Exists(FormattingFileName) or TFile.Exists(FilterFileName) then
        LoadCurrentSettings;
    end
    else
      SetStatus('No database yet - click the gear button to choose a SQLite ' +
        'file.');
  finally
    FInternalUpdate := False;
  end;
end;

procedure TMainForm.ConfigureOpenDialog;
begin
  OpenDialog.Title := 'Select a database file';
  OpenDialog.Filter :=
    'SQLite databases (*.sqlite;*.db;*.s3db;*.db3)|*.sqlite;*.db;*.s3db;*.db3|' +
    'All files (*.*)|*.*';
end;

procedure TMainForm.PrepareOutputFolder;
begin
  FOutputFolder := System.IOUtils.TPath.Combine(ExtractFilePath(ParamStr(0)), 'Output');
  ForceDirectories(FOutputFolder);
  LabelOutput.Text := 'Output  ' + FOutputFolder;
end;

procedure TMainForm.SetBusy(const AValue: Boolean; const AText: string);
begin
  ExportBar.Enabled := not AValue;
  if AText <> '' then
    SetStatus(AText);
end;

procedure TMainForm.SetStatus(const AText: string; const AIsError: Boolean);
begin
  LabelStatus.Text := AText;
  if AIsError then
    LabelStatus.TextSettings.FontColor := COLOR_RED
  else
    LabelStatus.TextSettings.FontColor := COLOR_NAVY;
end;

procedure TMainForm.StyleGrid;
begin
  Grid.Options.Sorting.Enabled := True;
  Grid.Options.Filtering.Enabled := True;
  Grid.Options.Selection.Mode := gsmSingleRow;
  Grid.Options.Column.Stretching.Enabled := True;
  Grid.Options.Mouse.ColumnSizing := True;
  Grid.Options.Mouse.ColumnAutoSizeOnDblClick := True;
  Grid.CellAppearance.FixedLayout.Fill.Color := COLOR_NAVY;
  Grid.CellAppearance.FixedLayout.Font.Color := gcWhite;
  Grid.CellAppearance.FixedLayout.Font.Style := [TFontStyle.fsBold];
  Grid.CellAppearance.BandLayout.Fill.Color := $FFF3F7FF;
  Grid.CellAppearance.SelectedLayout.Fill.Color := $FFDCE8FF;
  Grid.GlobalFont.Scale := 1.05;
end;

procedure TMainForm.ApplyViewOptions;
begin
  Grid.Options.Filtering.Enabled := CheckFiltering.IsChecked;
  Grid.Options.Filtering.Advanced := CheckAdvancedFilter.IsChecked;
  CheckAdvancedFilter.Enabled := CheckFiltering.IsChecked;

  ButtonEditRules.Enabled := CheckConditionalFormat.IsChecked;
  Grid.ConditionalFormatting.Enabled := CheckConditionalFormat.IsChecked;

  Grid.Invalidate;
end;

function TMainForm.LoadFormattingFromFile(const AFileName: string): Boolean;
var
  Stream: TFileStream;
begin
  { Load the conditional-formatting rules from a JSON file, replacing the
    current rules. Returns False when the file is not present. }
  Grid.ConditionalFormatting.Enabled := False;

  Result := TFile.Exists(AFileName);
  if not Result then
    Exit;
  Grid.ConditionalFormatting.ClearRules;
  Stream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyWrite);
  try
    Grid.ConditionalFormatting.LoadFromJSONStream(Stream);
  finally
    Stream.Free;
  end;

  Grid.ConditionalFormatting.Enabled := True;
end;

procedure TMainForm.CheckFilteringChange(Sender: TObject);
begin
  ApplyViewOptions;
end;

procedure TMainForm.CheckAdvancedFilterChange(Sender: TObject);
begin
  ApplyViewOptions;
end;

procedure TMainForm.CheckConditionalFormatChange(Sender: TObject);
begin
  ApplyViewOptions;
end;

procedure TMainForm.ButtonEditRulesClick(Sender: TObject);
begin
  { The built-in editor lets the user add, remove and tune rules at run time.
    Keep the edited rules instead of reloading the demo set. }
  if Grid.Root.ShowConditionalFormattingEditor then
  begin
    FInternalUpdate := True;
    try
      CheckConditionalFormat.IsChecked := True;
    finally
      FInternalUpdate := False;
    end;
    Grid.ConditionalFormatting.Enabled := True;
    ButtonEditRules.Enabled := True;
    Grid.Invalidate;
  end;
end;

procedure TMainForm.PopulateColumnsList;
var
  I: Integer;
  Item: TListBoxItem;
  Caption: string;
begin
  { Fill the picker with one checkable entry per column of the current query. }
  FInternalUpdate := True;
  try
    ColumnsListBox.Clear;
    for I := 0 to Grid.ColumnCount - 1 do
    begin
      Caption := Trim(TTMSFNCDataGridData.ValueToString(Grid.Root.Cells[I, 0]));
      if Caption = '' then
        Caption := 'Column ' + IntToStr(I + 1);
      Item := TListBoxItem.Create(ColumnsListBox);
      Item.Parent := ColumnsListBox;
      Item.Text := Caption;
      Item.Tag := I;
      Item.IsChecked := not Grid.IsColumnHidden(I);
    end;
  finally
    FInternalUpdate := False;
  end;
end;

procedure TMainForm.ButtonColumnsClick(Sender: TObject);
begin
  if Grid.ColumnCount = 0 then
  begin
    SetStatus('Run a query before choosing columns.', True);
    Exit;
  end;
  PopulateColumnsList;
  ColumnsPanel.Visible := True;
  ColumnsPanel.BringToFront;
end;

procedure TMainForm.ButtonColumnsCloseClick(Sender: TObject);
begin
  ColumnsPanel.Visible := False;
end;

procedure TMainForm.ColumnsListBoxChangeCheck(Sender: TObject);
var
  Item: TListBoxItem;
  Column, I, VisibleCount: Integer;
begin
  if FInternalUpdate or not (Sender is TListBoxItem) then
    Exit;

  Item := TListBoxItem(Sender);
  Column := Item.Tag;

  if Item.IsChecked then
    Grid.UnhideColumn(Column)
  else
  begin
    { Keep at least one column visible. }
    VisibleCount := 0;
    for I := 0 to Grid.ColumnCount - 1 do
      if not Grid.IsColumnHidden(I) then
        Inc(VisibleCount);
    if VisibleCount <= 1 then
    begin
      FInternalUpdate := True;
      try
        Item.IsChecked := True;
      finally
        FInternalUpdate := False;
      end;
      Exit;
    end;
    Grid.HideColumn(Column);
  end;

  Grid.Invalidate;
end;

procedure TMainForm.ButtonConfigClick(Sender: TObject);
begin
  { The gear button opens the data-source configuration popup. }
  ConnectionPanel.Visible := True;
  ConnectionPanel.BringToFront;
end;

procedure TMainForm.ButtonConnCancelClick(Sender: TObject);
begin
  { Persist the database file (and SQL) as the window closes. }
  SaveSettings;
  ConnectionPanel.Visible := False;
end;

function TMainForm.SettingsFileName: string;
begin
  Result := System.IOUtils.TPath.Combine(
    ExtractFilePath(ParamStr(0)), 'FNCDBReporterSettings.ini');
end;

procedure TMainForm.SaveSettings;
var
  Ini: TMemIniFile;
begin
  Ini := TMemIniFile.Create(SettingsFileName);
  try
    Ini.WriteString('Connection', 'Database', EditDatabase.Text);
    if ComboTable.ItemIndex >= 0 then
      Ini.WriteString('Connection', 'Table',
        ComboTable.Items[ComboTable.ItemIndex])
    else
      Ini.WriteString('Connection', 'Table', '');
    Ini.UpdateFile;
  finally
    Ini.Free;
  end;
end;

procedure TMainForm.LoadSettings;
var
  Ini: TMemIniFile;
begin
  if not TFile.Exists(SettingsFileName) then
    Exit;
  Ini := TMemIniFile.Create(SettingsFileName);
  try
    EditDatabase.Text := Ini.ReadString('Connection', 'Database', EditDatabase.Text);
    FInitialTable := Ini.ReadString('Connection', 'Table', '');
  finally
    Ini.Free;
  end;
end;

function TMainForm.FormattingFileName: string;
begin
  Result := System.IOUtils.TPath.Combine(
    ExtractFilePath(ParamStr(0)), 'FNCDBReporterFormatting.json');
end;

function TMainForm.FilterFileName: string;
begin
  Result := System.IOUtils.TPath.Combine(
    ExtractFilePath(ParamStr(0)), 'FNCDBReporterFilter.json');
end;

function TMainForm.FilterModeFileName: string;
begin
  Result := System.IOUtils.TPath.Combine(
    ExtractFilePath(ParamStr(0)), 'FNCDBReporterFilterMode.ini');
end;

procedure TMainForm.SaveFilter;
var
  Stream: TFileStream;
  Ini: TMemIniFile;
begin
  { The active filter lives in Grid.Filter (normal per-column) or in the
    FilterBuilder (advanced). Persist both so whichever is in use is captured. }
  Stream := TFileStream.Create(FilterFileName, fmCreate);
  try
    Grid.Filter.SaveToJSONStream(Stream);
  finally
    Stream.Free;
  end;

  Ini := TMemIniFile.Create(FilterModeFileName);
  try
    Ini.WriteBool('Filter', 'Advanced', Grid.Options.Filtering.Advanced);
    Ini.WriteString('Filter', 'Expression', Grid.FilterBuilder.FilterText);
    Ini.UpdateFile;
  finally
    Ini.Free;
  end;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  { Persist the current filter automatically as the form closes. }
  SaveFilter;
end;

procedure TMainForm.ButtonSaveSettingsClick(Sender: TObject);
var
  Stream: TFileStream;
begin
  { Conditional-formatting rules -> JSON. (The filter is saved on form close.) }
  Stream := TFileStream.Create(FormattingFileName, fmCreate);
  try
    Grid.ConditionalFormatting.SaveToJSONStream(Stream);
  finally
    Stream.Free;
  end;

  SetStatus('Saved conditional formatting to FNCDBReporterFormatting.json');
end;

procedure TMainForm.LoadCurrentSettings;
var
  Ini: TMemIniFile;
  Stream: TFileStream;
  Expression: string;
  Saved, Advanced: Boolean;
begin
  { Restore the conditional formatting and the filter (per-column rules plus the
    advanced expression) that a previous "Save current settings" wrote. }
  Saved := FInternalUpdate;
  FInternalUpdate := True;   { suppress check-box handlers while restoring }
  try
    if LoadFormattingFromFile(FormattingFileName) then
    begin
      CheckConditionalFormat.IsChecked := True;
      ButtonEditRules.Enabled := True;
    end;

    Advanced := True;
    Expression := '';
    if TFile.Exists(FilterModeFileName) then
    begin
      Ini := TMemIniFile.Create(FilterModeFileName);
      try
        Advanced := Ini.ReadBool('Filter', 'Advanced', True);
        Expression := Ini.ReadString('Filter', 'Expression', '');
      finally
        Ini.Free;
      end;
    end;

    if TFile.Exists(FilterFileName) then
    begin
      Grid.Filter.Clear;
      Stream := TFileStream.Create(FilterFileName, fmOpenRead or fmShareDenyWrite);
      try
        Grid.Filter.LoadFromJSONStream(Stream);
      finally
        Stream.Free;
      end;
    end;

    if TFile.Exists(FilterFileName) or TFile.Exists(FilterModeFileName) then
    begin
      { Use whichever representation actually holds a filter: an advanced
        expression takes precedence, otherwise the per-column rules. }
      if Trim(Expression) <> '' then
        Advanced := True
      else if Grid.Filter.Count > 0 then
        Advanced := False;

      CheckFiltering.IsChecked := True;
      CheckAdvancedFilter.IsChecked := Advanced;
      Grid.Options.Filtering.Enabled := True;
      Grid.Options.Filtering.Advanced := Advanced;
      try
        Grid.FilterBuilder.FilterText := Expression;
        Grid.RemoveFilter;
        if (Trim(Expression) <> '') or (Grid.Filter.Count > 0) then
          Grid.ApplyFilter;
      except
        { a saved filter that no longer fits the data is non-fatal }
      end;
    end;
  finally
    FInternalUpdate := Saved;
  end;
  Grid.Invalidate;
end;

procedure TMainForm.ButtonLoadSettingsClick(Sender: TObject);
var
  Dialog: TOpenDialog;
begin
  { Let the user pick a conditional-formatting JSON file to load. }
  Dialog := TOpenDialog.Create(Self);
  try
    Dialog.Title := 'Load conditional formatting';
    Dialog.Filter := 'Conditional formatting (*.json)|*.json|All files (*.*)|*.*';
    if TFile.Exists(FormattingFileName) then
      Dialog.FileName := FormattingFileName;
    if not Dialog.Execute then
      Exit;

    FInternalUpdate := True;
    try
      if LoadFormattingFromFile(Dialog.FileName) then
      begin
        CheckConditionalFormat.IsChecked := True;
        ButtonEditRules.Enabled := True;
        SetStatus('Loaded conditional formatting from ' +
          ExtractFileName(Dialog.FileName));
      end
      else
        SetStatus('File not found: ' + Dialog.FileName, True);
    finally
      FInternalUpdate := False;
    end;
    Grid.Invalidate;
  finally
    Dialog.Free;
  end;
end;

procedure TMainForm.ButtonConnDbBrowseClick(Sender: TObject);
begin
  if OpenDialog.Execute then
  begin
    EditDatabase.Text := RelativeDatabaseFile(OpenDialog.FileName);
    FInitialTable := '';  { new database - start at its first table }
    ConnectAndRun;
  end;
end;

procedure TMainForm.ComboTableChange(Sender: TObject);
begin
  { Picking a table shows it immediately. }
  if FInternalUpdate or (ComboTable.ItemIndex < 0) then
    Exit;
  ConnectAndRun;
end;

procedure TMainForm.PopulateTables;
var
  Saved: Boolean;
  ListQuery: TFDQuery;
  PrevSel: string;
begin
  { List the SQLite database's tables so the user can pick one instead of
    typing SQL. }
  Saved := FInternalUpdate;
  FInternalUpdate := True;
  try
    { Keep the current selection across a refresh; on the first load fall back
      to the table name restored from the INI. }
    if ComboTable.ItemIndex >= 0 then
      PrevSel := ComboTable.Items[ComboTable.ItemIndex]
    else
      PrevSel := FInitialTable;
    ComboTable.Items.Clear;

    if Connection.Connected then
    begin
      ListQuery := TFDQuery.Create(nil);
      try
        ListQuery.Connection := Connection;
        try
          ListQuery.SQL.Text :=
            'select name from sqlite_master where type = ''table'' ' +
            'and name not like ''sqlite_%'' order by name';
          ListQuery.Open;
          while not ListQuery.Eof do
          begin
            ComboTable.Items.Add(ListQuery.Fields[0].AsString);
            ListQuery.Next;
          end;
        except
          { enumeration is best-effort; ignore failures }
        end;
      finally
        ListQuery.Free;
      end;
    end;

    ComboTable.Enabled := ComboTable.Items.Count > 0;
    if PrevSel <> '' then
      ComboTable.ItemIndex := ComboTable.Items.IndexOf(PrevSel);
  finally
    FInternalUpdate := Saved;
  end;
end;

procedure TMainForm.ButtonConnectClick(Sender: TObject);
begin
  { Requires a SQLite database file. The popup stays open so the listed tables
    can be picked after connecting. }
  if Trim(EditDatabase.Text) = '' then
  begin
    SetStatus('Select a SQLite database file first.', True);
    Exit;
  end;
  ConnectAndRun;
end;

function TMainForm.ResolveDatabaseFile(const APath: string): string;
begin
  { A relative SQLite file is resolved against the executable's folder so the
    saved INI stays portable and connects regardless of the working directory. }
  Result := APath;
  if (APath = '') or (APath[1] = ':') or System.IOUtils.TPath.IsPathRooted(APath) then
    Exit;
  Result := System.IOUtils.TPath.Combine(ExtractFilePath(ParamStr(0)), APath);
end;

function TMainForm.RelativeDatabaseFile(const APath: string): string;
var
  Base: string;
begin
  { Store a file that lives in the app folder as just its relative name, so the
    INI stays portable; files elsewhere keep their absolute path. }
  Base := ExtractFilePath(ParamStr(0));
  if APath.StartsWith(Base, True) then
    Result := APath.Substring(Base.Length)
  else
    Result := APath;
end;

procedure TMainForm.ConnectAndRun;
var
  I: Integer;
  DbFile: string;
begin
  DbFile := Trim(EditDatabase.Text);
  if DbFile = '' then
  begin
    SetStatus('Select a SQLite database file (gear button).', True);
    Exit;
  end;

  SetBusy(True, 'Opening database and running query...');
  try
    try
      Adapter.Active := False;
      Query.Close;
      Connection.Connected := False;
      Connection.ConnectionString := 'DriverID=SQLite;Database=' +
        ResolveDatabaseFile(DbFile);
      Connection.Connected := True;

      { List the tables, then show the selected one (or the first table). }
      PopulateTables;
      if (ComboTable.ItemIndex < 0) and (ComboTable.Items.Count > 0) then
      begin
        FInternalUpdate := True;
        try
          ComboTable.ItemIndex := 0;
        finally
          FInternalUpdate := False;
        end;
      end;
      if ComboTable.ItemIndex < 0 then
      begin
        SetStatus('Connected, but this database has no tables.', True);
        Exit;
      end;

      Query.FetchOptions.Mode := fmAll;
      Query.SQL.Text := 'select * from "' +
        ComboTable.Items[ComboTable.ItemIndex] + '"';
      Query.Open;

      Adapter.DataSource := DataSource;
      Adapter.Renderer := Grid.Root;
      Adapter.AutoCreateColumns := True;
      Adapter.LoadMode := almAllRecords;
      Adapter.Active := True;

      for I := 0 to Grid.ColumnCount - 1 do
        Grid.ColumnWidths[I] := 130;
      if Grid.ColumnCount > 2 then
        Grid.ColumnWidths[2] := 180;
      if Grid.ColumnCount > 4 then
        Grid.ColumnWidths[4] := 190;

      ApplyViewOptions;
      if ColumnsPanel.Visible then
        PopulateColumnsList;

      SetStatus(Format(
        '%s records ready  |  Sort and filter in the grid, then export the current view.',
        [FormatFloat('#,##0', Query.RecordCount)]));
    except
      on E: Exception do
        SetStatus('Could not open the database or run the query: ' + E.Message,
          True);
    end;
  finally
    SetBusy(False);
  end;
end;

function TMainForm.OutputFile(const AExtension: string): string;
begin
  Result := System.IOUtils.TPath.Combine(FOutputFolder,
    'sales-report-' + FormatDateTime('yyyymmdd-hhnnss', Now) + '.' + AExtension);
end;

function TMainForm.ColumnWidthColGroup: string;
var
  I: Integer;
  Total, W: Double;
  Builder: TStringBuilder;
begin
  { Build an HTML <colgroup> whose columns carry the grid's current widths as
    percentages, so the exported table keeps the on-screen proportions. }
  Result := '';
  Total := 0;
  for I := 0 to Grid.ColumnCount - 1 do
    if not Grid.IsColumnHidden(I) then
      Total := Total + Grid.ColumnWidths[I];
  if Total <= 0 then
    Exit;

  Builder := TStringBuilder.Create;
  try
    Builder.Append('<colgroup>');
    for I := 0 to Grid.ColumnCount - 1 do
      if not Grid.IsColumnHidden(I) then
      begin
        W := Grid.ColumnWidths[I] / Total * 100;
        Builder.Append(Format('<col style="width:%.2f%%">', [W],
          TFormatSettings.Invariant));
      end;
    Builder.Append('</colgroup>');
    Result := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

procedure TMainForm.ExportHTML(const AFileName: string);
var
  HTML, LowerHTML, ColGroup: string;
  TagStart, TagEnd: Integer;
begin
  Grid.Root.SaveToHTMLData(AFileName, TEncoding.UTF8);

  ColGroup := ColumnWidthColGroup;
  if ColGroup = '' then
    Exit;

  { Insert the column widths right after the opening <table ...> tag. }
  HTML := TFile.ReadAllText(AFileName, TEncoding.UTF8);
  LowerHTML := LowerCase(HTML);
  TagStart := Pos('<table', LowerHTML);
  if TagStart > 0 then
  begin
    TagEnd := PosEx('>', LowerHTML, TagStart);
    if TagEnd > 0 then
    begin
      Insert(ColGroup, HTML, TagEnd + 1);
      TFile.WriteAllText(AFileName, HTML, TEncoding.UTF8);
    end;
  end;
end;

procedure TMainForm.ExportCSV(const AFileName: string);
begin
  Grid.Root.Options.IO.Delimiter := ',';
  Grid.Root.Options.IO.QuoteEmptyCells := True;
  Grid.Root.SaveToCSVData(AFileName, TEncoding.UTF8);
end;

procedure TMainForm.ExportPDF(const AFileName: string);
begin
  PDFIO.DataGrid := Grid;
  PDFIO.Options.Header := 'SALES PERFORMANCE REPORT';
  PDFIO.Options.Footer := 'Generated ' + FormatDateTime('dd mmm yyyy, hh:nn', Now);
  PDFIO.Options.PageNumber := pnFooter;
  PDFIO.Options.PageNumberFormat := 'Page %d';
  PDFIO.Options.FitToPage := True;
  PDFIO.Options.RepeatFixedRows := True;
  PDFIO.Options.OpenInPDFReader := False;
  PDFIO.Save(AFileName);
end;

procedure TMainForm.ExportXLS(const AFileName: string);
begin
  ExcelIO.Renderer := Grid.Root;
  ExcelIO.DataGridStartRow := 0;
  ExcelIO.DataGridStartCol := 0;
  ExcelIO.Options.ExportOverwrite := omAlways;
  ExcelIO.Options.ExportCellProperties := True;
  ExcelIO.Options.ExportShowInExcel := False;
  ExcelIO.XLSExport(AFileName, 'Sales report');
end;

procedure TMainForm.ExportXLSX(const AFileName: string);
{$IFDEF FNC_USE_FLEXCEL}
var
  Workbook: TXlsFile;
  IntermediateXLS: string;
begin
  { DataGridExcelIO writes BIFF/XLS. FlexCel performs the real XLSX conversion. }
  IntermediateXLS := ChangeFileExt(AFileName, '.intermediate.xls');
  ExportXLS(IntermediateXLS);
  try
    Workbook := TXlsFile.Create(IntermediateXLS, True);
    try
      Workbook.Save(AFileName, TFileFormats.Xlsx);
    finally
      Workbook.Free;
    end;
  finally
    if TFile.Exists(IntermediateXLS) then
      TFile.Delete(IntermediateXLS);
  end;
end;
{$ELSE}
begin
  ShowMessage('XLSX export requires FlexCel. Build the project with the FlexCel configuration.');
end;
{$ENDIF}

procedure TMainForm.ExportJSON(const AFileName: string);
var
  Document: TJSONObject;
  Rows: TJSONArray;
  RowObject: TJSONObject;
  Renderer: TTMSFNCDataGridRenderer;
  Column, Row: Integer;
  Key, Value: string;
begin
  Renderer := Grid.Root;
  Document := TJSONObject.Create;
  try
    Document.AddPair('report', 'Sales performance report');
    Document.AddPair('generated_at', DateToISO8601(Now, False));
    Rows := TJSONArray.Create;
    Document.AddPair('rows', Rows);

    for Row := Renderer.FixedRowCount to Renderer.RowCount - 1 do
      if not Renderer.IsRowHidden(Row) and not Renderer.IsRowFiltered(Row) then
      begin
        RowObject := TJSONObject.Create;
        for Column := 0 to Renderer.ColumnCount - 1 do
          if not Renderer.IsColumnHidden(Column) then
          begin
            Key := TTMSFNCDataGridData.ValueToString(Renderer.Cells[Column, 0]);
            if Key = '' then
              Key := 'column_' + IntToStr(Column + 1);
            Value := TTMSFNCDataGridData.ValueToString(Renderer.Cells[Column, Row]);
            RowObject.AddPair(Key, Value);
          end;
        Rows.AddElement(RowObject);
      end;

    TFile.WriteAllText(AFileName, Document.Format(2), TEncoding.UTF8);
  finally
    Document.Free;
  end;
end;

procedure TMainForm.ExportClick(Sender: TObject);
var
  FileName, Extension: string;
  ReportFormat: TReportFormat;
begin
  if not Query.Active then
  begin
    SetStatus('Run a query before exporting.', True);
    Exit;
  end;

  ReportFormat := TReportFormat(TControl(Sender).Tag);
  Extension := REPORT_EXTENSIONS[ReportFormat];
  FileName := OutputFile(Extension);

  SetBusy(True, 'Creating ' + UpperCase(Extension) + ' report...');
  try
    try
      case ReportFormat of
        rfHTML: ExportHTML(FileName);
        rfJSON: ExportJSON(FileName);
        rfPDF: ExportPDF(FileName);
        rfCSV: ExportCSV(FileName);
        rfXLS: ExportXLS(FileName);
        rfXLSX: ExportXLSX(FileName);
      end;
      SetStatus('Created  ' + FileName);
      TTMSFNCUtils.OpenFile(FileName);
    except
      on E: Exception do
        SetStatus('Export failed: ' + E.Message, True);
    end;
  finally
    SetBusy(False);
  end;
end;


end.
