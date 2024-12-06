unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.Menus, Vcl.Buttons,
  System.ImageList, Vcl.ImgList, Vcl.ExtCtrls, Vcl.Grids, IniFiles,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, UDataModule, PrintSetupForm,
  Vcl.WinXCtrls, PreviewForm,
  Data.DB, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  ExcelXP, Vcl.Imaging.pngimage, System.Generics.Collections,
  ComObj, Printers, Winapi.WinSpool, System.StrUtils, Clipbrd, System.IOUtils;

type
  TFormMain = class(TForm)
    clbToolButton: TCoolBar;
    panTitle: TPanel;
    labTitle: TLabel;
    panHeader: TPanel;
    tlbToolButton: TToolBar;
    mamMain: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    N1: TMenuItem;
    Printout1: TMenuItem;
    Setting1: TMenuItem;
    btnClose: TSpeedButton;
    imlToolBar: TImageList;
    btnPreview: TSpeedButton;
    edtbarcode: TEdit;
    lblBarCode: TLabel;
    stbBase: TStatusBar;
    stgMain: TStringGrid;
    MainMemTable: TFDMemTable;
    btnDelete: TSpeedButton;
    btnDeleteSelection: TSpeedButton;
    btnDeleteAll: TSpeedButton;
    Panel1: TPanel;
    btnSelectAll: TSpeedButton;
    btnDeSelectAll: TSpeedButton;
    btnPrintOut: TSpeedButton;
    btnPrintSetup: TSpeedButton;
    PrintDialog1: TPrintDialog;
    ProgressBar1: TProgressBar;
    procedure Setting1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject); // Added to handle cleanup
    procedure btnCloseClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure SetupStringGrid;
    procedure stgMainDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure stgMainMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnSelectAllClick(Sender: TObject);
    procedure btnDeSelectAllClick(Sender: TObject);
    procedure ProcessBarcode;
    procedure edtbarcodeKeyPress(Sender: TObject; var Key: Char);
    procedure btnPrintOutClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnDeleteSelectionClick(Sender: TObject);
    procedure btnDeleteAllClick(Sender: TObject);
    procedure btnPrintSetupClick(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Printout1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ClearGridData;
    procedure stgMainSelectCell(Sender: TObject; ACol, ARow: LongInt;
      var CanSelect: Boolean);
    procedure AdjustColumnWidths;
    procedure FormResize(Sender: TObject);

  private
    FDraggingColumn,EM: Integer; // Column currently being dragged
    PageCount: Integer;
    FirstPage,LastPage : Integer;
    FDragStartX: Integer; // Initial X position when the drag started
    FDragTargetColumn: Integer; // Target column to move to
    ColumnMapping: array of Integer;
    ClearAll:Boolean;
    // Array to track the display order of columns
    NextRow: Integer;
    DataModuleCIMT: TDataModuleCIMT; // Instance of the data module
    FIsUseRESTAPIState: TToggleSwitchState;
    FDebugSQL: TToggleSwitchState;
    FUser: string;
    FPass: string;
    procedure ReadIniFile;
    procedure WriteIniFile;
    procedure InitialProgram;
    procedure DebugSQLText(const SQLText: string);
    procedure InitConnection;
    procedure PreviewLabels;
    procedure AutoPopulateBarcodes;
    function GenerateLabelBitmaps: TList<TBitmap>;
    procedure PrintLabels;
    function CheckEmpty: Integer;
  public
  end;

var
  FormMain: TFormMain;

const
  COL_SELECTION = 0;
  COL_BARCODE = 1;
  COL_STATUS = 2;
  COL_JOB_NO = 3;
  COL_HASH = 4;
  COL_WEIGHT = 5;
  COL_COATING = 6;
  COL_VENDOR = 7;
  COL_PART_NO = 8;
  COL_PART_NAME = 9;
  COL_PART_QTY = 10;
  COL_MATERIAL = 11;
  COL_MATERIAL_SIZE = 12;
  COL_PLANNED_PROCESS_CD = 13;
  COL_PROCESS_NAME = 14;


implementation

uses
  SettingForm, DetailForm;

{$R *.dfm}
{$R DensoResource.RES}
{$R barcode.res}

// Click

procedure TFormMain.btnPreviewClick(Sender: TObject);
begin
  PreviewLabels;
end;

procedure TFormMain.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TFormMain.Copy1Click(Sender: TObject);
var
  Row, Col: Integer;
  TextToCopy: string;
begin
  TextToCopy := '';

  // Loop through each cell in the grid
  for Row := 0 to stgMain.RowCount - 1 do
  begin
    for Col := 1 to stgMain.ColCount - 1 do
    begin
      // Add cell content to TextToCopy with tab delimiters
      TextToCopy := TextToCopy + stgMain.Cells[Col, Row];
      if Col < stgMain.ColCount - 1 then
        TextToCopy := TextToCopy + #9; // Add a tab between columns
    end;
    TextToCopy := TextToCopy + sLineBreak; // Add a line break after each row
  end;
  Clipboard.AsText := TextToCopy;
end;

procedure TFormMain.btnPrintOutClick(Sender: TObject);
begin
  PrintLabels;
end;

procedure TFormMain.btnPrintSetupClick(Sender: TObject);
begin
  // Create the print setup form and show it as a modal dialog
  with TFormPrintSetup.Create(nil) do
    try
      ShowModal;
    finally
      Free;
    end;

end;

procedure TFormMain.Setting1Click(Sender: TObject);
begin
  if not Assigned(FormSetting) then
    FormSetting := TFormSetting.Create(Self);

  FormSetting.LoadSetting;

  if FormSetting.ShowModal = mrOk then
  begin
    ReadIniFile;
  end;
end;

procedure TFormMain.btnCloseClick(Sender: TObject);
begin
  WriteIniFile;
  Close;
end;

procedure TFormMain.btnDeleteAllClick(Sender: TObject);
begin
  ClearGridData;
end;

procedure TFormMain.btnDeleteClick(Sender: TObject);
var
  SelectedRow, i: Integer;
begin
  SelectedRow := stgMain.Row; // Get the currently selected row

  // Ensure there is a row to delete and it's not the header row
  if (SelectedRow > 0) and (SelectedRow < stgMain.RowCount) then
  begin
    // Shift rows up to delete the selected row
    for i := SelectedRow to stgMain.RowCount - 2 do
      stgMain.Rows[i].Assign(stgMain.Rows[i + 1]);

    // Clear the last row since it has been shifted up
    stgMain.Rows[stgMain.RowCount - 1].Clear;

    // Reduce row count to remove the empty row
    stgMain.RowCount := stgMain.RowCount - 1;
    dec(NextRow);
  end;
end;


procedure TFormMain.btnDeleteSelectionClick(Sender: TObject);
var
  Row, Col, LastRow: Integer;
begin
  LastRow := stgMain.RowCount - 1;

  // Loop from the first data row to the last row
  Row := 1;
  while Row <= LastRow do
  begin
    // Check if the row is selected by checking the "Selection" column (assuming it's column 0)
   if (stgMain.Cells[ColumnMapping[COL_SELECTION], Row] = '1') and
   (stgMain.Cells[ColumnMapping[COL_BARCODE], Row] <> '') then

    begin
      // Shift rows up to "delete" the selected row
      for var ShiftRow := Row to LastRow - 1 do
        for Col := 0 to stgMain.ColCount - 1 do
          stgMain.Cells[Col, ShiftRow] := stgMain.Cells[Col, ShiftRow + 1];

      // Clear the last row after shifting
      for Col := 0 to stgMain.ColCount - 1 do
        stgMain.Cells[Col, LastRow] := '';

      // Reduce RowCount to remove the empty row at the bottom
      Dec(LastRow);
      stgMain.RowCount := LastRow + 1;
      NextRow := NextRow - 1;
    end
    else
    begin
      Inc(Row); // Move to the next row if not selected
    end;
  end;
  ShowMessage('Selected rows have been deleted.');
end;

procedure TFormMain.btnDeSelectAllClick(Sender: TObject);
var
  Row: Integer;
  TempCellRect: TRect; // Local variable to avoid conflicts
begin
  for Row := 1 to stgMain.RowCount - 1 do
  begin
    stgMain.Cells[0, Row] := '0'; // Mark all as unselected
    TempCellRect := stgMain.CellRect(0, Row); // Use a unique name
    InvalidateRect(stgMain.Handle, @TempCellRect, False); // Redraw the cell
  end;
end;

procedure TFormMain.btnSelectAllClick(Sender: TObject);
var
  Row: Integer;
  TempCellRect: TRect; // Local variable to avoid conflicts
begin
  for Row := 1 to stgMain.RowCount - 1 do
  begin
    stgMain.Cells[0, Row] := '1'; // Mark all as unselected
    TempCellRect := stgMain.CellRect(0, Row); // Use a unique name
    InvalidateRect(stgMain.Handle, @TempCellRect, False); // Redraw the cell
  end;
end;


procedure TFormMain.stgMainMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  CellRect: TRect;
  CheckBoxRect: TRect;
  IsChecked: Boolean;
begin
  stgMain.MouseToCell(X, Y, ACol, ARow);

  // Check if the user clicked on the checkbox in the selection column (first column)
  if (ACol = 0) and (ARow > 0) then
  begin
    CellRect := stgMain.CellRect(ACol, ARow);

    // Define the CheckBoxRect centered within the CellRect
    CheckBoxRect := Rect(CellRect.Left + (CellRect.Width - 16) div 2,
      CellRect.Top + (CellRect.Height - 16) div 2,
      CellRect.Left + (CellRect.Width + 16) div 2,
      CellRect.Top + (CellRect.Height + 16) div 2);

    // Check if the point is within the checkbox rectangle
    if (X >= CheckBoxRect.Left) and (X <= CheckBoxRect.Right) and
      (Y >= CheckBoxRect.Top) and (Y <= CheckBoxRect.Bottom) then
    begin
      // Toggle checkbox state
      IsChecked := stgMain.Cells[ACol, ARow] = '1';
      if IsChecked then
        stgMain.Cells[ACol, ARow] := '0'
      else
        stgMain.Cells[ACol, ARow] := '1';

      // Redraw the cell
      InvalidateRect(stgMain.Handle, @CellRect, False);

      // Exit to avoid starting a column drag when clicking the checkbox
      Exit;
    end;
  end;
end;

procedure TFormMain.stgMainSelectCell(Sender: TObject; ACol, ARow: LongInt;var CanSelect: Boolean);
begin
  if (ARow > 0) and ((ACol = COL_WEIGHT) or (ACol = COL_COATING) or (ACol = COL_VENDOR)) then
    stgMain.Options := stgMain.Options + [goEditing] // Enable editing
  else
    stgMain.Options := stgMain.Options - [goEditing]; // Disable editing

  stgMain.Invalidate;

end;

procedure TFormMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    DataModuleCIMT.WriteSetting('Printer', 'Auto', 'False');
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  MainMemTable := TFDMemTable.Create(Self);
  InitialProgram;
  SetupStringGrid;
  NextRow := 1;
  stgMain.OnDrawCell := stgMainDrawCell;
  stgMain.OnMouseDown := stgMainMouseDown;

  //AutoPopulateBarcodes;
  ClearAll := SameText(DataModuleCIMT.readsetting('Setting', 'AutoClear', ''), 'True');
  AdjustColumnWidths;

end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  // Free the memory table when the form is destroyed
  MainMemTable.Free;
end;

procedure TFormMain.AdjustColumnWidths;
var
  Col, Row, TotalWidth, FixedWidth, RemainingWidth: Integer;
  MaxWidth, CellWidth: Integer;
  DynamicCols: Integer;
begin
  // Define fixed width for columns 5, 6, and 7
  FixedWidth := 100 * 3;

  // Calculate available width for dynamic columns
  TotalWidth := stgMain.ClientWidth;
  RemainingWidth := TotalWidth - FixedWidth;
  DynamicCols := stgMain.ColCount - 3; // Exclude columns 5, 6, and 7

  // Loop through each column
  for Col := 0 to stgMain.ColCount - 1 do
  begin
    // Columns 5, 6, and 7 have fixed widths
    if (Col = 5) or (Col = 6) or (Col = 7) then
    begin
      stgMain.ColWidths[Col] := 100;
      Continue;
    end;

    // Find the widest cell for dynamic columns
    MaxWidth := 0;
    for Row := 0 to stgMain.RowCount - 1 do
    begin
      CellWidth := stgMain.Canvas.TextWidth(stgMain.Cells[Col, Row]) + 20; // Add padding
      if CellWidth > MaxWidth then
        MaxWidth := CellWidth;
    end;

    // Set column width based on content width, constrained by available space
    stgMain.ColWidths[Col] := MaxWidth;
    RemainingWidth := RemainingWidth - stgMain.ColWidths[Col];
  end;

  // If any extra width remains, distribute it evenly among dynamic columns
  if RemainingWidth > 0 then
  begin
    for Col := 0 to stgMain.ColCount - 1 do
    begin
      if (Col <> 5) and (Col <> 6) and (Col <> 7) then
        stgMain.ColWidths[Col] := stgMain.ColWidths[Col] + (RemainingWidth div DynamicCols);
    end;
  end;
end;

procedure TFormMain.FormResize(Sender: TObject);
begin
  AdjustColumnWidths;
end;


procedure TFormMain.WriteIniFile;
var
  IniFile: TIniFile;
  IniFileName, IniFileDir: string;
begin
  IniFileDir := ExtractFilePath(Application.ExeName) + 'GRD\';
  IniFileName := IniFileDir + ChangeFileExt
    (ExtractFileName(Application.ExeName), '.ini');

  // Ensure the directory exists
  if not DirectoryExists(IniFileDir) then
    ForceDirectories(IniFileDir);

  IniFile := TIniFile.Create(IniFileName);
  try
    IniFile.WriteString('Parameter', 'SampleParam', edtbarcode.Text);
    IniFile.WriteInteger('Setting', 'IsUseRESTAPI',
      Integer(FIsUseRESTAPIState));
    IniFile.WriteInteger('Setting', 'DebugSQL', Integer(FDebugSQL));
    IniFile.WriteString('Setting', 'User', FUser);
    IniFile.WriteString('Setting', 'Pass', FPass);
  finally
    IniFile.Free;
  end;
end;

procedure TFormMain.ReadIniFile;
var
  IniFile: TIniFile;
  IniFileName: string;
begin
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\' +
    ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
  IniFile := TIniFile.Create(IniFileName);
  try
    edtbarcode.Text := IniFile.ReadString('Parameter', 'SampleParam', '');
    FIsUseRESTAPIState := TToggleSwitchState(IniFile.ReadInteger('Setting',
      'IsUseRESTAPI', Integer(tssOff)));
    FDebugSQL := TToggleSwitchState(IniFile.ReadInteger('Setting', 'DebugSQL',
      Integer(tssOff)));
    FUser := IniFile.ReadString('Setting', 'User', 'admin');
    FPass := IniFile.ReadString('Setting', 'Pass', 'admin');

  finally
    IniFile.Free;
  end;
end;

procedure TFormMain.InitialProgram;
var
  ExeFileName, FileVersion: string;

  function GetFileVersion(const FileName: TFileName): string;
  var
    Size, Handle: DWORD;
    Buffer: array of Byte;
    FileInfo: PVSFixedFileInfo;
    FileInfoSize: UINT;
  begin
    Size := GetFileVersionInfoSize(PChar(FileName), Handle);
    if Size = 0 then
      RaiseLastOSError;

    SetLength(Buffer, Size);
    if not GetFileVersionInfo(PChar(FileName), Handle, Size, Buffer) then
      RaiseLastOSError;

    if not VerQueryValue(Buffer, '\', Pointer(FileInfo), FileInfoSize) then
      RaiseLastOSError;

    Result := Format('%d.%d.%d.%d', [HiWord(FileInfo.dwFileVersionMS),
      LoWord(FileInfo.dwFileVersionMS), HiWord(FileInfo.dwFileVersionLS),
      LoWord(FileInfo.dwFileVersionLS)]);
  end;

begin
  ExeFileName := ExtractFileName(Application.ExeName);
  FileVersion := GetFileVersion(Application.ExeName);

  stbBase.Panels.Add;
  stbBase.Panels[0].Width := 210;
  stbBase.Panels.Add;
  stbBase.Panels[1].Width := 160;
  stbBase.Panels.Add;
  stbBase.Panels[2].Width := 600;

  stbBase.Panels[0].Text := ' ' + ExeFileName;
  stbBase.Panels[1].Text := FileVersion;

  DataModuleCIMT := TDataModuleCIMT.Create(Self);
  // Create the data module instance
  ReadIniFile;
  InitConnection;
end;

procedure TFormMain.InitConnection;
var
  ConnectionStatus: string;
begin
  // Initialize connection before any data operations
  try
    if not DataModuleCIMT.UniConnection1.Connected then
    begin
      ConnectionStatus := DataModuleCIMT.InitializeConnection
        (FIsUseRESTAPIState = tssOn);
      stbBase.Panels[2].Text := ConnectionStatus;
    end;
  except
    on E: Exception do
      ShowMessage('Error establishing Oracle connection: ' + E.Message);
  end;
end;

procedure TFormMain.DebugSQLText(const SQLText: string);
var
  DialogForm: TForm;
  Memo: TMemo;
begin
  DialogForm := TForm.Create(nil);
  try
    DialogForm.Width := 800;
    DialogForm.Height := 600;
    DialogForm.Position := poScreenCenter;
    DialogForm.Caption := 'SQL Debug';

    Memo := TMemo.Create(DialogForm);
    Memo.Parent := DialogForm;
    Memo.Align := alClient;
    Memo.Lines.Text := SQLText;
    Memo.ReadOnly := True;
    Memo.ScrollBars := ssVertical;
    Memo.Font.Name := 'Courier New';
    Memo.Font.Size := 10;

    Memo.SelectAll;

    DialogForm.ShowModal;
  finally
    DialogForm.Free;
  end;
end;

procedure TFormMain.ProcessBarcode;
var
  SQLText: string;
  barcode: string;
  Col,I: Integer;
  ResultMessage: string;
begin
  barcode := Trim(edtbarcode.Text);
    Delete(barcode, Length(barcode), 1);

  if barcode = '' then
  begin
  Exit;
  end;

  for I := 0 to stgMain.ColCount -1 do
    begin
         if edtbarcode.Text = stgMain.Cells[ColumnMapping[COL_BARCODE], I] then
      begin
        ShowMessage('This barcode already exists : '+ edtbarcode.Text);
        exit;
      end;
    end;
    SQLText :=
      'SELECT ' +
      '    kkk.KMSEQNO AS "Barcode", ' +
      '    kkk.waritukekbn AS "Type",' +
      '    k2.KANRYOFLG AS "Status Complete", ' +
      '    CASE ' +
      '        WHEN kkj.jkbn = 1 THEN ''Allotted'' ' +
      '        WHEN kkj.jkbn IS NULL THEN ''Not yet Allotted'' ' +
      '        WHEN kkj.jkbn = 2 AND kkj.jdankbn = 0 THEN ''Start'' ' +
      '        WHEN kkj.jkbn = 2 AND kkj.jdankbn = 9 THEN ''Discont'' ' +
      '        WHEN kkj.jkbn = 4 THEN ''Complete'' ' +
      '        ELSE ''Unknown'' ' +
      '    END AS "Status", ' +
      '    bhm.SEIZONO AS "Job No.", ' +
      '    bhm.BUBAN AS "Parts No", ' +
      '    bhm.BUNM AS "Parts Name", ' +
      '    bhm.BUBAN AS "#", ' +
      '    BHM.suryo AS "Part Qty", ' +
      '    BHM.ZAISITU AS "Material", ' +
      '    BHM.ZUBAN AS "Material Size", ' +
      '    k.KEIKOTEICD AS "Planned Process CD", ' +
      '    k.KOTEINM AS "Process Name", ' +
      '    ''1.00'' AS "Weight", ' +
      '    k.KEIKOTEICD AS "Coating" ' +
      'FROM ' +
      '    seizomst sei ' +
      'INNER JOIN ' +
      '    buhinkomst bhm ON sei.seizono = bhm.seizono ' +
      'INNER JOIN ' +
      '    keikakumst kkk ON bhm.seizono = kkk.seizono AND bhm.buno = kkk.buno ' +
      'INNER JOIN ' +
      '    KOUTEIKMST k ON k.KEIKOTEICD = kkk.KEIKOTEICD ' +
      'INNER JOIN ' +
      '    KEIKAKUOPT k2 ON k2.KMSEQNO = kkk.KMSEQNO ' +
      'LEFT JOIN ' +
      '    keikakujwmst kkj ON kkk.kmseqno = kkj.kmseqno ';


  SQLText := SQLText + Format(' WHERE kkk.KMSEQNO = %s ', [barcode]);

  if FDebugSQL = tssOn then
    DebugSQLText(SQLText);

  if FIsUseRESTAPIState = tssOff then
  begin
    DataModuleCIMT.FetchDataFromOracle(SQLText, MainMemTable);
  end
  else
  begin
    ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MainMemTable);
    if ResultMessage <> 'Success' then
    begin
      ShowMessage(ResultMessage);
      Exit;
    end;
  end;

  if not MainMemTable.IsEmpty and (MainMemTable.FieldByName('Status Complete').AsInteger <> 1) and  (MainMemTable.FieldByName('Type').AsInteger = 2)  then
  begin
    // Ensure enough rows in StringGrid
    if stgMain.RowCount <= NextRow then
      stgMain.RowCount := NextRow + 1;

    MainMemTable.First;

    stgMain.ColCount := MainMemTable.FieldCount + 1;
    // Ensure columns match data

    stgMain.Cells[0, NextRow] := '1';
    // Populate data based on SQL column mapping
    stgMain.Cells[ColumnMapping[COL_BARCODE], NextRow] :=
      edtbarcode.Text;
    stgMain.Cells[ColumnMapping[COL_STATUS], NextRow] :=
      MainMemTable.FieldByName('Status').AsString;
    stgMain.Cells[ColumnMapping[COL_JOB_NO], NextRow] :=
      MainMemTable.FieldByName('Job No.').AsString;
    stgMain.Cells[ColumnMapping[COL_PART_NO], NextRow] :=
      MainMemTable.FieldByName('Parts No').AsString;
    stgMain.Cells[ColumnMapping[COL_HASH], NextRow] :=
      MainMemTable.FieldByName('#').AsString;
    stgMain.Cells[ColumnMapping[COL_PART_NAME], NextRow] :=
      MainMemTable.FieldByName('Parts Name').AsString;
    stgMain.Cells[ColumnMapping[COL_PART_QTY], NextRow] :=
      MainMemTable.FieldByName('Part Qty').AsString;
    stgMain.Cells[ColumnMapping[COL_MATERIAL], NextRow] :=
      MainMemTable.FieldByName('Material').AsString;
    stgMain.Cells[ColumnMapping[COL_MATERIAL_SIZE], NextRow] :=
      MainMemTable.FieldByName('Material Size').AsString;
    stgMain.Cells[ColumnMapping[COL_PLANNED_PROCESS_CD], NextRow] :=
      MainMemTable.FieldByName('Planned Process CD').AsString;
    stgMain.Cells[ColumnMapping[COL_PROCESS_NAME], NextRow] :=
      MainMemTable.FieldByName('Process Name').AsString;
    stgMain.Cells[ColumnMapping[COL_WEIGHT], NextRow] := ''; // Fixed value
    stgMain.Cells[ColumnMapping[COL_COATING], NextRow] :=
     MainMemTable.FieldByName('Coating').AsString;
    stgMain.Cells[ColumnMapping[COL_VENDOR], NextRow] := ''; // Fixed value

    Inc(NextRow); // Move to the next row for the following barcode
    MainMemTable.Next;

  end
  else
  begin

    if MainMemTable.IsEmpty then
   begin
      ShowMessage('No data found for barcode: ' + edtbarcode.Text);
   end
   else if MainMemTable.FieldByName('Type').AsInteger <> 2 then
   begin
         ShowMessage('Cannot input in-house process barcode: ' + edtbarcode.Text);
   end
   else if (MainMemTable.FieldByName('Status Complete').AsInteger = 1) then
   begin
     ShowMessage('This process has already completed: ' + edtbarcode.Text);
   end;

  end;
  edtbarcode.Clear; // Clear the input box for the next barcode entry
end;

procedure TFormMain.edtbarcodeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then // Check if Enter key is pressed
  begin
    Key := #0; // Prevent the beep sound on Enter
    if edtbarcode.Text = '#PRINTALLSELECT' then
    begin
      printlabels;
      edtbarcode.Clear;
    end
    else if edtbarcode.Text = '#DELETEALL' then
    begin
      ClearGridData;
       edtbarcode.Clear;
    end
    else
    begin
       ProcessBarcode;
    end;
  end;
end;

procedure TFormMain.ClearGridData;
var
  Row, Col: Integer;
begin
  // Loop through each row starting from the first data row (assuming row 0 is header)
  for Row := 1 to stgMain.RowCount - 1 do
  begin
    // Clear each cell in the row
    for Col := 0 to stgMain.ColCount - 1 do
    begin
      stgMain.Cells[Col, Row] := '';
    end;
  end;

  // Set RowCount to 2 to reset the grid and leave one row below the header
  stgMain.RowCount := 2;

  // Update NextRow to the first data row
  NextRow := 1;
end;


procedure TFormMain.SetupStringGrid;
var
  i: Integer;
begin
  // Define the columns
  stgMain.ColCount := 15; // Total number of columns based on your headers
  SetLength(ColumnMapping, stgMain.ColCount);
  for i := 0 to stgMain.ColCount - 1 do
  begin
    ColumnMapping[i] := i;
    stgMain.ColWidths[i] := 100;
  end;

  stgMain.FixedRows := 1;
  stgMain.FixedCols := 0;

  stgMain.Cells[COL_SELECTION, 0] := 'Selection';
  stgMain.Cells[COL_BARCODE, 0] := 'Barcode';
  stgMain.Cells[COL_STATUS, 0] := 'Status';
  stgMain.Cells[COL_JOB_NO, 0] := 'Job No.';
  stgMain.Cells[COL_WEIGHT, 0] := 'Weight';
  stgMain.Cells[COL_COATING, 0] := 'Coating';
  stgMain.Cells[COL_VENDOR, 0] := 'VENDOR';
  stgMain.Cells[COL_HASH, 0] := '#';
  stgMain.Cells[COL_PART_NO, 0] := 'Part No';
  stgMain.Cells[COL_PART_NAME, 0] := 'Part Name';
  stgMain.Cells[COL_PART_QTY, 0] := 'Part Qty';
  stgMain.Cells[COL_MATERIAL, 0] := 'Material';
  stgMain.Cells[COL_MATERIAL_SIZE, 0] := 'Material Size';
  stgMain.Cells[COL_PLANNED_PROCESS_CD, 0] := 'Planned Process CD';
  stgMain.Cells[COL_PROCESS_NAME, 0] := 'Process Name';


  stgMain.DefaultDrawing := False; // We'll handle drawing ourselves

end;

procedure TFormMain.stgMainDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  CheckBoxRect: TRect;
  IsChecked: Boolean;
  DrawState: Integer;
  StatusText: string;
begin
  with stgMain.Canvas do
  begin
    // Default background colors for header and alternating rows
    if ARow = 0 then
    begin
      Brush.Color := clBtnFace; // Header background
      Font.Style := [fsBold];
      Font.Color := clBlack; // Header font color
    end
    else if ARow mod 2 = 0 then
      Brush.Color := RGB($D0, $E7, $FD) // Light blue for even rows
    else
      Brush.Color := clWhite; // White for odd rows

    // Apply conditional formatting based on the Status column value
    if (ARow > 0) and (ACol = ColumnMapping[COL_STATUS]) then
    begin
      StatusText := stgMain.Cells[ACol, ARow];

      // Set background color based on Status
      if StatusText = 'Allotted' then
        Brush.Color := RGB($00, $99, $00) // Dark green
      else if StatusText = 'Not yet Allotted' then
        Brush.Color := clWhite
      else if StatusText = 'Start' then
        Brush.Color := RGB($ff, $66, $ff) // Light purple
      else if StatusText = 'Discont' then
        Brush.Color := RGB($ff, $3d, $3d) // Red
      else if StatusText = 'Complete' then
        Brush.Color := clYellow
      else if StatusText = 'Unknown' then
        Brush.Color := clGray;
    end;
    if (ARow > 0) and (ARow = stgMain.Row) then
    begin
      Brush.Color := RGB($98, $AC, $FE);
      Font.Color := clWhite;
    end
    else
    begin
      Font.Color := clBlack; // Reset font color for non-selected rows
    end;

    FillRect(Rect); // Fill cell background with the chosen color

    // Draw header text
    if ARow = 0 then
    begin
      TextOut(Rect.Left + 2, Rect.Top + 2, stgMain.Cells[ACol, ARow]);
      Font.Style := []; // Reset font style after header
    end
    // Draw checkbox in the first column for each data row
    else if (ACol = 0) and (ARow > 0) then
    begin
      // Determine if the checkbox is checked based on cell value ('1' for checked, '0' for unchecked)
      IsChecked := stgMain.Cells[ACol, ARow] = '1';
      CheckBoxRect.Left := Rect.Left + (Rect.Width - 16) div 2;
      CheckBoxRect.Top := Rect.Top + (Rect.Height - 16) div 2;
      CheckBoxRect.Right := CheckBoxRect.Left + 16;
      CheckBoxRect.Bottom := CheckBoxRect.Top + 16;

      // Set the checkbox style
      DrawState := DFCS_BUTTONCHECK;
      if IsChecked then
        DrawState := DrawState or DFCS_CHECKED;

      // Draw the checkbox
      DrawFrameControl(Handle, CheckBoxRect, DFC_BUTTON, DrawState);
    end
    else
    begin
      // Draw cell text for other columns
      TextOut(Rect.Left + 2, Rect.Top + 2, stgMain.Cells[ACol, ARow]);
    end;

    // Draw black border if the cell is in columns 5, 6, or 7
    if (gdSelected in State) then
    begin
      Pen.Color := clBlack;
      Pen.Width := 1;
      Brush.Style := bsClear;
      Rectangle(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);
    end;

  end;

end;

function TFormMain.CheckEmpty: Integer;
var
  i: Integer;
begin
  EM := 0;
  for i := 1 to stgMain.RowCount - 1 do
  begin
    if stgMain.Cells[ColumnMapping[COL_SELECTION], i] = '1' then
    begin
      // Retrieve values from the current row
      if stgMain.Cells[ColumnMapping[COL_BARCODE], i] <> '' then
      begin
        Inc(EM);
      end;
    end;
  end;

  // Return the value of EM
  Result := EM;
end;

function TFormMain.GenerateLabelBitmaps: TList<TBitmap>;
const
  LabelWidthMM = 90;
  LabelHeightMM = 45;
  BaseDPI = 203; // Base DPI for original coordinates
var
  i: Integer;
  LabelBitmap: TBitmap;
  LabelWidthPx, LabelHeightPx, PrinterDPI: Integer;
  ScaleFactor, FontScaleFactor: Double;
  barcode, Status, JobNo, PartsNo, PartsName, Qty, Weight, Coating,
    Vendor: string;
  ResourceStream: TResourceStream;
  DensoLogo: TPngImage;
  BitmapList: TList<TBitmap>;
begin
    // Get the DPI of the default printer
    PrinterDPI := StrToIntDef(DataModuleCIMT.ReadSetting('Printer', 'DPI', '203'), 203);
    // Calculate scale factors for drawing and fonts
    ScaleFactor := PrinterDPI / BaseDPI;
    FontScaleFactor := 1 + (ScaleFactor - 1) / 32; // Smooth scaling for fonts

    // Calculate width and height based on the printer's DPI
    LabelWidthPx := Round(LabelWidthMM * PrinterDPI / 25.4);
    LabelHeightPx := Round(LabelHeightMM * PrinterDPI / 25.4);

    // Create a list to store all bitmaps
    BitmapList := TList<TBitmap>.Create;
    try
      for i := 1 to stgMain.RowCount - 1 do
      begin
        if stgMain.Cells[ColumnMapping[COL_SELECTION], i] = '1' then
        begin
          // Retrieve values from the current row
          if stgMain.Cells[ColumnMapping[COL_BARCODE], i] = '' then
          begin
              continue;
          end;
          inc(LastPage);

          barcode := stgMain.Cells[ColumnMapping[COL_BARCODE], i];
          Status := stgMain.Cells[ColumnMapping[COL_STATUS], i];
          JobNo := stgMain.Cells[ColumnMapping[COL_JOB_NO], i];
          PartsNo := stgMain.Cells[ColumnMapping[COL_PART_NO], i];
          Qty := stgMain.Cells[ColumnMapping[COL_PART_QTY], i];
          Weight := stgMain.Cells[ColumnMapping[COL_WEIGHT], i];
          Coating := stgMain.Cells[ColumnMapping[COL_COATING], i];
          Vendor := stgMain.Cells[ColumnMapping[COL_VENDOR], i];

          // Create a new bitmap for the label
          LabelBitmap := TBitmap.Create;
          try
            LabelBitmap.SetSize(LabelWidthPx, LabelHeightPx);
            LabelBitmap.Canvas.Brush.Color := clWhite;
            LabelBitmap.Canvas.FillRect(Rect(0, 0, LabelWidthPx, LabelHeightPx));

            with LabelBitmap.Canvas do
            begin
              Font.Name := 'Arial';
              Font.PixelsPerInch := PrinterDPI; // Set font DPI to match printer DPI

              // Draw borders with fixed sizes, scaled by ScaleFactor
              Pen.Color := clBlack;
              Pen.Width := Round(2 * ScaleFactor);
              Rectangle(
                Round(10 * ScaleFactor), Round(10 * ScaleFactor),
                LabelWidthPx - Round(10 * ScaleFactor), LabelHeightPx - Round(10 * ScaleFactor)
              );
              Pen.Width := Round(1 * ScaleFactor);
              Rectangle(
                Round(30 * ScaleFactor), Round(30 * ScaleFactor),
                LabelWidthPx - Round(30 * ScaleFactor), LabelHeightPx - Round(30 * ScaleFactor)
              );

              // Load and draw DENSO logo
              try
                ResourceStream := TResourceStream.Create(HInstance, 'DENSOLOGO', RT_RCDATA);
                try
                  DensoLogo := TPngImage.Create;
                  try
                    DensoLogo.LoadFromStream(ResourceStream);
                    StretchDraw(Rect(
                      Round(40 * ScaleFactor), Round(35 * ScaleFactor),
                      Round(150 * ScaleFactor), Round(60 * ScaleFactor)
                    ), DensoLogo);
                  finally
                    DensoLogo.Free;
                  end;
                finally
                  ResourceStream.Free;
                end;
              except
                on E: Exception do
                  ShowMessage('Error loading DENSOLOGO resource: ' + E.Message);
              end;

              // Draw text with adjusted font sizes
              Font.Size := Round(8*FontScaleFactor); // Smaller font scaling  +
              Font.Style := [];
              TextOut(Round(168 * ScaleFactor), Round(40 * ScaleFactor),
              'Innovative Manufacturing Solution Asia Co.,Ltd');

              Font.Size := Round(10*FontScaleFactor);
              Font.Style := [fsBold];
              TextOut((LabelWidthPx div 2) - TextWidth('Special Coating') div 2,
              Round(90 * ScaleFactor), 'Special Coating');
              Font.Style := [];
              Font.Size := Round(8*FontScaleFactor);

              // Draw labels and lines, scaled by ScaleFactor
              TextOut(Round(40 * ScaleFactor), Round(150 * ScaleFactor), 'JOB. NO.');
              MoveTo(Round(150 * ScaleFactor), Round(175 * ScaleFactor));
              LineTo(Round(400 * ScaleFactor), Round(175 * ScaleFactor));
              TextOut(Round(410 * ScaleFactor), Round(150 * ScaleFactor), 'Q''ty');
              MoveTo(Round(450 * ScaleFactor), Round(175 * ScaleFactor));
              LineTo(Round(590 * ScaleFactor), Round(175 * ScaleFactor));
              TextOut(Round(600 * ScaleFactor), Round(150 * ScaleFactor), 'Pcs.');
              TextOut(Round(40 * ScaleFactor), Round(210 * ScaleFactor), 'Parts No.');
              MoveTo(Round(150 * ScaleFactor), Round(235 * ScaleFactor));
              LineTo(Round(400 * ScaleFactor), Round(235 * ScaleFactor));
              TextOut(Round(410 * ScaleFactor), Round(210 * ScaleFactor), 'Weight');
              MoveTo(Round(490 * ScaleFactor), Round(235 * ScaleFactor));
              LineTo(Round(590 * ScaleFactor), Round(235 * ScaleFactor));
              TextOut(Round(600 * ScaleFactor), Round(210 * ScaleFactor), 'Kgs.');
              TextOut(Round(40 * ScaleFactor), Round(270 * ScaleFactor), 'Coating');
              MoveTo(Round(150 * ScaleFactor), Round(295 * ScaleFactor));
              LineTo(Round(400 * ScaleFactor), Round(295 * ScaleFactor));
              TextOut(Round(410 * ScaleFactor), Round(270 * ScaleFactor), 'Vendor');
              MoveTo(Round(490 * ScaleFactor), Round(295 * ScaleFactor));
              LineTo(Round(650 * ScaleFactor), Round(295 * ScaleFactor));
              Font.Size := Round(7*FontScaleFactor);
              Font.Style := [fsBold];
              // Insert dynamic content with scaled positions
              TextOut(Round(170 * ScaleFactor), Round(145 * ScaleFactor), JobNo);
              TextOut(Round(490 * ScaleFactor), Round(145 * ScaleFactor), Qty);
              TextOut(Round(170 * ScaleFactor), Round(205 * ScaleFactor), PartsNo);
              TextOut(Round(540 * ScaleFactor), Round(205 * ScaleFactor), Weight);
              TextOut(Round(170 * ScaleFactor), Round(265 * ScaleFactor), Coating);
              TextOut(Round(510 * ScaleFactor), Round(265 * ScaleFactor), Vendor);
            end;

            // Add the LabelBitmap to the list
            BitmapList.Add(LabelBitmap);
          except
            LabelBitmap.Free;
            raise;
          end;
        end;
      end;

      Result := BitmapList;
    except
      // Clean up if there was an error generating bitmaps
      for LabelBitmap in BitmapList do
        LabelBitmap.Free;
      BitmapList.Free;
      raise;
    end;
end;


procedure TFormMain.PreviewLabels;
var
  BitmapList: TList<TBitmap>;
begin
    if CheckEmpty = 0 then
    begin
        ShowMessage('NO DATA');
        exit;
    end;

  BitmapList := GenerateLabelBitmaps;
  try
    FormPreview := TFormPreview.Create(Self);
    try
      FormPreview.LoadImages(BitmapList); // Load all bitmaps at once
      FormPreview.ShowModal;
    finally
      FormPreview.Free;
    end;
  finally
    // Free all bitmaps in the list
    for var LabelBitmap in BitmapList do
      LabelBitmap.Free;
    BitmapList.Free;
  end;
end;

procedure TFormMain.PrintLabels;
  function SelectPrinterByName(const PrinterName: string): Boolean;
  var
    i: Integer;
  begin
      if CheckEmpty = 0 then
    begin
        ShowMessage('NO DATA');
        exit;
    end;

    Result := False;
    for i := 0 to Printer.Printers.Count - 1 do
    begin
      if SameText(Printer.Printers[i], PrinterName) then
      begin
        Printer.PrinterIndex := i; // Select the printer by setting PrinterIndex
        Result := True;
        Exit;
      end;
    end;
    ShowMessage('Printer not found: ' + PrinterName);
  end;

var
  BitmapList: TList<TBitmap>;
  i: Integer;
begin
  ClearAll := (StrToIntDef(DataModuleCIMT.ReadSetting('Setting', 'AutoClear', IntToStr(Integer(tssOff))), Integer(tssOff)) <> 0);
  // Select the printer
  if not SelectPrinterByName(DataModuleCIMT.ReadSetting('Printer', 'Name', ''))
  then
  begin
    Exit; // Stop if the printer is not found
  end;


  BitmapList := GenerateLabelBitmaps;

  try
    Printer.BeginDoc;
    try
      for i := 0 to BitmapList.Count - 1 do
      begin
        if i > 0 then
          Printer.NewPage;
          Printer.Canvas.Draw(0, 0, BitmapList[i]);
        // Adjust (0, 0) to position the image on the page
      end;
    finally
      Printer.EndDoc;
    end;
  finally
    // Free all bitmaps in the list
    if ClearAll then
    begin
      ClearGridData;
    end;

    for var LabelBitmap in BitmapList do
      LabelBitmap.Free;
    BitmapList.Free;
  end;
end;

procedure TFormMain.Printout1Click(Sender: TObject);
begin
  PrintLabels;
end;

// For debug
procedure TFormMain.AutoPopulateBarcodes;
const
  Barcodes: array [0 .. 3] of string = ('31776860', '32180314', '27072921',
    '2211255');
var
  i: Integer;
begin
  for i := Low(Barcodes) to High(Barcodes)-3 do
  begin
    edtbarcode.Text := Barcodes[i];
    ProcessBarcode;
  end;
end;

end.
