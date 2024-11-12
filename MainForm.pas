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
  ExcelXP, Vcl.Imaging.pngimage,
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
    Edit1: TMenuItem;
    Copy1: TMenuItem;
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
    procedure readPrinterSetup(Sheet: Variant);
    procedure SetDefaultPrinter(const PrinterName: string);
    procedure Copy1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);

  private
    FDraggingColumn: Integer; // Column currently being dragged
    PageCount: Integer;
    FDragStartX: Integer; // Initial X position when the drag started
    FDragTargetColumn: Integer; // Target column to move to
    ColumnMapping: array of Integer;
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
    procedure PrintImagePreview;
    procedure PrintCoatingLabel;
    procedure PreviewLabels;
    procedure AutoPopulateBarcodes;
  public
  end;

var
  FormMain: TFormMain;

const
  COL_SELECTION = 0;
  COL_BARCODE = 1;
  COL_STATUS = 2;
  COL_JOB_NO = 3;
  COL_PARTS_NO = 4;
  COL_PARTS_NAME = 5;
  COL_HASH = 6;
  COL_PART_NAME = 7;
  COL_PART_QTY = 8;
  COL_MATERIAL = 9;
  COL_MATERIAL_SIZE = 10;
  COL_PLANNED_PROCESS_CD = 11;
  COL_PROCESS_NAME = 12;
  COL_WEIGHT = 13;
  COL_COATING = 14;

implementation

uses
  SettingForm, DetailForm;

{$R *.dfm}
{$R DensoResource.RES}
// Click

procedure TFormMain.btnPreviewClick(Sender: TObject);
begin
  PreviewLabels;
end;

procedure TFormMain.Exit1Click(Sender: TObject);
begin
  halt;
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
  PrintCoatingLabel;
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

  ShowMessage('All rows have been cleared.');
end;

procedure TFormMain.btnDeleteClick(Sender: TObject);
begin
  // print out delete
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
    if stgMain.Cells[ColumnMapping[COL_SELECTION], Row] = '1' then
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

procedure TFormMain.FormCreate(Sender: TObject);
begin
  // Initialize the memory table and other components
  MainMemTable := TFDMemTable.Create(Self);
  InitialProgram;
  SetupStringGrid;
  NextRow := 1;
  stgMain.OnDrawCell := stgMainDrawCell;
  stgMain.OnMouseDown := stgMainMouseDown;

  // Call AutoPopulateBarcodes to automatically fill the grid
  AutoPopulateBarcodes;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  // Free the memory table when the form is destroyed
  MainMemTable.Free;
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
  Col: Integer;
  ResultMessage: string;
begin
  barcode := Trim(edtbarcode.Text);
  if barcode = '' then
    Exit;

  SQLText :=
    'SELECT kkk.KMSEQNO AS "Barcode",k2.KANRYOFLG AS "status",bhm.SEIZONO AS "job no." '
    + ',bhm.BUBAN AS "parts no",bhm.BUNM  AS "parts name",bhm.BUBAN AS "#" ,bhm.BUNM  AS "part name" '
    + ',BHM.ZUBAN2  AS "part qty",BHM.ZAISITU AS "material" ,BHM.ZUBAN  AS "material size" ,k.KEIKOTEICD  AS "Planned Process CD" '
    + ',k.KOTEINM  AS "Process Name", ''1'' AS "Weight" , ''2'' AS "Coating" FROM'
    + ' seizomst sei ' +
    'INNER JOIN buhinkomst bhm ON sei.seizono = bhm.seizono ' +
    'INNER JOIN keikakumst kkk ON bhm.seizono = kkk.seizono AND bhm.buno = kkk.buno '
    + 'INNER JOIN KOUTEIKMST k ON k.KEIKOTEICD = kkk.KEIKOTEICD ' +
    'INNER JOIN KEIKAKUOPT k2 ON k2.KMSEQNO = kkk.KMSEQNO';

  SQLText := SQLText + Format(' WHERE kkk.KMSEQNO = %s', [barcode]);

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

  if not MainMemTable.IsEmpty then
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
      MainMemTable.FieldByName('Barcode').AsString;
    stgMain.Cells[ColumnMapping[COL_STATUS], NextRow] :=
      MainMemTable.FieldByName('Status').AsString;
    stgMain.Cells[ColumnMapping[COL_JOB_NO], NextRow] :=
      MainMemTable.FieldByName('Job No.').AsString;
    stgMain.Cells[ColumnMapping[COL_PARTS_NO], NextRow] :=
      MainMemTable.FieldByName('Parts No').AsString;
    stgMain.Cells[ColumnMapping[COL_PARTS_NAME], NextRow] :=
      MainMemTable.FieldByName('Parts Name').AsString;
    stgMain.Cells[ColumnMapping[COL_HASH], NextRow] :=
      MainMemTable.FieldByName('#').AsString;
    stgMain.Cells[ColumnMapping[COL_PART_NAME], NextRow] :=
      MainMemTable.FieldByName('Part Name').AsString;
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
    stgMain.Cells[ColumnMapping[COL_WEIGHT], NextRow] := '1'; // Fixed value
    stgMain.Cells[ColumnMapping[COL_COATING], NextRow] := '2'; // Fixed value

    Inc(NextRow); // Move to the next row for the following barcode
    MainMemTable.Next;

  end
  else
  begin
    ShowMessage('No data found for barcode: ' + barcode);
  end;
  edtbarcode.Clear; // Clear the input box for the next barcode entry
end;

procedure TFormMain.edtbarcodeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then // Check if Enter key is pressed
  begin
    Key := #0; // Prevent the beep sound on Enter
    ProcessBarcode; // Call the existing procedure to execute the query
  end;
end;

procedure TFormMain.SetupStringGrid;
var
  i: Integer;
begin
  // Define the columns
  stgMain.ColCount := 15; // Total number of columns based on your headers
  SetLength(ColumnMapping, stgMain.ColCount);
  for i := 0 to stgMain.ColCount - 1 do
    ColumnMapping[i] := i;

  stgMain.ColWidths[0] := 80;
  stgMain.ColWidths[1] := 100;
  stgMain.ColWidths[2] := 100;
  stgMain.ColWidths[3] := 100;
  stgMain.ColWidths[4] := 100;
  stgMain.ColWidths[5] := 100;
  stgMain.ColWidths[6] := 100;
  stgMain.ColWidths[7] := 100;
  stgMain.ColWidths[8] := 100;
  stgMain.ColWidths[9] := 100;
  stgMain.ColWidths[10] := 100;
  stgMain.ColWidths[12] := 130;
  stgMain.ColWidths[13] := 100;
  stgMain.ColWidths[14] := 100;

  stgMain.FixedRows := 1;
  stgMain.FixedCols := 0;

  stgMain.Cells[COL_SELECTION, 0] := 'Selection';
  stgMain.Cells[COL_BARCODE, 0] := 'Barcode';
  stgMain.Cells[COL_STATUS, 0] := 'Status';
  stgMain.Cells[COL_JOB_NO, 0] := 'Job No.';
  stgMain.Cells[COL_PARTS_NO, 0] := 'Parts No.';
  stgMain.Cells[COL_PARTS_NAME, 0] := 'Parts Name';
  stgMain.Cells[COL_HASH, 0] := '#';
  stgMain.Cells[COL_PART_NAME, 0] := 'Part Name';
  stgMain.Cells[COL_PART_QTY, 0] := 'Part Qty';
  stgMain.Cells[COL_MATERIAL, 0] := 'Material';
  stgMain.Cells[COL_MATERIAL_SIZE, 0] := 'Material Size';
  stgMain.Cells[COL_PLANNED_PROCESS_CD, 0] := 'Planned Process CD';
  stgMain.Cells[COL_PROCESS_NAME, 0] := 'Process Name';
  stgMain.Cells[COL_WEIGHT, 0] := 'Weight';
  stgMain.Cells[COL_COATING, 0] := 'Coating';

  stgMain.DefaultDrawing := False; // We'll handle drawing ourselves

end;

procedure TFormMain.stgMainDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  CheckBoxRect: TRect;
  IsChecked: Boolean;
  DrawState: Integer;
begin
  with stgMain.Canvas do
  begin
    // Set background colors
    if ARow = 0 then
    begin
      // Header row background color and bold text
      Brush.Color := clBtnFace;
      Font.Style := [fsBold];
    end
    else if ARow mod 2 = 0 then
      Brush.Color := RGB($D0, $E7, $FD) // Light blue for even rows
    else
      Brush.Color := clWhite; // White for odd rows

    FillRect(Rect); // Fill cell background

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

      // Define checkbox rectangle
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
  end;
end;

procedure TFormMain.PreviewLabels;
const
  DPI = 203; // Target DPI for the label image
  LabelWidthMM = 90;
  LabelHeightMM = 45;
var
  i: Integer;
  LabelBitmap: TBitmap;
  LabelWidthPx, LabelHeightPx: Integer;
  Barcode, Status, JobNo, PartsNo, PartsName, Qty, Weight, Coating, Vendor: string;
  ResourceStream: TResourceStream;
  DensoLogo: TPngImage;
begin
  LabelWidthPx := Round(LabelWidthMM * DPI / 25.4);
  LabelHeightPx := Round(LabelHeightMM * DPI / 25.4);

  FormPreview := TFormPreview.Create(Self);
  try
    FormPreview.TotalPages := stgMain.RowCount - 1; // Set total pages for pagination
    FormPreview.CurrentPage := 1;

    // Loop through each row in the StringGrid, skipping the header row (Row 0)
    for i := 1 to stgMain.RowCount - 1 do
    begin
      // Retrieve values from the current row
      Barcode := stgMain.Cells[COL_BARCODE, i];
      Status := stgMain.Cells[COL_STATUS, i];
      JobNo := stgMain.Cells[COL_JOB_NO, i];
      PartsNo := stgMain.Cells[COL_PARTS_NO, i];
      PartsName := stgMain.Cells[COL_PARTS_NAME, i];
      Qty := stgMain.Cells[COL_PART_QTY, i];
      Weight := stgMain.Cells[COL_WEIGHT, i];
      Coating := stgMain.Cells[COL_COATING, i];
      Vendor := 'Some Vendor'; // Set vendor as needed

      // Create a new bitmap for the label
      LabelBitmap := TBitmap.Create;
      try
        LabelBitmap.SetSize(LabelWidthPx, LabelHeightPx);
        LabelBitmap.Canvas.Brush.Color := clWhite;
        LabelBitmap.Canvas.FillRect(Rect(0, 0, LabelWidthPx, LabelHeightPx));

        with LabelBitmap.Canvas do
        begin
          Font.Name := 'Arial';

          // Draw outer and inner borders
          Pen.Color := clBlack;
          Pen.Width := 2;
          Rectangle(10, 10, LabelWidthPx - 10, LabelHeightPx - 10);
          Pen.Width := 1;
          Rectangle(20, 20, LabelWidthPx - 20, LabelHeightPx - 20);

          // Load and draw DENSO logo
          try
            ResourceStream := TResourceStream.Create(HInstance, 'DENSOLOGO', RT_RCDATA);
            try
              DensoLogo := TPngImage.Create;
              try
                DensoLogo.LoadFromStream(ResourceStream);
                StretchDraw(Rect(30, 25, 80, 45), DensoLogo); // Adjust the logo size and position
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

          // Draw company name
          Font.Size := 8;
          TextOut(90, 30, 'Innovative Manufacturing Solution Asia Co.,Ltd');

          // Draw title "Special Coating"
          Font.Size := 10;
          Font.Style := [fsBold];
          TextOut(90, 50, 'Special Coating');
          Font.Style := []; // Reset style

          // Draw JOB. NO., Q'ty, and Parts No.
          Font.Size := 8;
          TextOut(30, 80, 'JOB. NO.');
          MoveTo(110, 90); LineTo(290, 90); // Adjusted the line width for JOB. NO.
          TextOut(320, 80, 'Q''ty');
          MoveTo(350, 90); LineTo(420, 90); // Adjusted the line width for Q'ty
          TextOut(440, 80, 'Pcs.');

          TextOut(30, 100, 'Parts No.');
          MoveTo(110, 110); LineTo(290, 110); // Adjusted the line width for Parts No.
          TextOut(320, 100, 'Weight');
          MoveTo(370, 110); LineTo(420, 110); // Adjusted the line width for Weight
          TextOut(440, 100, 'Kgs.');

          // Draw Coating and Vendor fields
          TextOut(30, 120, 'Coating');
          MoveTo(110, 130); LineTo(290, 130); // Adjusted the line width for Coating
          TextOut(320, 120, 'Vendor');
          MoveTo(370, 130); LineTo(500, 130); // Adjusted the line width for Vendor

          // Insert dynamic content
          Font.Style := [];
          TextOut(115, 80, JobNo);
          TextOut(355, 80, Qty);
          TextOut(115, 100, PartsNo);
          TextOut(375, 100, Weight);
          TextOut(115, 120, Coating);
          TextOut(375, 120, Vendor);
        end;

        // Send bitmap to PreviewForm
        FormPreview.LoadImageFromBitmap(LabelBitmap, i);
        Application.ProcessMessages;

      finally
        LabelBitmap.Free; // Free the bitmap after each loop iteration
      end;
    end;

    FormPreview.ShowModal; // Display the preview form

  finally
    FormPreview.Free;
  end;
end;


procedure TFormMain.PrintCoatingLabel();

  procedure SelectPrinterByName(const PrinterName: string);
  var
    i: Integer;
  begin
    for i := 0 to Printer.Printers.Count - 1 do
    begin
      if SameText(Printer.Printers[i], PrinterName) then
      begin
        Printer.PrinterIndex := i; // Select the printer by setting PrinterIndex
        Exit;
      end;
    end;

    // If not found, show a message or handle the error
    ShowMessage('Printer not found: ' + PrinterName);
  end;

begin
  // SelectPrinterByName('Microsoft Print to PDF');
  SelectPrinterByName('SATO WS408');
  PrintImagePreview;
end;

procedure TFormMain.readPrinterSetup(Sheet: Variant);
var
  Ini: TIniFile;
  IniFileName, PaperStyle, PrinterName, FullPrinterName: string;
  UseDefaultPrinter: Boolean;
  PrinterFound: Boolean;
  PaperSizes: array of Word;
  PaperNames: array of array [0 .. 63] of Char;
  NumPaperSizes, i, DesiredPaperSize: Integer;
  hPrinter: THandle;
  SelectedPaperSize: Integer;
  ExcelApp: Variant;
begin
  IniFileName := ExtractFilePath(Application.ExeName) +
    'GRD\DS13LBC100_printer.ini';
  Ini := TIniFile.Create(IniFileName);
  try
    // Read settings from the INI file
    UseDefaultPrinter := Ini.ReadBool('Settings', 'UseDefaultPrinter', False);
    PaperStyle := Ini.ReadString('Settings', 'PaperStyle', 'Horizontal');
    DesiredPaperSize := Ini.ReadInteger('Settings', 'PaperSizeNum', 1);
    PrinterName := Ini.ReadString('Settings', 'Printer', '');

    // Set orientation based on PaperStyle
    if PaperStyle = 'Vertical' then
      Sheet.PageSetup.Orientation := 1 // xlPortrait
    else
      Sheet.PageSetup.Orientation := 2; // xlLandscape

    // Set the printer as the default and confirm with Excel
    SetDefaultPrinter(PrinterName);

    Sheet.PageSetup.PaperSize := DesiredPaperSize;

  finally
    Ini.Free;
  end;
end;

procedure TFormMain.PrintImagePreview;
const
  LabelWidthMM = 90; // Label width in mm
  LabelHeightMM = 45; // Label height in mm
  DPI = 203; // Set this to your printer's DPI (203 or 300 typically)
var
  LabelWidthPx, LabelHeightPx: Integer;
  PrinterRect: TRect;
begin
  // Calculate label dimensions in pixels based on DPI
  LabelWidthPx := Round(LabelWidthMM * DPI / 25.4);
  LabelHeightPx := Round(LabelHeightMM * DPI / 25.4);

  // Begin the print job
  Printer.BeginDoc;
  try
    // Define the rectangle on the printer's canvas where the image will be drawn
    PrinterRect := Rect(0, 0, LabelWidthPx, LabelHeightPx);

    // Draw the TImage component's bitmap directly to the printer canvas, scaling as needed
    // Printer.Canvas.StretchDraw(PrinterRect, ImagePreview.Picture.Bitmap);

  finally
    Printer.EndDoc; // End the print job
  end;
end;

procedure TFormMain.SetDefaultPrinter(const PrinterName: string);
begin
  if not Winapi.WinSpool.SetDefaultPrinter(PChar(PrinterName)) then
    ShowMessage('Failed to set default printer: ' + PrinterName);
end;

// For debug
procedure TFormMain.AutoPopulateBarcodes;
const
  Barcodes: array [0 .. 2] of string = ('3179222', '3179223', '3179224');
var
  i: Integer;
begin
  for i := Low(Barcodes) to High(Barcodes) do
  begin
    edtbarcode.Text := Barcodes[i];
    ProcessBarcode;
  end;
end;

end.
