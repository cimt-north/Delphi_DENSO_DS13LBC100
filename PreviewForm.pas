unit PreviewForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.jpeg,
  Vcl.Imaging.pngimage,
  VclTee.TeeGDIPlus, VclTee.TeeProcs, VclTee.TeePreviewPanel, Vcl.OleCtrls,
  IniFiles,
  SHDocVw, Vcl.Buttons, Vcl.ComCtrls, Vcl.ToolWin,
  System.ImageList, Vcl.ImgList, Vcl.Menus, Vcl.StdCtrls;

type
  TFormPreview = class(TForm)
    Panel1: TPanel;
    clbToolButton: TCoolBar;
    tlbToolButton: TToolBar;
    btnNextPage: TSpeedButton;
    btnPreviousPage: TSpeedButton;
    btnClose: TSpeedButton;
    imlToolBar: TImageList;
    MainMenu1: TMainMenu;
    FileF1: TMenuItem;
    Panel2: TPanel;
    lblPagetoprint: TLabel;
    edtFrom: TEdit;
    edtto: TEdit;
    edtFirst: TEdit;
    edtLast: TEdit;
    btnPrintOut: TSpeedButton;
    PrintoutP1: TMenuItem;
    CloseX1: TMenuItem;
    ImagePreview: TImage;
    procedure btnNextPageClick(Sender: TObject);
    procedure btnPreviousPageClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure edtLastClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnPrintOutClick(Sender: TObject);
    procedure edtFirstClick(Sender: TObject);
    procedure CloseX1Click(Sender: TObject);
    procedure PrintoutP1Click(Sender: TObject);
  private
    { Private declarations }

  public
    CurrentPage: Integer;
    TotalPages: Integer;
    BaseFileName: string;
    procedure readPrinterSetup;
    procedure LoadImageFromBitmap(Bitmap: TBitmap; PageNumber: Integer);

  end;

var
  FormPreview: TFormPreview;

implementation

{$R *.dfm}

procedure TFormPreview.btnCloseClick(Sender: TObject);
begin
  Close; // Close the form instead of halting the application
end;

procedure TFormPreview.btnNextPageClick(Sender: TObject);
begin
  if CurrentPage < TotalPages then
  begin
    Inc(CurrentPage);
    btnPreviousPage.Enabled := True; // Enable the Previous button
  end;

  if CurrentPage >= TotalPages then
  begin
    btnNextPage.Enabled := False; // Disable the Next button if at the last page
  end;
end;

procedure TFormPreview.btnPreviousPageClick(Sender: TObject);
begin
  if CurrentPage > 1 then
  begin
    Dec(CurrentPage);
    btnNextPage.Enabled := True; // Enable the Next button
  end;

  if CurrentPage <= 1 then
  begin
    btnPreviousPage.Enabled := False;
    // Disable the Previous button if at the first page
  end;
end;

procedure TFormPreview.btnPrintOutClick(Sender: TObject);
var
  i, FirstPage, lastpage: Integer;
  FileName: string;
begin
  FirstPage := StrToIntDef(edtFirst.Text, 1); // Read first number from edtFirst
  lastpage := StrToIntDef(edtLast.Text, TotalPages);
  // Read last number from edtLast

  for i := FirstPage to lastpage do
  begin
    FileName := BaseFileName + IntToStr(i) + '.pdf';
    if FileExists(FileName) then
    begin

    end
    else
    begin
      ShowMessage('File not found: ' + FileName);
    end;
  end;
end;

procedure TFormPreview.CloseX1Click(Sender: TObject);
begin
  Close;
end;

procedure TFormPreview.edtFirstClick(Sender: TObject);
begin
  edtFirst.Color := RGB(252, 252, 164);
  edtLast.Color := clWhite;
end;

procedure TFormPreview.edtLastClick(Sender: TObject);
begin
  edtLast.Color := RGB(252, 252, 164);
  edtFirst.Color := clWhite;
end;

procedure TFormPreview.FormCreate(Sender: TObject);
begin
  readPrinterSetup;
  edtFirst.Text := '1';
end;

procedure TFormPreview.LoadImageFromBitmap(Bitmap: TBitmap;
  PageNumber: Integer);
begin
  ImagePreview.Picture.Assign(Bitmap); // Assign bitmap to ImagePreview
  lblPagetoprint.Caption := Format('Page %d of %d', [PageNumber, TotalPages]);
  // Update label with page number
end;

procedure TFormPreview.PrintoutP1Click(Sender: TObject);
var
  i, FirstPage, lastpage: Integer;
  FileName: string;
begin
  FirstPage := StrToIntDef(edtFirst.Text, 1); // Read first number from edtFirst
  lastpage := StrToIntDef(edtLast.Text, TotalPages);
  // Read last number from edtLast

  for i := FirstPage to lastpage do
  begin
    FileName := BaseFileName + IntToStr(i) + '.pdf';
    if FileExists(FileName) then
    begin

    end
    else
    begin
      ShowMessage('File not found: ' + FileName);
    end;
  end;
end;

procedure TFormPreview.readPrinterSetup;
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
    // 1  xlPortrait
    // 2  xlLandscape

  finally
    Ini.Free;
  end;
end;

end.
