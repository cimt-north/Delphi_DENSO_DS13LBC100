unit PreviewForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.jpeg, Vcl.Imaging.pngimage,
  VclTee.TeeGDIPlus, VCLTee.TeeProcs, VCLTee.TeePreviewPanel, Vcl.OleCtrls, IniFiles,
  SHDocVw, AcroPDFLib_TLB, Vcl.Buttons, Vcl.ComCtrls, Vcl.ToolWin,
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
    lblPage: TLabel;
    WebBrowser1: TWebBrowser;
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
    CurrentPage: Integer;
    TotalPages: Integer;
    BaseFileName: string;
  public
    procedure LoadImageFromFile(const FileName: string; lastpage: Integer);
    procedure readPrinterSetup;
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
    WebBrowser1.Navigate('file://' + BaseFileName + IntToStr(CurrentPage) + '.pdf'); // Use WebBrowser1
    btnPreviousPage.Enabled := True; // Enable the Previous button
    lblPage.Caption := 'Page: ' + IntToStr(CurrentPage);
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
    WebBrowser1.Navigate('file://' + BaseFileName + IntToStr(CurrentPage) + '.pdf'); // Use WebBrowser1
    btnNextPage.Enabled := True; // Enable the Next button
    lblPage.Caption := 'Page: ' + IntToStr(CurrentPage);
  end;

  if CurrentPage <= 1 then
  begin
    btnPreviousPage.Enabled := False; // Disable the Previous button if at the first page
  end;
end;

procedure TFormPreview.btnPrintOutClick(Sender: TObject);
var
  i, FirstPage, LastPage: Integer;
  FileName: string;
begin
  FirstPage := StrToIntDef(edtFirst.Text, 1); // Read first number from edtFirst
  LastPage := StrToIntDef(edtLast.Text, TotalPages); // Read last number from edtLast

  for i := FirstPage to LastPage do
  begin
    FileName := BaseFileName + IntToStr(i) + '.pdf';
    if FileExists(FileName) then
    begin
      WebBrowser1.Navigate('file://' + FileName); // Use WebBrowser1 for display
      // Add code to trigger the print dialog if needed
    end
    else
    begin
      ShowMessage('File not found: ' + FileName);
    end;
  end;
end;

procedure TFormPreview.CloseX1Click(Sender: TObject);
begin
   halt;
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
    lblPage.Caption := 'Page: ' + IntToStr(CurrentPage+1);
end;

procedure TFormPreview.LoadImageFromFile(const FileName: string; lastpage: Integer);
begin
  if FileExists(FileName) then
  begin
    BaseFileName := ExtractFilePath(FileName) + 'PreviewPage_';
    TotalPages := lastpage;
    edtLast.Text :=  inttostr(lastpage);
    CurrentPage := 1;
    WebBrowser1.Navigate('file://' + BaseFileName + IntToStr(CurrentPage) + '.pdf'); // Use WebBrowser1
    btnPreviousPage.Enabled := False; // Start with Previous button disabled
    btnNextPage.Enabled := (TotalPages > 1); // Enable Next if more than one page
  end
  else
    ShowMessage('File not found: ' + FileName);
end;

procedure TFormPreview.PrintoutP1Click(Sender: TObject);
var
  i, FirstPage, LastPage: Integer;
  FileName: string;
begin
  FirstPage := StrToIntDef(edtFirst.Text, 1); // Read first number from edtFirst
  LastPage := StrToIntDef(edtLast.Text, TotalPages); // Read last number from edtLast

  for i := FirstPage to LastPage do
  begin
    FileName := BaseFileName + IntToStr(i) + '.pdf';
    if FileExists(FileName) then
    begin
      WebBrowser1.Navigate('file://' + FileName); // Use WebBrowser1 for display
      // Add code to trigger the print dialog if needed
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
  PaperNames: array of array[0..63] of Char;
  NumPaperSizes, i ,DesiredPaperSize: Integer;
  hPrinter: THandle;
  SelectedPaperSize: Integer;
  ExcelApp: Variant;
begin
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\DS13LBC100_printer.ini';
  Ini := TIniFile.Create(IniFileName);
  try
    // Read settings from the INI file
    UseDefaultPrinter := Ini.ReadBool('Settings', 'UseDefaultPrinter', False);
    PaperStyle := Ini.ReadString('Settings', 'PaperStyle', 'Horizontal');
    DesiredPaperSize := Ini.ReadInteger('Settings', 'PaperSizeNum', 1);
    PrinterName := Ini.ReadString('Settings', 'Printer', '');
    //1  xlPortrait
    //2  xlLandscape

  finally
    Ini.Free;
  end;
end;

end.

