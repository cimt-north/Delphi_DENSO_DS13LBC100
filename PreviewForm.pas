unit PreviewForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, UDataModule,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.jpeg,
  Vcl.Imaging.pngimage,
  VclTee.TeeGDIPlus, VclTee.TeeProcs, VclTee.TeePreviewPanel, Vcl.OleCtrls,
  IniFiles, Printers, System.Generics.Collections,
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
    edtFirst: TEdit;
    edtLast: TEdit;
    btnPrintOut: TSpeedButton;
    PrintoutP1: TMenuItem;
    CloseX1: TMenuItem;
    ImagePreview: TImage;
    lblPrintFrom: TLabel;
    lblPrintTo: TLabel;
    procedure btnNextPageClick(Sender: TObject);
    procedure btnPreviousPageClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure edtLastClick(Sender: TObject);
    procedure edtFirstClick(Sender: TObject);
    procedure CloseX1Click(Sender: TObject);
    procedure PrintoutP1Click(Sender: TObject);
    procedure btnPrintOutClick(Sender: TObject);
  private
    FImages: TList<TBitmap>; // Store multiple images
    procedure DisplayCurrentPage;
    procedure PrintLabels;
    { Private declarations }

  public
    CurrentPage: Integer;
    TotalPages: Integer;
    BaseFileName: string;
    procedure LoadImages(Images: TList<TBitmap>);

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
    DisplayCurrentPage;
  end;
end;

procedure TFormPreview.LoadImages(Images: TList<TBitmap>);
begin
  FImages := TList<TBitmap>.Create;
  FImages.AddRange(Images); // Store the images
  TotalPages := FImages.Count;
  CurrentPage := 1;

  edtFirst.Text := '1';
  edtLast.Text := IntToStr(TotalPages);
  DisplayCurrentPage;
end;

procedure TFormPreview.PrintoutP1Click(Sender: TObject);
begin
  PrintLabels;
end;

procedure TFormPreview.btnPreviousPageClick(Sender: TObject);
begin
  if CurrentPage > 1 then
  begin
    Dec(CurrentPage);
    DisplayCurrentPage;
  end;
end;

procedure TFormPreview.btnPrintOutClick(Sender: TObject);
begin
  PrintLabels;
end;

procedure TFormPreview.DisplayCurrentPage;
begin
  if (CurrentPage >= 1) and (CurrentPage <= TotalPages) then
  begin
    ImagePreview.Picture.Assign(FImages[CurrentPage - 1]); // Show current image
    lblPagetoprint.Caption := Format('Page %d of %d',
      [CurrentPage, TotalPages]);
  end;
  btnPreviousPage.Enabled := CurrentPage > 1;
  btnNextPage.Enabled := CurrentPage < TotalPages;
end;

procedure TFormPreview.PrintLabels;
  function SelectPrinterByName(const PrinterName: string): Boolean;
  var
    i: Integer;
  begin
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

    // If not found, show a message
    ShowMessage('Printer not found: ' + PrinterName);
  end;

var
  FirstPage, LastPage, i: Integer;
begin
  FirstPage := StrToIntDef(edtFirst.Text, 1);
  // Read the start page from edtFirst
  LastPage := StrToIntDef(edtLast.Text, TotalPages);
  // Read the end page from edtLast

  if (FirstPage < 1) or (LastPage > TotalPages) or (FirstPage > LastPage) then
  begin
    ShowMessage('Invalid page range.');
    Exit;
  end;

  // Select the printer
  if not SelectPrinterByName(DataModuleCIMT.ReadSetting('Settings',
    'Printer', '')) then
  begin
    Exit; // Stop if the printer is not found
  end;

  for i := FirstPage to LastPage do
  begin
    Printer.BeginDoc;
    try
      Printer.Canvas.Draw(0, 0, FImages[i - 1]);
    finally
      Printer.EndDoc;
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

end.
