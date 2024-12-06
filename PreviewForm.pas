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
  System.ImageList, Vcl.ImgList, Vcl.Menus, Vcl.StdCtrls,Math;

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
const
  LabelWidthMM = 90; // Label width in millimeters
  LabelHeightMM = 45; // Label height in millimeters
var
  SourceBitmap, AdjustedBitmap: TBitmap;
  LabelWidthPx, LabelHeightPx,BaseDPI,TargetDPI: Integer;
  ScaleFactor,FontScaleFactor :Double;
begin
    BaseDPI := 203;
    TargetDPI := StrToIntDef(DataModuleCIMT.ReadSetting('Printer', 'DPI', '203'), 203);

    ScaleFactor := TargetDPI / BaseDPI;
    FontScaleFactor := 1 + (ScaleFactor - 1) / 30; // Smooth scaling for fonts

  // Calculate the pixel dimensions based on the target DPI
  LabelWidthPx := Round(LabelWidthMM * TargetDPI / 25.4);
  LabelHeightPx := Round(LabelHeightMM * TargetDPI / 25.4);

  FImages := TList<TBitmap>.Create;

  for SourceBitmap in Images do
  begin
    AdjustedBitmap := TBitmap.Create;
    try
      AdjustedBitmap.SetSize(LabelWidthPx, LabelHeightPx);
      AdjustedBitmap.Canvas.Brush.Color := clWhite;
      AdjustedBitmap.Canvas.FillRect(Rect(0, 0, LabelWidthPx, LabelHeightPx));

      // Set the font DPI for consistent font rendering
      AdjustedBitmap.Canvas.Font.PixelsPerInch := TargetDPI;
      AdjustedBitmap.Canvas.Font.Name := 'Arial';
      AdjustedBitmap.Canvas.Font.Size := 8;
      AdjustedBitmap.Canvas.StretchDraw(Rect(0, 0, LabelWidthPx, LabelHeightPx), SourceBitmap);

      // Add the adjusted bitmap to FImages
      FImages.Add(AdjustedBitmap);
    except
      AdjustedBitmap.Free;
      raise;
    end;
  end;

  // Update page information
  TotalPages := FImages.Count;
  CurrentPage := 1;

  edtFirst.Text := '1';
  edtLast.Text := IntToStr(TotalPages);
  DisplayCurrentPage;
end;

procedure TFormPreview.DisplayCurrentPage;
var
  DisplayBitmap: TBitmap;
  WidthScale, HeightScale, ScaleFactor: Single;
begin
  if (CurrentPage >= 1) and (CurrentPage <= TotalPages) then
  begin
    // Calculate the scale factors for both width and height
    WidthScale := 575 / FImages[CurrentPage - 1].Width;
    HeightScale := 288 / FImages[CurrentPage - 1].Height;
    ScaleFactor := Min(WidthScale, HeightScale);
    DisplayBitmap := TBitmap.Create;
    try
      DisplayBitmap.SetSize(Round(FImages[CurrentPage - 1].Width * ScaleFactor),
                            Round(FImages[CurrentPage - 1].Height * ScaleFactor));
      DisplayBitmap.Canvas.StretchDraw(Rect(0, 0, DisplayBitmap.Width, DisplayBitmap.Height), FImages[CurrentPage - 1]);
      ImagePreview.Picture.Assign(DisplayBitmap);
    finally
      DisplayBitmap.Free;
    end;

    lblPagetoprint.Caption := Format('Page %d of %d', [CurrentPage, TotalPages]);
  end;

  btnPreviousPage.Enabled := CurrentPage > 1;
  btnNextPage.Enabled := CurrentPage < TotalPages;
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
  if not SelectPrinterByName(DataModuleCIMT.ReadSetting('Printer',
    'Name', '')) then
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
