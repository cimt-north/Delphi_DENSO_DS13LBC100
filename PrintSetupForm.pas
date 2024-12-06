unit PrintSetupForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, UDataModule,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Printers,
  IniFiles, System.StrUtils, Winapi.WinSpool;

type
  TFormPrintSetup = class(TForm)
    lblprinter: TLabel;
    pnbottom: TPanel;
    btnClose: TButton;
    cbbPrinter: TComboBox;
    lblLabelSize: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure PopulatePrinters;
    procedure LoadPrinterSettings;
    procedure SavePrinterSettings;
    function GetPrinterDPI: Integer;
  public
  end;

var
  FormPrintSetup: TFormPrintSetup;

implementation

{$R *.dfm}

procedure TFormPrintSetup.FormCreate(Sender: TObject);
begin
  cbbPrinter.Style := csDropDownList;
  PopulatePrinters;
  LoadPrinterSettings;
end;

procedure TFormPrintSetup.PopulatePrinters;
var
  I: Integer;
begin
  // Populate cbbPrinter with available printers
  cbbPrinter.Items.Clear;
  for I := 0 to Printer.Printers.Count - 1 do
    cbbPrinter.Items.Add(Printer.Printers[I]);

  // Set the currently selected printer to the system default
  if Printer.Printers.Count > 0 then
  begin
    cbbPrinter.ItemIndex := Printer.PrinterIndex;
  end;
end;

function TFormPrintSetup.GetPrinterDPI: Integer;
var
  PrinterHandle: THandle;
  PrinterName: string;
  PrinterDC: HDC;
begin
  Result := 203; // Default DPI if retrieval fails
  PrinterName := cbbPrinter.Text;

  // Open the printer and get a handle to it
  if OpenPrinter(PChar(PrinterName), PrinterHandle, nil) then
  try
    // Create a device context for the printer
    PrinterDC := CreateDC('WINSPOOL', PChar(PrinterName), nil, nil);
    if PrinterDC <> 0 then
    try
      // Get the horizontal DPI of the printer
      Result := GetDeviceCaps(PrinterDC, LOGPIXELSX);
    finally
      DeleteDC(PrinterDC); // Delete the device context after use
    end;
  finally
    ClosePrinter(PrinterHandle); // Close the printer handle
  end;
end;

procedure TFormPrintSetup.SavePrinterSettings;
var
  PrinterDPI: Integer;
begin
  // Save the selected printer to the centralized settings
  DataModuleCIMT.WriteSetting('Printer', 'Name', cbbPrinter.Text);
  PrinterDPI := GetPrinterDPI;
  DataModuleCIMT.WriteSetting('Printer', 'DPI', IntToStr(PrinterDPI));
end;

procedure TFormPrintSetup.LoadPrinterSettings;
var
  SavedPrinter: string;
  I: Integer;
begin
  // Load the saved printer from the centralized settings
  SavedPrinter := DataModuleCIMT.ReadSetting('Printer', 'Name', '');
  if SavedPrinter <> '' then
  begin
    I := cbbPrinter.Items.IndexOf(SavedPrinter);
    if I >= 0 then
      cbbPrinter.ItemIndex := I;
  end;
end;

procedure TFormPrintSetup.btnCloseClick(Sender: TObject);
begin
  SavePrinterSettings; // Save settings before closing
  Close; // Close the form when the Close button is clicked
end;

procedure TFormPrintSetup.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SavePrinterSettings; // Ensure settings are saved on form close
end;

end.
