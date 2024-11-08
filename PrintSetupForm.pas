unit PrintSetupForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Printers, IniFiles, System.StrUtils, Winapi.WinSpool;

type
  TFormPrintSetup = class(TForm)
    chkDefaultPrinter: TCheckBox;
    lblprinter: TLabel;
    lblpapersize: TLabel;
    pnbottom: TPanel;
    btnClose: TButton;
    cbbPrinter: TComboBox;
    cbbPapersize: TComboBox;
    rdgPaperStyle: TRadioGroup;
    rdbVertical: TRadioButton;
    rdbHorizontal: TRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure chkDefaultPrinterClick(Sender: TObject);
    procedure cbbPrinterChange(Sender: TObject);
  private
    procedure PopulatePrinters;
    procedure PopulatePaperSizes(const PrinterName: string);
    procedure WriteLog;
    procedure ReadSettings;
    procedure SetDefaultPrinter;
  public
  end;

var
  FormPrintSetup: TFormPrintSetup;

implementation

{$R *.dfm}


procedure TFormPrintSetup.FormCreate(Sender: TObject);
begin
  // Populate printer list and paper sizes
  PopulatePrinters;

  // Set default printer selection if needed
  if chkDefaultPrinter.Checked then
    SetDefaultPrinter;

  // Load saved settings from the INI file
  ReadSettings;
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
    PopulatePaperSizes(Printer.Printers[Printer.PrinterIndex]);
  end;
end;

procedure TFormPrintSetup.PopulatePaperSizes(const PrinterName: string);
var
  hPrinter: THandle;
  PaperSizes: array of Word;
  PaperNames: array of array[0..63] of Char;
  NumPaperSizes, i: Integer;
begin
  // Open the printer
  if OpenPrinter(PChar(PrinterName), hPrinter, nil) then
  try
    // Get the number of supported paper sizes
    NumPaperSizes := DeviceCapabilities(PChar(PrinterName), nil, DC_PAPERS, nil, nil);
    if NumPaperSizes > 0 then
    begin
      // Set the size of arrays based on the number of paper sizes
      SetLength(PaperSizes, NumPaperSizes);
      SetLength(PaperNames, NumPaperSizes);

      // Retrieve the list of supported paper sizes and names
      DeviceCapabilities(PChar(PrinterName), nil, DC_PAPERS, @PaperSizes[0], nil);
      DeviceCapabilities(PChar(PrinterName), nil, DC_PAPERNAMES, @PaperNames[0], nil);

      // Clear and populate the paper sizes combo box
      cbbPapersize.Items.Clear;
      for i := 0 to NumPaperSizes - 1 do
      begin
        cbbPapersize.Items.Add(Format('%s (%d)', [StrPas(PaperNames[i]), PaperSizes[i]]));
      end;

      // Optionally, set the first item as selected
      if cbbPapersize.Items.Count > 0 then
        cbbPapersize.ItemIndex := 0;
    end
    else
      ShowMessage('No supported paper sizes found for this printer.');
  finally
    ClosePrinter(hPrinter);
  end
  else
    ShowMessage('Could not open printer: ' + PrinterName);
end;

procedure TFormPrintSetup.SetDefaultPrinter;
begin
  // Set cbbPrinter to the default system printer and disable it
  cbbPrinter.ItemIndex := Printer.PrinterIndex;
  cbbPrinter.Enabled := False;
  PopulatePaperSizes(Printer.Printers[Printer.PrinterIndex]);
end;

procedure TFormPrintSetup.cbbPrinterChange(Sender: TObject);
begin
  // Update paper sizes based on selected printer
  if cbbPrinter.ItemIndex >= 0 then
    PopulatePaperSizes(cbbPrinter.Items[cbbPrinter.ItemIndex]);
end;

procedure TFormPrintSetup.WriteLog;
var
  Ini: TIniFile;
  IniFileName, PaperStyle, PaperSizeText: string;
  PaperSizeNum: Integer;
  StartPos, EndPos: Integer;
begin
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\DS13LBC100_printer.ini';
  Ini := TIniFile.Create(IniFileName);
  try
    // Save settings to the INI file
    Ini.WriteBool('Settings', 'UseDefaultPrinter', chkDefaultPrinter.Checked);
    Ini.WriteString('Settings', 'Printer', cbbPrinter.Text);
    Ini.WriteString('Settings', 'PaperSize', cbbPapersize.Text);

    // Extract the last number in parentheses from PaperSize text
    PaperSizeText := cbbPapersize.Text;
    StartPos := LastDelimiter('(', PaperSizeText); // Find the last '('
    EndPos := LastDelimiter(')', PaperSizeText);   // Find the last ')'

    // Only proceed if valid positions were found for both parentheses
    if (StartPos > 0) and (EndPos > StartPos) then
    begin
      // Extract the substring between the parentheses and convert it to an integer
      PaperSizeNum := StrToIntDef(Trim(Copy(PaperSizeText, StartPos + 1, EndPos - StartPos - 1)), -1);
    end
    else
      PaperSizeNum := -1; // Default if parsing fails

    // Write PaperSizeNum to the INI file
    Ini.WriteInteger('Settings', 'PaperSizeNum', PaperSizeNum);

    // Save paper style
    if rdbVertical.Checked then
      PaperStyle := 'Vertical'
    else
      PaperStyle := 'Horizontal';
    Ini.WriteString('Settings', 'PaperStyle', PaperStyle);
  finally
    Ini.Free;
  end;
end;




procedure TFormPrintSetup.ReadSettings;
var
  Ini: TIniFile;
  IniFileName: string;
  SavedPrinter, SavedPaperSize, PaperStyle: string;
  I: Integer;
begin
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\DS13LBC100_printer.ini';
  Ini := TIniFile.Create(IniFileName);
  try
    // Load settings from the INI file
    chkDefaultPrinter.Checked := Ini.ReadBool('Settings', 'UseDefaultPrinter', false);

    SavedPrinter := Ini.ReadString('Settings', 'Printer', '');
    if SavedPrinter <> '' then
    begin
      I := cbbPrinter.Items.IndexOf(SavedPrinter);
      if I >= 0 then
      begin
        cbbPrinter.ItemIndex := I;
        PopulatePaperSizes(SavedPrinter);
      end;
    end;

    SavedPaperSize := Ini.ReadString('Settings', 'PaperSize', '');
    if SavedPaperSize <> '' then
    begin
      I := cbbPapersize.Items.IndexOf(SavedPaperSize);
      if I >= 0 then
        cbbPapersize.ItemIndex := I;
    end;

    PaperStyle := Ini.ReadString('Settings', 'PaperStyle', 'Horizontal');
    if PaperStyle = 'Vertical' then
      rdbVertical.Checked := True
    else
      rdbHorizontal.Checked := True;
  finally
    Ini.Free;
  end;
end;



procedure TFormPrintSetup.chkDefaultPrinterClick(Sender: TObject);
begin
  // If Use Default Printer is checked, set the default printer and disable the combo box
  if chkDefaultPrinter.Checked then
    SetDefaultPrinter
  else
    cbbPrinter.Enabled := True; // Enable cbbPrinter if the checkbox is unchecked
end;

procedure TFormPrintSetup.btnCloseClick(Sender: TObject);
begin
  WriteLog; // Save settings before closing
  Close;    // Close the form when the Close button is clicked
end;

procedure TFormPrintSetup.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  WriteLog; // Ensure settings are saved on form close
end;

end.

