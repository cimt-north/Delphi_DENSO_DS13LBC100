unit SettingForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, System.IOUtils,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.WinXCtrls, IniFiles;

type
  TFormSetting = class(TForm)
    lblDebugSQL: TLabel;
    lblUserRESTAPI: TLabel;
    tgsDebugSQL: TToggleSwitch;
    tgsUseRESTAPI: TToggleSwitch;
    btnSave: TButton;
    btnCancel: TButton;
    lblUser: TLabel;
    lblPass: TLabel;
    edtUser: TEdit;
    edtPass: TEdit;
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure tgsUseRESTAPIClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure SaveSettingToIni;
    procedure UpdateAuthVisibility;

    { Private declarations }
  public
    procedure LoadSetting;
    { Public declarations }
  end;

var
  FormSetting: TFormSetting;

implementation

{$R *.dfm}

procedure TFormSetting.btnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TFormSetting.btnSaveClick(Sender: TObject);
begin
  SaveSettingToIni;
  ModalResult := mrOk; // Only close the form if Setting are valid
end;

procedure TFormSetting.FormCreate(Sender: TObject);
begin
  LoadSetting; // Load Setting when the form is created
  UpdateAuthVisibility;
end;

procedure TFormSetting.LoadSetting;
var
  IniFile: TIniFile;
  IniFileName: string;
begin
  // Load Setting from the INI file
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\' +
    ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
  IniFile := TIniFile.Create(IniFileName);
  try
    edtUser.Text := IniFile.ReadString('Setting', 'User', 'admin');
    edtPass.Text := IniFile.ReadString('Setting', 'Pass', 'admin');
    tgsUseRESTAPI.State := TToggleSwitchState(IniFile.ReadInteger('Setting',
      'IsUseRESTAPI', Integer(tssOff)));
    tgsDebugSQL.State := TToggleSwitchState(IniFile.ReadInteger('Setting',
      'DebugSQL', Integer(tssOff)));
  finally
    IniFile.Free;
  end;
end;

procedure TFormSetting.SaveSettingToIni;
var
  IniFile: TIniFile;
  IniFileName, IniFileDir: string;
begin
  // Define the INI file name and directory
  IniFileDir := ExtractFilePath(Application.ExeName) + 'GRD\';
  IniFileName := IniFileDir + ChangeFileExt
    (ExtractFileName(Application.ExeName), '.ini');

  // Check if the directory exists, if not, create it
  if not DirectoryExists(IniFileDir) then
    ForceDirectories(IniFileDir);

  // Save Setting to the INI file
  IniFile := TIniFile.Create(IniFileName);
  try
    IniFile.WriteString('Setting', 'User', edtUser.Text);
    IniFile.WriteString('Setting', 'Pass', edtPass.Text);
    IniFile.WriteInteger('Setting', 'IsUseRESTAPI',
      Integer(tgsUseRESTAPI.State));
    IniFile.WriteInteger('Setting', 'DebugSQL', Integer(tgsDebugSQL.State));
  finally
    IniFile.Free;
  end;
end;

procedure TFormSetting.tgsUseRESTAPIClick(Sender: TObject);
begin
  UpdateAuthVisibility;
end;

procedure TFormSetting.UpdateAuthVisibility;
begin
  // Hide or show the user and password fields based on the REST API toggle state
  if tgsUseRESTAPI.State = tssOn then
  begin
    lblUser.Visible := True;
    lblPass.Visible := True;
    edtUser.Visible := True;
    edtPass.Visible := True;
  end
  else
  begin
    lblUser.Visible := False;
    lblPass.Visible := False;
    edtUser.Visible := False;
    edtPass.Visible := False;
  end;
end;

end.
