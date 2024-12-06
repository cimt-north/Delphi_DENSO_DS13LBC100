unit SettingForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, System.IOUtils, UDataModule,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.WinXCtrls;

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
    tgsCAP: TToggleSwitch;
    lblClearAll: TLabel;
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
begin
  // Use centralized ReadSetting method to load values
  edtUser.Text := DataModuleCIMT.ReadSetting('Setting', 'User', 'admin');
  edtPass.Text := DataModuleCIMT.ReadSetting('Setting', 'Pass', 'admin');

  // Use StrToIntDef to safely convert the read integer settings for toggle switches
  tgsUseRESTAPI.State := TToggleSwitchState
    (StrToIntDef(DataModuleCIMT.ReadSetting('Setting', 'IsUseRESTAPI',
    IntToStr(Integer(tssOff))), Integer(tssOff)));

  tgsDebugSQL.State := TToggleSwitchState
    (StrToIntDef(DataModuleCIMT.ReadSetting('Setting', 'DebugSQL',
    IntToStr(Integer(tssOff))), Integer(tssOff)));
  tgsCAP.State := TToggleSwitchState
    (StrToIntDef(DataModuleCIMT.ReadSetting('Setting', 'AutoClear',
    IntToStr(Integer(tssOff))), Integer(tssOff)));

end;

procedure TFormSetting.SaveSettingToIni;
begin
  // Use centralized WriteSetting method to save values
  DataModuleCIMT.WriteSetting('Setting', 'User', edtUser.Text);
  DataModuleCIMT.WriteSetting('Setting', 'Pass', edtPass.Text);

  // Convert toggle switch states to strings for storage
  DataModuleCIMT.WriteSetting('Setting', 'IsUseRESTAPI',
    IntToStr(Integer(tgsUseRESTAPI.State)));
  DataModuleCIMT.WriteSetting('Setting', 'DebugSQL',
    IntToStr(Integer(tgsDebugSQL.State)));

  DataModuleCIMT.WriteSetting('Setting', 'AutoClear',
    IntToStr(Integer(tgsCAP.State)));

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
