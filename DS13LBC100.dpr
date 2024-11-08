program DS13LBC100;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {FormMain},
  UDataModule in 'UDataModule.pas' {DataModuleCIMT: TDataModule},
  SettingForm in 'SettingForm.pas' {FormSetting},
  DetailForm in 'DetailForm.pas' {FormDetail},
  PrintSetupForm in 'PrintSetupForm.pas' {FormPrintSetup},
  PreviewForm in 'PreviewForm.pas' {FormPreview};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
