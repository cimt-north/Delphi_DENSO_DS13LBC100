object FormSetting: TFormSetting
  Left = 0
  Top = 0
  Margins.Left = 4
  Margins.Top = 4
  Margins.Right = 4
  Margins.Bottom = 4
  Caption = 'FormSetting'
  ClientHeight = 224
  ClientWidth = 406
  Color = 16635850
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 20
  object lblDebugSQL: TLabel
    Left = 30
    Top = 30
    Width = 75
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Debug SQL'
  end
  object lblUserRESTAPI: TLabel
    Left = 34
    Top = 70
    Width = 87
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Use REST API'
  end
  object lblUser: TLabel
    Left = 250
    Top = 70
    Width = 29
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'User'
  end
  object lblPass: TLabel
    Left = 250
    Top = 110
    Width = 27
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Pass'
  end
  object tgsDebugSQL: TToggleSwitch
    Left = 130
    Top = 24
    Width = 90
    Height = 25
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    SwitchHeight = 25
    SwitchWidth = 63
    TabOrder = 0
    ThumbWidth = 19
  end
  object tgsUseRESTAPI: TToggleSwitch
    Left = 130
    Top = 70
    Width = 90
    Height = 25
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    SwitchHeight = 25
    SwitchWidth = 63
    TabOrder = 1
    ThumbWidth = 19
    OnClick = tgsUseRESTAPIClick
  end
  object btnSave: TButton
    Left = 170
    Top = 170
    Width = 94
    Height = 31
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Save'
    TabOrder = 2
    OnClick = btnSaveClick
  end
  object btnCancel: TButton
    Left = 290
    Top = 170
    Width = 94
    Height = 31
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Cancel'
    TabOrder = 3
    OnClick = btnCancelClick
  end
  object edtUser: TEdit
    Left = 303
    Top = 66
    Width = 97
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    TabOrder = 4
  end
  object edtPass: TEdit
    Left = 303
    Top = 103
    Width = 97
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    PasswordChar = '*'
    TabOrder = 5
  end
end
