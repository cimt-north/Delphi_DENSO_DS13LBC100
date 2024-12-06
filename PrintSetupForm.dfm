object FormPrintSetup: TFormPrintSetup
  Left = 0
  Top = 0
  Caption = 'Print Setup'
  ClientHeight = 140
  ClientWidth = 300
  Color = 16706018
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnClose = FormClose
  OnCreate = FormCreate
  TextHeight = 15
  object lblprinter: TLabel
    Left = 24
    Top = 30
    Width = 35
    Height = 15
    Caption = 'Printer'
  end
  object lblLabelSize: TLabel
    Left = 88
    Top = 58
    Width = 145
    Height = 15
    Caption = 'Label Size = 90mm x 45mm'
    Color = 16706275
    ParentColor = False
  end
  object pnbottom: TPanel
    Left = 0
    Top = 100
    Width = 300
    Height = 40
    Align = alBottom
    Color = 16706018
    ParentBackground = False
    TabOrder = 0
    object btnClose: TButton
      Left = 208
      Top = 10
      Width = 73
      Height = 20
      Caption = 'Close'
      TabOrder = 0
      OnClick = btnCloseClick
    end
  end
  object cbbPrinter: TComboBox
    Left = 88
    Top = 27
    Width = 193
    Height = 23
    TabOrder = 1
  end
end
