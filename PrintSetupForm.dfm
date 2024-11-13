object FormPrintSetup: TFormPrintSetup
  Left = 0
  Top = 0
  Margins.Left = 4
  Margins.Top = 4
  Margins.Right = 4
  Margins.Bottom = 4
  Caption = 'Print Setup'
  ClientHeight = 175
  ClientWidth = 377
  Color = 16706018
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 20
  object lblprinter: TLabel
    Left = 30
    Top = 38
    Width = 43
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Printer'
  end
  object lblLabelSize: TLabel
    Left = 110
    Top = 73
    Width = 184
    Height = 20
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = 'Label Size = 90mm x 45mm'
    Color = 16706275
    ParentColor = False
  end
  object pnbottom: TPanel
    Left = 0
    Top = 125
    Width = 377
    Height = 50
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Align = alBottom
    Color = 16706018
    ParentBackground = False
    TabOrder = 0
    object btnClose: TButton
      Left = 260
      Top = 13
      Width = 91
      Height = 25
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Close'
      TabOrder = 0
      OnClick = btnCloseClick
    end
  end
  object cbbPrinter: TComboBox
    Left = 110
    Top = 34
    Width = 241
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    TabOrder = 1
  end
end
