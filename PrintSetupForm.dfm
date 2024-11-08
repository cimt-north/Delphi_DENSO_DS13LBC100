object FormPrintSetup: TFormPrintSetup
  Left = 0
  Top = 0
  Caption = 'Print Setup'
  ClientHeight = 310
  ClientWidth = 555
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
    Left = 88
    Top = 56
    Width = 41
    Height = 15
    Caption = 'Printer'
  end
  object lblpapersize: TLabel
    Left = 68
    Top = 96
    Width = 61
    Height = 15
    Caption = 'Paper Size'
  end
  object chkDefaultPrinter: TCheckBox
    Left = 152
    Top = 16
    Width = 137
    Height = 17
    Caption = 'Use default printer'
    TabOrder = 0
    OnClick = chkDefaultPrinterClick
  end
  object pnbottom: TPanel
    Left = 0
    Top = 270
    Width = 555
    Height = 40
    Align = alBottom
    Color = 16706018
    ParentBackground = False
    TabOrder = 1
    ExplicitTop = 275
    object btnClose: TButton
      Left = 470
      Top = 10
      Width = 73
      Height = 20
      Caption = 'Close'
      TabOrder = 0
      OnClick = btnCloseClick
    end
  end
  object cbbPrinter: TComboBox
    Left = 152
    Top = 53
    Width = 193
    Height = 23
    TabOrder = 2
    OnChange = cbbPrinterChange
  end
  object cbbPapersize: TComboBox
    Left = 152
    Top = 93
    Width = 193
    Height = 23
    TabOrder = 3
  end
  object rdgPaperStyle: TRadioGroup
    Left = 152
    Top = 137
    Width = 121
    Height = 97
    Caption = 'Paper style'
    Color = 16706275
    ParentBackground = False
    ParentColor = False
    TabOrder = 4
  end
  object rdbVertical: TRadioButton
    Left = 168
    Top = 160
    Width = 81
    Height = 17
    Caption = 'Vertical'
    Color = 16706275
    ParentColor = False
    TabOrder = 5
  end
  object rdbHorizontal: TRadioButton
    Left = 168
    Top = 200
    Width = 81
    Height = 17
    Caption = 'Horizontal'
    Color = 16706275
    ParentColor = False
    TabOrder = 6
  end
end
