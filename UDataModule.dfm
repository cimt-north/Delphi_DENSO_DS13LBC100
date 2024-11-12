object DataModuleCIMT: TDataModuleCIMT
  Height = 960
  Width = 1280
  PixelsPerInch = 192
  object UniConnection1: TUniConnection
    ProviderName = 'Oracle'
    Left = 1008
    Top = 416
  end
  object UniQuery1: TUniQuery
    Connection = UniConnection1
    Left = 832
    Top = 416
  end
  object OracleUniProvider1: TOracleUniProvider
    Left = 640
    Top = 496
  end
  object FDMemTable1: TFDMemTable
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvSilentMode]
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    Left = 1168
    Top = 416
  end
end
