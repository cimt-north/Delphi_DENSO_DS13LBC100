object DataModuleCIMT: TDataModuleCIMT
  Height = 480
  Width = 640
  object UniConnection1: TUniConnection
    ProviderName = 'Oracle'
    Left = 504
    Top = 208
  end
  object UniQuery1: TUniQuery
    Connection = UniConnection1
    Left = 416
    Top = 208
  end
  object OracleUniProvider1: TOracleUniProvider
    Left = 320
    Top = 248
  end
  object FDMemTable1: TFDMemTable
    FetchOptions.AssignedValues = [evMode]
    FetchOptions.Mode = fmAll
    ResourceOptions.AssignedValues = [rvSilentMode]
    ResourceOptions.SilentMode = True
    UpdateOptions.AssignedValues = [uvCheckRequired, uvAutoCommitUpdates]
    UpdateOptions.CheckRequired = False
    UpdateOptions.AutoCommitUpdates = True
    Left = 584
    Top = 208
  end
end
