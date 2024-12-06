unit UDataModule;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Data.DB,
  Uni, MemDS, DBAccess, REST.Client, System.JSON, System.Net.HttpClient,
  System.NetEncoding, Vcl.Forms, UniProvider, OracleUniProvider, IniFiles,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TDataModuleCIMT = class(TDataModule)
    UniConnection1: TUniConnection;
    UniQuery1: TUniQuery;
    OracleUniProvider1: TOracleUniProvider;
    FDMemTable1: TFDMemTable;
  private
    function LoadOracleParameters: string;
    function LoadRestAPIParameters: string;
    function GetSessionID(const DoctorURL: string): string;
    function GetSQLJson(const SQLText: string): string;
    function Base64Encode(const InputStr: string): string;

  public
    constructor Create(AOwner: TComponent); reintroduce;
    function ReadSetting(const Section, Key, DefaultValue: string): string;
    procedure WriteSetting(const Section, Key, Value: string);
    function InitializeConnection(IsRESTAPI: Boolean): string;
    procedure FetchDataFromOracle(const SQLText: string; MemTable: TFDMemTable);
    function FetchDataFromREST(const SQLText: string;
      MemTable: TFDMemTable): string;
  end;

var
  DataModuleCIMT: TDataModuleCIMT;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}
{$R *.dfm}

constructor TDataModuleCIMT.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

function TDataModuleCIMT.LoadOracleParameters: string;
var
  DirectDBName, User, Pass: string;
  IniFile: TIniFile;
  IniFileName: string;
begin
  // Use Setup folder for Oracle parameters
  IniFileName := ExtractFilePath(Application.ExeName) + '/Setup/SetUp.Ini';
  IniFile := TIniFile.Create(IniFileName);
  try
    // Retrieve Oracle parameters
    DirectDBName := IniFile.ReadString('Setting', 'DIRECTDBNAME', '');
    User := IniFile.ReadString('Setting', 'USERNAME', '');
    Pass := IniFile.ReadString('Setting', 'PASSWORD', '');

    Result := DirectDBName + ' : ' + User;

    UniConnection1.ProviderName := 'Oracle';
    UniConnection1.SpecificOptions.Values['Direct'] := 'True';
    UniConnection1.Server := DirectDBName;
    UniConnection1.Username := User;
    UniConnection1.Password := Pass;

    if not UniConnection1.Connected then
      UniConnection1.Connect;
  finally
    IniFile.Free;
  end;
end;

function TDataModuleCIMT.LoadRestAPIParameters: string;
var
  IniFile: TMemIniFile;
  IniFileName: string;
begin
  // Use Setup folder for REST parameters
  IniFileName := ExtractFilePath(Application.ExeName) + '/Setup/SetUp.Ini';
  IniFile := TMemIniFile.Create(IniFileName);
  try
    Result := IniFile.ReadString('Setting', 'DOCTOR_URL', '');
  finally
    IniFile.Free;
  end;
end;

function TDataModuleCIMT.GetSessionID(const DoctorURL: string): string;
var
  HttpClient: THTTPClient;
  Response: IHTTPResponse;
  JSONValue: TJSONValue;
  DataJSON: TJSONObject;
  User, Pass: string;
begin
  // Retrieve User and Pass
  User := ReadSetting('Setting', 'User', 'admin');
  Pass := ReadSetting('Setting', 'Pass', 'admin');

  HttpClient := THTTPClient.Create;
  try
    DataJSON := TJSONObject.Create;
    try
      DataJSON.AddPair('loginID', User);
      DataJSON.AddPair('passWord', Pass);

      Response := HttpClient.Post(DoctorURL + '/api/login',
        TStringStream.Create(DataJSON.ToString), nil);

      if Response.StatusCode = 200 then
      begin
        JSONValue := TJSONObject.ParseJSONValue
          (Response.ContentAsString(TEncoding.UTF8));
        try
          if JSONValue.TryGetValue('sessionID', Result) then
            Exit(Result);
        finally
          JSONValue.Free;
        end;
      end
      else
        raise Exception.Create('Error: ' + Response.StatusCode.ToString + ' - '
          + Response.StatusText);
    finally
      DataJSON.Free;
    end;
  finally
    HttpClient.Free;
  end;
  Result := '';
end;

function TDataModuleCIMT.GetSQLJson(const SQLText: string): string;
var
  PayloadArray: TJSONArray;
  SQLData, TantoCD: TJSONObject;
begin
  PayloadArray := TJSONArray.Create;
  try
    SQLData := TJSONObject.Create;
    SQLData.AddPair('key', 'SQLDATA');
    SQLData.AddPair('value1', Base64Encode(SQLText));
    SQLData.AddPair('value2', TJSONNull.Create);
    PayloadArray.Add(SQLData);

    TantoCD := TJSONObject.Create;
    TantoCD.AddPair('key', 'TANTOCD');
    TantoCD.AddPair('value1', 'admin');
    TantoCD.AddPair('value2', '');
    PayloadArray.Add(TantoCD);

    Result := PayloadArray.ToJSON;
  finally
    PayloadArray.Free;
  end;
end;

function TDataModuleCIMT.Base64Encode(const InputStr: string): string;
begin
  Result := TNetEncoding.Base64.Encode(InputStr);
end;

procedure TDataModuleCIMT.FetchDataFromOracle(const SQLText: string;
  MemTable: TFDMemTable);
var
  Field: TField;
  MemField: TField;
begin
  if not UniConnection1.Connected then
    UniConnection1.Connect;

  UniQuery1.SQL.Text := SQLText;

  try
    UniQuery1.Open;

    // Clear and setup the TFDMemTable structure
    MemTable.Close; // Close the MemTable if it is open
    MemTable.Fields.Clear; // Clear existing fields
    MemTable.FieldDefs.Clear; // Clear field definitions

    // Define the fields in the MemTable based on UniQuery
    for Field in UniQuery1.Fields do
    begin
      // Add field definitions to the MemTable
      MemTable.FieldDefs.Add(Field.FieldName, Field.DataType, Field.Size,
        Field.Required);
    end;

    // Create the fields in the MemTable
    MemTable.CreateDataSet;

    // Populate the MemTable with data from UniQuery
    UniQuery1.First;
    while not UniQuery1.Eof do
    begin
      MemTable.Append;
      for Field in UniQuery1.Fields do
      begin
        MemField := MemTable.FieldByName(Field.FieldName);
        MemField.Value := Field.Value;
      end;
      MemTable.Post; // Save the new record in the MemTable
      UniQuery1.Next;
    end;
  finally
    UniQuery1.Close;
  end;
end;

function TDataModuleCIMT.FetchDataFromREST(const SQLText: string;
  MemTable: TFDMemTable): string;
var
  Client: THTTPClient;
  Response: IHTTPResponse;
  JSONValue: TJSONValue;
  DataArray: TJSONArray;
  SessionID, DoctorURL: string;
  I: Integer;
  Item: TJSONValue;
  Pair: TJSONPair;
  FieldDef: TFieldDef;
  FieldType: TFieldType;
  StringValue: string;
begin
  // Retrieve User, Pass and DoctorURL
  DoctorURL := LoadRestAPIParameters();
  SessionID := GetSessionID(DoctorURL);

  Client := THTTPClient.Create;
  try
    Client.ContentType := 'application/json';
    Client.CustomHeaders['SESSIONID'] := SessionID;
    Client.CustomHeaders['Accept'] := 'application/json';
    Client.CustomHeaders['Authorization'] := 'Bearer ' + SessionID;

    Response := Client.Post(DoctorURL + '/api/sql/sqltool/open',
      TStringStream.Create(GetSQLJson(SQLText), TEncoding.UTF8), nil);

    if Response.StatusCode = 200 then
    begin
      JSONValue := TJSONObject.ParseJSONValue
        (Response.ContentAsString(TEncoding.UTF8));
      try
        if JSONValue is TJSONObject then
        begin
          DataArray := TJSONObject(JSONValue).GetValue('data') as TJSONArray;

          // Clear and setup the TFDMemTable structure
          MemTable.Close;
          MemTable.Fields.Clear;
          MemTable.FieldDefs.Clear;

          // Define fields in the MemTable based on the first JSON object
          if DataArray.Count > 0 then
          begin
            for Pair in TJSONObject(DataArray.Items[0]) do
            begin
              // Determine the field type based on the JSON value type
              if Pair.JSONValue is TJSONString then
                FieldType := ftWideString
              else if Pair.JSONValue is TJSONNumber then
              begin
                if TJSONNumber(Pair.JSONValue).AsInt64 = TJSONNumber
                  (Pair.JSONValue).AsDouble then
                  FieldType := ftInteger
                else
                  FieldType := ftFloat;
              end
              else if Pair.JSONValue is TJSONBool then
                FieldType := ftBoolean
              else if Pair.JSONValue is TJSONNull then
                FieldType := ftWideString
              else
                FieldType := ftWideString;
              // Default to WideString for unknown types

              FieldDef := MemTable.FieldDefs.AddFieldDef;
              FieldDef.Name := Pair.JsonString.Value;
              FieldDef.DataType := FieldType;
              FieldDef.Size := 255; // Set a reasonable size for string fields
            end;

            // Create the dataset in memory
            MemTable.CreateDataSet;

            // Populate the MemTable with data from the JSON array
            for I := 0 to DataArray.Count - 1 do
            begin
              Item := DataArray.Items[I];
              MemTable.Append;
              for Pair in TJSONObject(Item) do
              begin
                if FieldType = ftWideString then
                begin
                  StringValue := TEncoding.UTF8.GetString
                    (TEncoding.UTF8.GetBytes(Pair.JSONValue.Value));
                  MemTable.FieldByName(Pair.JsonString.Value).AsWideString :=
                    StringValue;
                end
                else
                begin
                  MemTable.FieldByName(Pair.JsonString.Value).Value :=
                    Pair.JSONValue.Value;
                end;
              end;
              MemTable.Post;
            end;
          end;
        end;
      finally
        JSONValue.Free;
      end;
    end
    else
    begin
      Result := 'Error: ' + Response.StatusCode.ToString + ' - ' +
        Response.StatusText;
      Exit(Result);
    end;
  finally
    Client.Free;
  end;
  Result := 'Success';
end;

function TDataModuleCIMT.InitializeConnection(IsRESTAPI: Boolean): string;
begin
  if IsRESTAPI then
    Result := LoadRestAPIParameters
  else
    Result := LoadOracleParameters;
end;

// Centralized method for reading a setting from the INI file
function TDataModuleCIMT.ReadSetting(const Section, Key,
  DefaultValue: string): string;
var
  IniFile: TIniFile;
  IniFileName, IniFileDir: string;
begin
  IniFileDir := ExtractFilePath(Application.ExeName) + 'GRD\';
  IniFileName := IniFileDir + ChangeFileExt
    (ExtractFileName(Application.ExeName), '.ini');

  // Ensure the directory exists
  if not DirectoryExists(IniFileDir) then
    ForceDirectories(IniFileDir);

  IniFile := TIniFile.Create(IniFileName);
  try
    Result := IniFile.ReadString(Section, Key, DefaultValue);
  finally
    IniFile.Free;
  end;
end;

procedure TDataModuleCIMT.WriteSetting(const Section, Key, Value: string);
var
  IniFile: TIniFile;
  IniFileName, IniFileDir: string;
begin
  IniFileDir := ExtractFilePath(Application.ExeName) + 'GRD\';
  IniFileName := IniFileDir + ChangeFileExt
    (ExtractFileName(Application.ExeName), '.ini');

  // Ensure the directory exists
  if not DirectoryExists(IniFileDir) then
    ForceDirectories(IniFileDir);

  IniFile := TIniFile.Create(IniFileName);
  try
    IniFile.WriteString(Section, Key, Value);
  finally
    IniFile.Free;
  end;
end;

end.
