unit DetailForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TFormDetail = class(TForm)
    stgDetail: TStringGrid;
  public
    procedure SetFilteredMemTable(MemTable: TFDMemTable);
  end;

var
  FormDetail: TFormDetail;

implementation

{$R *.dfm}

procedure TFormDetail.SetFilteredMemTable(MemTable: TFDMemTable);
var
  Row, Col: Integer;
begin
  // Populate StringGrid from the passed MemTable
  if not MemTable.IsEmpty then
  begin
    // Set the row and column count based on the MemTable structure
    stgDetail.RowCount := MemTable.RecordCount + 1; // +1 for header row
    stgDetail.ColCount := MemTable.FieldCount;

    // Populate header row
    for Col := 0 to MemTable.FieldCount - 1 do
    begin
      stgDetail.Cells[Col, 0] := MemTable.Fields[Col].DisplayName;
    end;

    // Populate data rows
    MemTable.First;
    Row := 1;
    while not MemTable.Eof do
    begin
      for Col := 0 to MemTable.FieldCount - 1 do
      begin
        stgDetail.Cells[Col, Row] := MemTable.Fields[Col].AsString;
      end;
      Inc(Row);
      MemTable.Next;
    end;
  end
  else
  begin
    ShowMessage('No data found.');
  end;
end;

end.

