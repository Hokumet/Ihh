unit HEdit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, HolderEdits, AdoDB, Contnrs, Data.DB;

type
  TfrmHEdit = class(TForm)
    Panel3: TPanel;
    LblAangemaakDoor: TLabel;
    lblAangemaaktOp: TLabel;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnReset: TBitBtn;
    pnlLabels: TPanel;
    pnlFields: TPanel;
    CurrQuery: TADOQuery;
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnResetClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    sFieldName: String;
    sFieldValue: String;
    ID: Integer;
    CurrTable: TADOTable;
    MasterKey: String;
    TableObjectList: TObjectList;
    procedure fillCmb(cmb:TComboBox; def:String);
  protected
    function getFieldName(edtField: TControl):String;
    function getFieldValue(edtField: TControl):String;
    procedure loadField(eField: TControl);
    procedure saveField(eField:TControl);
    procedure loadEditField(edtField:TEdit);
    procedure loadComboField(comboField:TComboBox);
    procedure loadCurrencyField(currField:THCurrencyEdit);
    procedure loadDateField(dateField:TDateTimePicker);
    procedure loadMemoField(memoField:TMeMo);
    procedure loadDetails();
    procedure loadDetailsTables();
    procedure loadFields();
    procedure saveEditField(edtField:TEdit);
    procedure saveComboField(comboField:TComboBox);
    procedure saveCurrencyField(currField:THCurrencyEdit);
    procedure saveDateField(dateField:TDateTimePicker);
    procedure saveMemoField(memoField:TMeMo);
    procedure saveFields();
    procedure cancelFields();
    procedure loadOnce(); virtual;
  public
    constructor Create(Owner: TComponent; ID: Integer; AdoTable: TADOTable); overload;
    constructor Create(Owner: TComponent; ID: Integer; AdoTable: TADOTable; Key: String); overload;
  end;

var
  frmHEdit: TfrmHEdit;

implementation

uses Main;
{$R *.dfm}
{ TfrmEditAlgemeen }

constructor TfrmHEdit.Create(Owner: TComponent; ID: Integer;
  AdoTable: TADOTable);
begin
  inherited Create(Owner);
  TableObjectList := TObjectList.Create;
  CurrTable := AdoTable;
  CurrQuery.Connection := CurrTable.Connection;
  loadOnce();
  Self.Id := Id;
  if Id = 0 then begin
    CurrTable.Insert;
    loadFields();
  end
  else if Id = -1 then begin
  end
  else begin
    loadFields();
    loadDetailsTables();
    loadDetails();
    CurrTable.Edit;
  end;
  btnReset.Visible := not(Id = 0);
end;

procedure TfrmHEdit.btnCancelClick(Sender: TObject);
begin
  cancelFields();
end;

procedure TfrmHEdit.btnResetClick(Sender: TObject);
begin
  CurrTable.Cancel;
  loadFields();
  loadDetailsTables();
  loadDetails();
  CurrTable.edit;
end;

procedure TfrmHEdit.btnSaveClick(Sender: TObject);
begin
  if CurrTable.FieldByName('AangemaaktDoor').AsString = '' then begin
    CurrTable.FieldByName('AangemaaktDoor').AsString := TfrmMain(Owner).user;
    CurrTable.FieldByName('AangemaaktOp').AsDateTime := Date;
  end;
  saveFields();
end;

procedure TfrmHEdit.cancelFields;
var I: Integer;
begin
  for I := 0 to TableObjectList.Count -1 do begin
    try
      TADOTable(TableObjectList.Items[I]).CancelBatch;
    except
    on E: Exception do
       //
    end;
  end;
  CurrTable.Cancel;
end;

constructor TfrmHEdit.Create(Owner: TComponent; ID: Integer;
  AdoTable: TADOTable; Key: String);
begin
  inherited Create(Owner);
  TableObjectList := TObjectList.Create;
  CurrTable := AdoTable;
  CurrQuery.Connection := CurrTable.Connection;
  loadOnce();
  Self.Id := Id;
  MasterKey := Key;
  if Id = 0  then begin
    CurrTable.Insert;
    loadFields();
  end
  else begin
    loadFields();
    loadDetailsTables();
    loadDetails();
    CurrTable.Edit;
  end;
  btnReset.Visible := not(Id = 0);
end;

procedure TfrmHEdit.fillCmb(cmb: TComboBox; def: String);
var I: Integer;
    table:String;
    field:String;
begin
  table :=  cmb.Hint;
  field := cmb.TextHint;
  CurrQuery.SQL.Clear;
  CurrQuery.SQL.Add('Select *  from '+ table);
  CurrQuery.ExecSQL;

  CurrQuery.Open;
  CurrQuery.First;

  for I := 0 to CurrQuery.RecordCount -1 do begin
    cmb.AddItem(CurrQuery.FieldByName(field).AsString, Pointer(CurrQuery.FieldByName('Id').AsInteger));
    if(CurrQuery.FieldByName(field).AsString = def) then begin
      cmb.ItemIndex := I;
    end;
    CurrQuery.Next;
  end;
  if cmb.ItemIndex = -1 then
    cmb.ItemIndex := 0;
end;

procedure TfrmHEdit.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 then begin
    CurrTable.Cancel;
    Close
  end;
end;

function TfrmHEdit.getFieldName(edtField: TControl): String;
var
  edt: TEdit;
  curr: THCurrencyEdit;
  combo: TComboBox;
begin
  if edtField is TEdit then  begin
    edt := TEdit(edtField);
    Result := edt.HelpKeyword;
  end
  else if edtField is THCurrencyEdit then  begin
    curr := THCurrencyEdit(edtField);
    Result := curr.HelpKeyword;
  end
  else if edtField is TComboBox then  begin
    combo := TComboBox(edtField);
    Result := combo.HelpKeyword;
  end;
end;

function TfrmHEdit.getFieldValue(edtField: TControl): String;
var
  value: String;
  edt: TEdit;
  curr: THCurrencyEdit;
  combo: TComboBox;
begin
  if edtField is TEdit then  begin
    edt := TEdit(edtField);
    value := edt.HelpKeyword;
  end
  else if edtField is THCurrencyEdit then  begin
    curr := THCurrencyEdit(edtField);
    value := curr.HelpKeyword;
  end
  else if edtField is TComboBox then  begin
    combo := TComboBox(edtField);
    value := combo.HelpKeyword;
  end;
end;

procedure TfrmHEdit.loadComboField(comboField: TComboBox);
var
  value: String;
begin
  value := CurrTable.FieldByName(comboField.HelpKeyword).AsString;
  fillCmb(comboField,  value);
end;

procedure TfrmHEdit.loadCurrencyField(currField: THCurrencyEdit);
var
  value: Double;
begin
  value := CurrTable.FieldByName(currField.HelpKeyword).AsFloat;
  currField.Value := value;
end;

procedure TfrmHEdit.loadDateField(dateField: TDateTimePicker);
var
  value: TDateTime;
begin
  value := CurrTable.FieldByName(dateField.HelpKeyword).AsDateTime;
  dateField.Date := value;
end;

procedure TfrmHEdit.loadDetails;
var I: Integer;
begin
  for I := 0 to TableObjectList.Count -1 do begin
    TADOTable(TableObjectList.Items[I]).Filtered := False;
    TADOTable(TableObjectList.Items[I]).Filter := Masterkey+ '= '+ IntToStr(Id);
    TADOTable(TableObjectList.Items[I]).Filtered := True;
  end;
  // verder vullen bij kind
//  frameDagen.lvwItems.Clear;
end;

procedure TfrmHEdit.loadDetailsTables;
begin
//vullen bij kind
//  TDagen := TfrmMain.GetTableDagen;
//  TableObjectList.Add(TDagen);
//  frameDagen.FTable := TDagen;
end;

procedure TfrmHEdit.loadEditField(edtField: TEdit);
var
  value: String;
begin
  value := CurrTable.FieldByName(edtField.HelpKeyword).AsString;
  edtField.Text := value;
end;

procedure TfrmHEdit.loadField(eField: TControl);
begin
  if eField is TEdit then  begin
    loadEditField(TEdit(eField));
  end
  else if eField is THCurrencyEdit then  begin
    loadCurrencyField(THCurrencyEdit(eField));
  end
  else if eField is TComboBox then  begin
    loadComboField(TComboBox(eField));
  end
  else if eField is TDateTimePicker then  begin
    loadDateField(TDateTimePicker(eField));
  end
  else if eField is TMemo then  begin
    loadMemoField(TMemo(eField));
  end;
end;

procedure TfrmHEdit.loadFields;
var I: Integer;
begin
  for I := 0 to pnlFields.ControlCount - 1 do begin
    loadField(pnlFields.Controls[I]);
  end;
  LblAangemaakDoor.Caption := 'Aangemaakt door: ' + CurrTable.FieldByName('AangemaaktDoor').AsString ;
  lblAangemaaktOp.Caption := 'Aangemaakt op: ' + CurrTable.FieldByName('AangemaaktOp').AsString;
end;


procedure TfrmHEdit.loadMemoField(memoField: TMeMo);
var
  value: String;
begin
  value := CurrTable.FieldByName(memoField.HelpKeyword).AsString;
  memoField.Lines.Add(value);
end;

procedure TfrmHEdit.loadOnce;
begin

end;

procedure TfrmHEdit.saveComboField(comboField: TComboBox);
var
  field: String;
begin                     
  field := comboField.HelpKeyword;
  CurrTable.FieldByName(field).AsString := comboField.Text;
end;

procedure TfrmHEdit.saveCurrencyField(currField: THCurrencyEdit);
var
  field: String;
begin                     
  field := currField.HelpKeyword;
  CurrTable.FieldByName(field).AsCurrency := currField.Value;
end;

procedure TfrmHEdit.saveDateField(dateField: TDateTimePicker);
var
  field: String;
begin                     
  field := dateField.HelpKeyword;
  CurrTable.FieldByName(field).AsDateTime := dateField.DateTime;
end;

procedure TfrmHEdit.saveEditField(edtField: TEdit);
var
  field: String;
begin
  field := edtField.HelpKeyword;
  CurrTable.FieldByName(field).AsString := edtField.Text;
end;

procedure TfrmHEdit.saveField(eField: TControl);
begin
  if eField is TEdit then  begin
    saveEditField(TEdit(eField));
  end
  else if eField is THCurrencyEdit then  begin
    saveCurrencyField(THCurrencyEdit(eField));
  end
  else if eField is TComboBox then  begin
    saveComboField(TComboBox(eField));
  end
  else if eField is TDateTimePicker then  begin
    saveDateField(TDateTimePicker(eField));
  end
  else if eField is TMemo then  begin
    saveMemoField(TMemo(eField));
  end;
end;

procedure TfrmHEdit.saveFields;
var I: Integer;
begin
  for I := 0 to pnlFields.ControlCount - 1 do begin
    saveField(pnlFields.Controls[I]);
  end;

  if CurrTable.FieldByName('AangemaaktDoor').AsString = '' then begin
    CurrTable.FieldByName('AangemaaktDoor').AsString := TfrmMain(Owner).user;
    CurrTable.FieldByName('AangemaaktOp').AsDateTime := Date;
  end;

  for I := 0 to TableObjectList.Count -1 do begin
    TADOTable(TableObjectList.Items[I]).UpdateBatch;
  end;
  CurrTable.UpdateBatch;
end;


procedure TfrmHEdit.saveMemoField(memoField: TMeMo);
var
  field: String;
begin
  field := memoField.HelpKeyword;
  CurrTable.FieldByName(field).AsString := memoField.Lines.Text;
end;

end.
