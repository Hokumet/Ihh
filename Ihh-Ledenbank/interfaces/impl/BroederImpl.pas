unit BroederImpl;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, IBroeder, Data.DB, Data.Win.ADODB,
  Vcl.StdCtrls, HolderEdits, Vcl.ComCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TfrmBroederImpl = class(TfrmIBroeder)
    lblBroeder: TLabel;
    edtYFONaam: TEdit;
    edtWees: TEdit;
    edtArm: TEdit;
  private
    { Private declarations }

  public
    constructor Create(Owner: TComponent; ID: Integer; AdoTable: TADOTable; Extra: String); overload;
  end;

var
  frmBroederImpl: TfrmBroederImpl;

implementation

{$R *.dfm}

{ TfrmBroederImpl }

constructor TfrmBroederImpl.Create(Owner: TComponent; ID: Integer;
  AdoTable: TADOTable; Extra: String);
begin
    inherited Create(Owner, ID, AdoTable);
    if(Extra = 'Wees') then begin
      edtArm.Text := 'false';
      lblBroeder.Caption := 'Yetimin Adi:';
    end
    else begin
      edtWees.Text := 'false';
      lblBroeder.Caption := 'Fakirin Adi:';
    end;
end;

end.
