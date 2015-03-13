unit EditProject;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, HEdit, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls, Data.DB, Data.Win.ADODB;

type
  TfrmEditProject = class(TfrmHEdit)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEditProject: TfrmEditProject;

implementation

{$R *.dfm}

end.
