unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, MainAncestor, Vcl.Menus, Data.Win.ADODB,
  Data.DB, Vcl.ImgList, Vcl.ComCtrls, MyListView, Vcl.ToolWin, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.ShellAnimations, HolderEdits, HEdit,Vcl.Buttons;

// Replace Flags
type
  TWordReplaceFlags = set of (wrfReplaceAll, wrfMatchCase, wrfMatchWildcards);

type
  TfrmMain = class(TfrmMainAncestor)
    btnVrijwilliger: TToolButton;
    ATable: TADOTable;
    btnDonatie: TToolButton;
    btnWeesBroeder: TToolButton;
    btnArmBroeder: TToolButton;
    btnSpaarpot: TToolButton;
    btnBeurs: TToolButton;
    lblPlusMin: TLabel;
    zkLidnr: TEdit;
    zkOmschrijving: TEdit;
    zkNaam: TEdit;
    zkStraat: TEdit;
    zkPostcode: TEdit;
    zkPlaats: TEdit;
    zkTelefoon: TEdit;
    zkEmail: TEdit;
    DBTLanden: TADOTable;
    DBTBetalingen: TADOTable;
    DBTBanen: TADOTable;
    DBTCategorien: TADOTable;
    DBTLandenId: TAutoIncField;
    DBTLandenNaam: TWideStringField;
    DBTBetalingenId: TAutoIncField;
    DBTBetalingenType: TWideStringField;
    DBTBanenId: TAutoIncField;
    DBTBanenNaam: TWideStringField;
    DBTCategorienId: TAutoIncField;
    DBTCategorienCategorie: TWideStringField;
    DBTDoelen: TADOTable;
    DBTProjecten: TADOTable;
    DBTProjectenId: TAutoIncField;
    DBTProjectenNaam: TWideStringField;
    btnZoeken: TBitBtn;
    zkLidDatum: TEdit;
    edtLidMaand: TEdit;
    zkLidJaar: TEdit;
    zkBetaling: TComboBox;
    zkBedrag: THCurrencyEdit;
    zkBanknr: TEdit;
    zkDoel: TComboBox;
    zkLand: TComboBox;
    DBTDoelenId: TAutoIncField;
    DBTDoelenNaam: TWideStringField;
    zkYFONaam: TEdit;
    zkSpaarpotNr: TEdit;
    Landen1: TMenuItem;
    Doelen1: TMenuItem;
    BetalingTypes1: TMenuItem;
    Lists1: TMenuItem;
    Categorieen1: TMenuItem;
    Banen1: TMenuItem;
    Keywords1: TMenuItem;
    ATableId: TAutoIncField;
    ATableLidnr: TWideStringField;
    ATableNaam: TWideStringField;
    ATableInschrijfdatum: TDateTimeField;
    ATableLidDag: TWideStringField;
    ATableLidMaand: TWideStringField;
    ATableLidJaar: TWideStringField;
    ATableStraat: TWideStringField;
    ATablePostcode: TWideStringField;
    ATablePlaats: TWideStringField;
    ATableBanknr: TWideStringField;
    ATableBedrag: TBCDField;
    ATableBaan: TWideStringField;
    ATableTelefoonnr: TWideStringField;
    ATableEmail: TWideStringField;
    ATableDoel: TWideStringField;
    ATableLand: TWideStringField;
    ATableYFONaam: TWideStringField;
    ATableBetaling: TWideStringField;
    ATableCategorie: TWideStringField;
    ATableProject: TWideStringField;
    ATableOmschrijving: TWideStringField;
    ATableSpaarpotNr: TWideStringField;
    ATableAlgemeen: TWideStringField;
    ATableVrijwilliger: TWideStringField;
    ATableSpaarpot: TWideStringField;
    ATableDonatie: TWideStringField;
    ATableArm: TWideStringField;
    ATableWees: TWideStringField;
    ATableAangemaaktDoor: TWideStringField;
    ATableAangemaaktOp: TDateTimeField;
    procedure btnPrintenClick(Sender: TObject);
    procedure btnBeginClick(Sender: TObject);
    procedure Instellingen1Click(Sender: TObject);
    procedure btnDonatieClick(Sender: TObject);
    procedure btnVrijwilligerClick(Sender: TObject);
    procedure btnSpaarpotClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure lblPlusMinClick(Sender: TObject);
    procedure btnWeesBroederClick(Sender: TObject);
    procedure btnArmBroederClick(Sender: TObject);
    procedure Keywords1Click(Sender: TObject);
    procedure btnZoekenClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure Landen1Click(Sender: TObject);
    procedure BetalingTypes1Click(Sender: TObject);
  private
    procedure FindUpdateField(WordApp: OLEVariant; SearchString: String;
      ReplaceString: String; Flags: TWordReplaceFlags);
    procedure LoadAlgemeen(refreshAll: boolean);
    procedure LoadDonatie(refreshAll: boolean);
    procedure LoadWees(refreshAll: boolean);
    procedure LoadArm(refreshAll: boolean);
    procedure LoadVrijwilliger(refreshAll: boolean);
    procedure LoadSpaarpot(refreshAll: boolean);
    procedure LoadBeurs(refreshAll: boolean);

    procedure ReloadData(filter: boolean);
    procedure ReloadExtra();

    procedure SetAlgemeenColums();
    procedure SetDonatieColums();
    procedure SetWeesColums();
    procedure SetArmColums();
    procedure SetVrijwilligerColums();
    procedure SetSpaarpotColums();
    procedure SetBeursColums();

    procedure ShowEditForm(frmHEdit: TfrmHEdit);
  protected
    procedure Refresh; override;
    procedure RefreshData(refreshAll: boolean);
    procedure LoadFiltered(SQl: String); override;
    function GetFilterFromPanel(panel: TPanel):String;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ComObj, SettingImp, AlgemeenImpl, VrijwilligerImpl, SpaarpotImpl,
      DonatieImpl, BroederImpl, EditLand, EditBetaling;

Const
  wdSendToNewDocument = 0;
  Algemeen = 'Algemeen';
  Donatie = 'Donatie';
  Wees = 'Wees';
  Arm = 'Arm';
  Vrijwilliger = 'Vrijwilliger';
  Spaarpot = 'Spaarpot';
  Beurs = 'Beurs';
{$R *.dfm}

procedure TfrmMain.BetalingTypes1Click(Sender: TObject);
begin
  objectType := 'Betaling';
  CurrentTable := DBTBetalingen;
  lvwItems.Columns.Clear;
  FieldCaptionAndFieldName.Clear;

  lvwItems.Columns.BeginUpdate;
  addColumn('Aidat takip','Type', 100);
  lvwItems.Columns.EndUpdate;


  ReloadExtra();

end;

procedure TfrmMain.btnArmBroederClick(Sender: TObject);
begin
  objectType := btnArmBroeder.HelpKeyword;
  RefreshData(true);
//  objectType := btnArmBroeder.HelpKeyword;
//  frmEditBroederImp := TfrmEditBroederImp.Create(Self, 1, ATable);
//  try
//      if frmEditBroederImp.ShowModal = MrOk then
//      // Refresh;
//   finally
//      frmEditBroederImp.Free;
//   end;
end;

procedure TfrmMain.btnBeginClick(Sender: TObject);
begin
    objectType := btnBegin.HelpKeyword;
    RefreshData(true);
    //inherited;

//    frmISpaarpot := TfrmISpaarpot.Create(Self, 1, ATable);
//    try
//      if frmISpaarpot.ShowModal = MrOk then
//        // Refresh;
//     finally
//        frmISpaarpot.Free;
//     end;
//    frmISpaarpot := TfrmISpaarpot.Create(Self, 1, ATable);
//    try
//      if frmISpaarpot.ShowModal = MrOk then
//        // Refresh;
//     finally
//        frmISpaarpot.Free;
//     end;
//
//
//
//
//    frmEditAlgemeenImp := TfrmEditAlgemeenImp.Create(Self, 1, ATable);
//    try
//      if frmEditAlgemeenImp.ShowModal = MrOk then
//        // Refresh;
//     finally
//        frmEditAlgemeenImp.Free;
//     end;

end;

procedure TfrmMain.btnDonatieClick(Sender: TObject);
begin
  objectType := btnDonatie.HelpKeyword;
  RefreshData(true);

//  inherited;
//  frmEditDonatieImp := TfrmEditDonatieImp.Create(Self, 1, ATable);
//  try
//    if frmEditDonatieImp.ShowModal = MrOk then
//      // Refresh;
//    finally
//      frmEditDonatieImp.Free;
//    end;
end;

procedure TfrmMain.btnEditClick(Sender: TObject);
begin
  if objectType = Algemeen then
    frmHEdit := TfrmAlgemeenImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable)
  else if objectType = Vrijwilliger then
    frmHEdit := TfrmVrijwilligerImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable)
  else if objectType = Spaarpot then
    frmHEdit := TfrmSpaarpotImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable)
  else if objectType = Donatie then
    frmHEdit := TfrmDonatieImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable)
  else if objectType = Arm then
    frmHEdit := TfrmBroederImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable, objectType)
  else if objectType = Wees then
    frmHEdit := TfrmBroederImpl.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable, objectType)
  else if objectType = 'Land' then
    frmHEdit := TfrmEditLand.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable, objectType)
  else if objectType = 'Betaling' then
    frmHEdit := TfrmEditBetaling.Create(Self, Integer(lvwItems.Selected.Data), CurrentTable, objectType);

  ShowEditForm(frmHEdit);
end;

procedure TfrmMain.btnPrintenClick(Sender: TObject);
var
  WordApp, WordDoc: OLEVariant;
  template, rFieldName, rFieldValue:String;
  I, Y: Integer;
  printSelected: Boolean;
begin
    printSelected := false;

    if lvwItems.SelCount = 0 then begin
      if MessageDlg(IntToStr(lvwItems.Items.Count) + ' kisinin bilgisi print yapilacak.' + #13#10 +
        'Hepsini print yapmak istediginden emin misin?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        exit;
    end
    else if lvwItems.SelCount = 1 then
      printSelected := true
    else if MessageDlg(IntToStr(lvwItems.SelCount) + ' kisinin bilgisi print yapilacak.' + #13#10 +
      'Hepsini print yapmak istediginden emin misin?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      exit
    else
      printSelected := true;

    template :=  Inifile.ReadString('app','worddoc', '');
    Screen.Cursor := crHourglass;
    try
      WordApp := CreateOleObject('Word.Application');
    except
      ShowMessage('Kan MS Word niet starten.');
      Screen.Cursor := crDefault;
      exit;
    end;

    try
      Screen.Cursor := crHourGlass;
      //WordApp.Visible := true;

      with lvwItems do
        for I := 0 to Items.Count - 1 do begin
          // Open a Word Document
          WordDoc := WordApp.Documents.Add(template);
          CurrentTable.Locate('ID', Integer(Items[I].Data), []);
          if printSelected = true then begin
            if Items[I].Selected then begin
              //ATable.Locate('ID', Integer(Items[I].Data), []);
              for Y := 0 to CurrentTable.Fields.Count - 1 do  begin
                rFieldName := CurrentTable.Fields.Fields[I].FieldName;
                rFieldValue := CurrentTable.FieldByName(rFieldName).AsString;
                FindUpdateField(WordApp, '__' + rFieldName + '__', rFieldValue, [wrfReplaceAll]);
              end;
              WordDoc.PrintPreview;
              // WordDoc.PrintOut;
              ShowMessage('Previewing');
            end;
          end
          else  begin
            for Y := 0 to CurrentTable.Fields.Count - 1 do  begin
               rFieldName := CurrentTable.Fields.Fields[I].FieldName;
               rFieldValue := CurrentTable.FieldByName(rFieldName).AsString;
               FindUpdateField(WordApp, '__' + rFieldName + '__', rFieldValue, [wrfReplaceAll]);
            end;
            WordDoc.PrintPreview;
          end;
          WordApp.ActiveDocument.Close(0);
        end;

    finally
      // Close the original form document
      WordApp.Quit;
      WordApp := Unassigned;
      Screen.Cursor := crDefault;
    end;

    // for X := 0 to lvwItems.Items.Count - 1 do
    // begin
    // FName := lvwItems.Items[X].Caption;
    // FCaption := lvwItems.Items[X].SubItems[0];
    // NField :=  ATable.FieldByName(FName);
    //
    // if NField.DataType = ftString then
    // ShowMessage(NField.FieldName + ' = ' + FCaption +  ' = String' )
    // else if NField.DataType = ftInteger then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = Integer')
    // else if NField.DataType = ftDate then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = Date')
    // else if NField.DataType = ftDateTime then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = DateTime')
    // else if NField.DataType = ftCurrency then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = Currentcy')
    // else if NField.DataType = ftWideString then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftWideString')
    // else if NField.DataType = ftSmallint then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftSmallint')
    // else if NField.DataType = ftWord then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftWord')
    // else if NField.DataType = ftBoolean then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftBoolean')
    // else if NField.DataType = ftWord then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftWord')
    // else if NField.DataType = ftFloat then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftFloat')
    // else if NField.DataType = ftAutoInc then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftAutoInc')
    // else if NField.DataType = ftBytes then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftBytes')
    // else if NField.DataType = ftBCD then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftBCD')
    // else if NField.DataType = ftVarBytes then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftVarBytes')
    // else if NField.DataType = ftFixedChar then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftFixedChar')
    // else if NField.DataType = ftExtended then
    // ShowMessage(NField.FieldName + ' = ' + FCaption + ' = ftExtended')
    // else
    // ShowMessage(NField.FieldName)
    // end;

end;

procedure TfrmMain.btnSpaarpotClick(Sender: TObject);
begin
  objectType := btnSpaarpot.HelpKeyword;
  RefreshData(true);
//  objectType := btnSpaarpot.HelpKeyword;
//
//  frmEditSpaarpotImp := TfrmEditSpaarpotImp.Create(Self, 1, ATable);
//  try
//  if frmEditSpaarpotImp.ShowModal = MrOk then
//    // Refresh;
//  finally
//    frmEditSpaarpotImp.Free;
//  end;
end;

procedure TfrmMain.btnVrijwilligerClick(Sender: TObject);
begin
  objectType := btnVrijwilliger.HelpKeyword;
  RefreshData(true);

//  objectType := btnVrijwilliger.HelpKeyword;
//  inherited;
//  frmEditVrijwilligerImp := TfrmEditVrijwilligerImp.Create(Self,
//    1, ATable);
//  try
//    if frmEditVrijwilligerImp.ShowModal = MrOk then
//      // Refresh;
//    finally
//      frmEditVrijwilligerImp.Free;
//    end;
end;

procedure TfrmMain.btnWeesBroederClick(Sender: TObject);
begin
  objectType := btnWeesBroeder.HelpKeyword;
  RefreshData(true);
//  objectType := btnWeesBroeder.HelpKeyword;
//  frmEditBroederImp := TfrmEditBroederImp.Create(Self, 1, ATable);
//  try
//    if frmEditBroederImp.ShowModal = MrOk then
//      // Refresh;
//    finally
//      frmEditBroederImp.Free;
//    end;
end;

procedure TfrmMain.FindUpdateField(WordApp: OLEVariant;
  SearchString: String; ReplaceString: String;
  Flags: TWordReplaceFlags);
const
  wdFindContinue = 1;
  wdReplaceOne = 1;
  wdReplaceAll = 2;
  wdDoNotSaveChanges = 0;
begin
  try
    { Initialize parameters }
    WordApp.Selection.Find.ClearFormatting;
    WordApp.Selection.Find.Text := SearchString;
    WordApp.Selection.Find.Replacement.Text := ReplaceString;
    WordApp.Selection.Find.Forward := true;
    WordApp.Selection.Find.Wrap := wdFindContinue;
    WordApp.Selection.Find.Format := false;
    WordApp.Selection.Find.MatchCase := wrfMatchCase in Flags;
    WordApp.Selection.Find.MatchWholeWord := false;
    WordApp.Selection.Find.MatchWildcards :=
      wrfMatchWildcards in Flags;
    WordApp.Selection.Find.MatchSoundsLike := false;
    WordApp.Selection.Find.MatchAllWordForms := false;
    { Perform the search }
    if wrfReplaceAll in Flags then
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll)
    else
      WordApp.Selection.Find.Execute(Replace := wdReplaceOne);

    // WordApp.ActiveDocument.PrintPreview;
    { Save word }
    // WordApp.ActiveDocument.SaveAs(ADocument);
    { Assume that successful }
    // Result := True;
    { Close the document }
  finally
    { Quit Word }
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  inherited;
  ATable.Open;
  DBTLanden.Open;
  DBTBetalingen.Open;
  CurrentTable := ATable;
end;

function TfrmMain.GetFilterFromPanel(panel: TPanel): String;
var
  I: Integer;
  edt: TEdit;
  curr: THCurrencyEdit;
  combo: TComboBox;
  FieldName, FilterString: String;
  pnl:TPanel;
begin
  if panel = nil then
    pnl := pnlHeader
  else
    pnl := panel;

  for I := 0 to pnl.ControlCount - 1 do
  begin
    if pnl.Controls[I].Visible then
    begin
      if pnl.Controls[I] is TEdit then
      begin
        edt := TEdit(pnl.Controls[I]);
        FieldName := edt.HelpKeyword;
        FilterString := FilterString +
          GetEdtFieldFilter(FieldName, edt);
      end
      else if pnl.Controls[I] is THCurrencyEdit then
      begin
        curr := THCurrencyEdit(pnl.Controls[I]);
        FieldName := curr.HelpKeyword;
        FilterString := FilterString + GetCurrencyFieldFilter
          (FieldName, curr);
      end
      else if pnl.Controls[I] is TComboBox then
      begin
        combo := TComboBox(pnl.Controls[I]);
        FieldName := combo.HelpKeyword;
        FilterString := FilterString + GetComboFieldFilter
          (FieldName, combo);
      end;
    end;
  end;


end;

procedure TfrmMain.Instellingen1Click(Sender: TObject);
begin
  frmSettingImp := TfrmSettingImp.Create(Self);
  try
    if frmSettingImp.ShowModal = MrOk then
      RefreshBackground;
  finally
    frmSettingImp.Free;
  end
end;

procedure TfrmMain.Keywords1Click(Sender: TObject);
begin
    MessageDlg('' + #13 + 'Word programinda kullanabileceginiz anahtar sözcükler:' +
    #13 + #13 + 'Bu sözcükler yerine asil veriler geçecektir ' + #13
    + 'Genel Adres: Lidnr, Naam, Straat, Postcode, Plaats, Telefoonnr, Email, Omschrijving '+ #13
    + 'Aylik Bagis: Lidnr, Naam, Straat, Postcode, Plaats, Telefoonnr, Email, Omschrijving '+ #13+ #13
    + 'Yukaridaki sözcükler Word dosyanizda 4 tane __ __ arasina aliniz. '+ #13
    + 'Örnek ==>   __Naam__ veya __LidNr__ veya __Email__  ',
    mtCustom, [mbOK], 0 );
end;

procedure TfrmMain.Landen1Click(Sender: TObject);
begin
  objectType := 'Land';
  CurrentTable := DBTLanden;
  lvwItems.Columns.Clear;
  FieldCaptionAndFieldName.Clear;

  lvwItems.Columns.BeginUpdate;
  addColumn('Ülke ismi','Naam', 100);
  lvwItems.Columns.EndUpdate;

  ReloadExtra();
end;

procedure TfrmMain.lblPlusMinClick(Sender: TObject);
begin
  if lblPlusMin.Caption = '+' then begin
    pnlHeader.Height := 51;
    lblPlusMin.Caption := '-';
  end
  else  begin
    pnlHeader.Height := 28;
    lblPlusMin.Caption := '+';
  end;
end;

procedure TfrmMain.LoadAlgemeen(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;
    FieldCaptionAndFieldName.Clear;

    lvwItems.Columns.BeginUpdate;
    SetAlgemeenColums;
    lvwItems.Columns.EndUpdate;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadArm(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;

    lvwItems.Columns.BeginUpdate;
    SetArmColums;
    lvwItems.Columns.EndUpdate;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadDonatie(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;

    lvwItems.Columns.BeginUpdate;
    SetDonatieColums;
    lvwItems.Columns.EndUpdate;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadFiltered(SQl: String);
begin
  CurrentTable.Filtered := false;
  CurrentTable.Filter := SQl;
  CurrentTable.Filtered := true;
end;

procedure TfrmMain.LoadBeurs(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;
    SetBeursColums;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadSpaarpot(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;

    lvwItems.Columns.BeginUpdate;
    SetSpaarpotColums;
    lvwItems.Columns.EndUpdate;

  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadVrijwilliger(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;

    lvwItems.Columns.BeginUpdate;
    SetVrijwilligerColums;
    lvwItems.Columns.EndUpdate;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.LoadWees(refreshAll: boolean);
begin
  if refreshAll then begin
    lvwItems.Columns.Clear;

    lvwItems.Columns.BeginUpdate;
    SetWeesColums;
    lvwItems.Columns.EndUpdate;
  end;

  ReloadData(not(refreshAll));
end;

procedure TfrmMain.Refresh;
begin
  RefreshData(false);
end;

procedure TfrmMain.RefreshData(refreshAll: boolean);
begin
  CurrentTable := ATable;
  if objectType = Algemeen  then
    LoadAlgemeen(refreshAll)
  else if objectType = Donatie  then
    LoadDonatie(refreshAll)
  else if objectType = Wees  then
    LoadWees(refreshAll)
  else if objectType = Arm  then
    LoadArm(refreshAll)
  else if objectType = Vrijwilliger  then
    LoadVrijwilliger(refreshAll)
  else if objectType = Spaarpot  then
    LoadSpaarpot(refreshAll)
  else if objectType = Beurs  then
    LoadBeurs(refreshAll)
end;

procedure TfrmMain.ReloadData(filter: boolean);
var sFilter:String;
  LI: TListItem;
  I,X: Integer;
  column: TListColumn;
begin
  if filter then
    sFilter := GetFilterFromPanel(nil);
  LoadFiltered(objectType + '=' + QuotedStr('true') + sFilter);

  lvwItems.Clear;
  CurrentTable.Last;
  for I := 0 to CurrentTable.RecordCount - 1 do
  begin
    LI := lvwItems.Items.Add;
    for X := 0 to lvwItems.Columns.Count -1 do begin
      column := lvwItems.Columns.Items[X];
      if X = 0 then
        LI.Caption := CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsString
      else begin
        if FieldCaptionAndFieldType.Values[column.DisplayName] = 'curr' then
          LI.SubItems.Add('€ '+FormatFloat('0.00',CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsFloat))
        else
          LI.SubItems.Add(CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsString);
      end;
    end;
    LI.Data := Pointer(CurrentTable.FieldByName('ID').asInteger);
    CurrentTable.Prior;
  end;
end;

procedure TfrmMain.ReloadExtra;
var sFilter:String;
  LI: TListItem;
  I,X: Integer;
  column: TListColumn;
begin
  lvwItems.Clear;
  CurrentTable.Filtered := False;
  CurrentTable.Last;
  for I := 0 to CurrentTable.RecordCount - 1 do
  begin
    LI := lvwItems.Items.Add;
    for X := 0 to lvwItems.Columns.Count -1 do begin
      column := lvwItems.Columns.Items[X];
      if X = 0 then
        LI.Caption := CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsString
      else begin
        if FieldCaptionAndFieldType.Values[column.DisplayName] = 'curr' then
          LI.SubItems.Add('€ '+FormatFloat('0.00',CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsFloat))
        else
          LI.SubItems.Add(CurrentTable.FieldByName(FieldCaptionAndFieldName.Values[column.DisplayName]).AsString);
      end;
    end;
    LI.Data := Pointer(CurrentTable.FieldByName('ID').asInteger);
    CurrentTable.Prior;
  end;

end;

procedure TfrmMain.SetAlgemeenColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Kategori', 'Categorie', 75);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.SetArmColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Gün','LidDag', 34);
  addColumn('Ay','LidMaand', 27);
  addColumn('Yil','LidJaar', 40);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Banka No', 'Banknr', 100);
  addColumn('Miktar', 'Bedrag', 'curr', 100);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Ülke', 'Land', 100);
  addColumn('Fakirin Adi', 'YFONaam', 100);
  addColumn('Aidat Takip', 'Betaling', 75);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.SetDonatieColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Gün','LidDag', 34);
  addColumn('Ay','LidMaand', 27);
  addColumn('Yil','LidJaar', 40);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Banka No', 'Banknr', 100);
  addColumn('Miktar', 'Bedrag', 'curr', 100);
  addColumn('Amaç', 'Doel', 100);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Aidat Takip', 'Betaling', 75);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.SetBeursColums;
begin

end;

procedure TfrmMain.SetSpaarpotColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Gün','LidDag', 34);
  addColumn('Ay','LidMaand', 27);
  addColumn('Yil','LidJaar', 40);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Kumbara No', 'SpaarpotNr', 75);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Ülke', 'Land', 75);
  addColumn('Ögrenci ', 'YFONaam', 75);
  addColumn('Aidat Takip', 'Betaling', 75);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.SetVrijwilligerColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Gün','LidDag', 34);
  addColumn('Ay','LidMaand', 27);
  addColumn('Yil','LidJaar', 40);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Meslek', 'Baan', 75);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.SetWeesColums;
begin
  addColumn('Sira No','LidNr', 50);
  addColumn('Gün','LidDag', 34);
  addColumn('Ay','LidMaand', 27);
  addColumn('Yil','LidJaar', 40);
  addColumn('Adi Soyadi', 'Naam', 100);
  addColumn('Adres', 'Straat', 100);
  addColumn('Posta kodu', 'Postcode', 75);
  addColumn('Sehir', 'Plaats', 75);
  addColumn('Banka No', 'Banknr', 100);
  addColumn('Miktar', 'Bedrag', 'curr', 100);
  addColumn('Telefon No', 'Telefoonnr', 75);
  addColumn('Email', 'Email', 100);
  addColumn('Ülke', 'Land', 100);
  addColumn('Yetimin Adi', 'YFONaam', 100);
  addColumn('Aidat Takip', 'Betaling', 75);
  addColumn('Açiklama', 'Omschrijving', 100);
end;

procedure TfrmMain.ShowEditForm(frmHEdit: TfrmHEdit);
begin
  try
    if frmHEdit.ShowModal = mrOk then
      Refresh;
  finally
    frmHEdit.Free;
  end;
end;

procedure TfrmMain.btnZoekenClick(Sender: TObject);
begin
  RefreshData(False);
//  GetFilterFromPanel(pnlHeader);
end;

end.
