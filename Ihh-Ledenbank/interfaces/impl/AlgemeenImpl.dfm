inherited frmAlgemeenImpl: TfrmAlgemeenImpl
  Caption = 'frmAlgemeenImpl'
  ClientHeight = 459
  ExplicitHeight = 497
  PixelsPerInch = 96
  TextHeight = 13
  inherited Panel3: TPanel
    Top = 380
  end
  inherited pnlLabels: TPanel
    Height = 380
    inherited Label11: TLabel
      Top = 216
      ExplicitTop = 216
    end
    inherited Label12: TLabel
      Top = 248
      ExplicitTop = 248
    end
    inherited Label13: TLabel
      Top = 280
      ExplicitTop = 280
    end
    object Label2: TLabel
      Left = 16
      Top = 184
      Width = 44
      Height = 13
      Caption = 'Kategori:'
    end
  end
  inherited pnlFields: TPanel
    Height = 380
    inherited edtTelefoon: TEdit
      Top = 208
      ExplicitTop = 208
    end
    inherited edtEmail: TEdit
      Top = 240
      ExplicitTop = 240
    end
    inherited edtOmschrijving: TMemo
      Top = 273
      ExplicitTop = 273
    end
    inherited edtAlgemeen: TEdit
      TabOrder = 9
    end
    object cmbCategorie: TComboBox
      Left = 6
      Top = 176
      Width = 145
      Height = 21
      Hint = 'Categorieen'
      HelpType = htKeyword
      HelpKeyword = 'Categorie'
      TabOrder = 8
      TextHint = 'Categorie'
    end
  end
end
