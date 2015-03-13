inherited frmSpaarpotImpl: TfrmSpaarpotImpl
  Caption = 'frmSpaarpotImpl'
  ExplicitWidth = 415
  PixelsPerInch = 96
  TextHeight = 13
  inherited pnlLabels: TPanel
    inherited Label3: TLabel
      Top = 372
      ExplicitTop = 372
    end
    inherited Label11: TLabel
      Top = 248
      ExplicitTop = 248
    end
    inherited Label12: TLabel
      Top = 280
      ExplicitTop = 280
    end
    inherited Label13: TLabel
      Top = 402
      ExplicitTop = 402
    end
    object Label4: TLabel
      Left = 16
      Top = 216
      Width = 62
      Height = 13
      Caption = 'Kumbara No:'
    end
    object Label5: TLabel
      Left = 16
      Top = 312
      Width = 24
      Height = 13
      Caption = #220'lke:'
    end
    object Label6: TLabel
      Left = 16
      Top = 342
      Width = 58
      Height = 13
      Caption = #214'grenci adi:'
    end
  end
  inherited pnlFields: TPanel
    inherited cmbBetaling: TComboBox
      Top = 364
      TabOrder = 13
      ExplicitTop = 364
    end
    inherited edtTelefoon: TEdit
      Top = 240
      ExplicitTop = 240
    end
    inherited edtEmail: TEdit
      Top = 272
      ExplicitTop = 272
    end
    inherited edtOmschrijving: TMemo
      Top = 396
      ExplicitTop = 396
    end
    object edtSpaarpotNr: TEdit
      Left = 6
      Top = 208
      Width = 121
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'SpaarpotNr'
      TabOrder = 10
    end
    object cmbLand: TComboBox
      Left = 6
      Top = 304
      Width = 145
      Height = 21
      Hint = 'Landen'
      HelpType = htKeyword
      HelpKeyword = 'Land'
      TabOrder = 11
      TextHint = 'Naam'
    end
    object edtYFONaam: TEdit
      Left = 6
      Top = 334
      Width = 219
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'YFONaam'
      TabOrder = 12
    end
    object edtSpaarpot: TEdit
      Left = 158
      Top = 144
      Width = 43
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'Spaarpot'
      TabOrder = 14
      Text = 'true'
      Visible = False
    end
  end
end
