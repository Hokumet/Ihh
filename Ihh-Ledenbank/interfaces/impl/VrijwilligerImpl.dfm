inherited frmVrijwilligerImpl: TfrmVrijwilligerImpl
  Caption = 'frmVrijwilligerImpl'
  ExplicitWidth = 415
  PixelsPerInch = 96
  TextHeight = 13
  inherited pnlLabels: TPanel
    inherited Label11: TLabel
      Top = 248
      ExplicitTop = 248
    end
    inherited Label12: TLabel
      Top = 280
      ExplicitTop = 280
    end
    inherited Label13: TLabel
      Top = 312
      ExplicitTop = 312
    end
    object Label3: TLabel
      Left = 16
      Top = 216
      Width = 36
      Height = 13
      Caption = 'Meslek:'
    end
  end
  inherited pnlFields: TPanel
    inherited edtTelefoon: TEdit
      Top = 240
      ExplicitTop = 240
    end
    inherited edtEmail: TEdit
      Top = 272
      ExplicitTop = 272
    end
    inherited edtOmschrijving: TMemo
      Top = 305
      TabOrder = 10
      ExplicitTop = 305
    end
    object cmbBaan: TComboBox
      Left = 6
      Top = 208
      Width = 219
      Height = 21
      Hint = 'Banen'
      HelpType = htKeyword
      HelpKeyword = 'Baan'
      TabOrder = 9
      TextHint = 'Naam'
    end
    object edtVrijwilliger: TEdit
      Left = 158
      Top = 144
      Width = 43
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'Vrijwilliger'
      TabOrder = 11
      Text = 'true'
      Visible = False
    end
  end
end
