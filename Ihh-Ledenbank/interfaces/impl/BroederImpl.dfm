inherited frmBroederImpl: TfrmBroederImpl
  Caption = 'frmBroederImpl'
  ExplicitTop = 8
  PixelsPerInch = 96
  TextHeight = 13
  inherited pnlLabels: TPanel
    inherited Label3: TLabel
      Top = 404
      ExplicitTop = 404
    end
    inherited Label13: TLabel
      Top = 434
      ExplicitTop = 434
    end
    object lblBroeder: TLabel
      Left = 16
      Top = 374
      Width = 56
      Height = 13
      Caption = 'Yetimin Adi:'
    end
  end
  inherited pnlFields: TPanel
    inherited cmbBetaling: TComboBox
      Top = 396
      ExplicitTop = 396
    end
    inherited edtTelefoon: TEdit
      TabOrder = 14
    end
    inherited edtOmschrijving: TMemo
      Top = 431
      ExplicitTop = 431
    end
    object edtYFONaam: TEdit
      Left = 6
      Top = 366
      Width = 219
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'YFONaam'
      TabOrder = 13
    end
    object edtWees: TEdit
      Left = 158
      Top = 144
      Width = 43
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'Wees'
      TabOrder = 15
      Text = 'true'
      Visible = False
    end
    object edtArm: TEdit
      Left = 158
      Top = 176
      Width = 43
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'Arm'
      TabOrder = 16
      Text = 'true'
      Visible = False
    end
  end
end
