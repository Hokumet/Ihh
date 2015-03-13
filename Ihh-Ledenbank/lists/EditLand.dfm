inherited frmEditLand: TfrmEditLand
  Caption = 'frmEditLand'
  ClientHeight = 143
  ExplicitHeight = 181
  PixelsPerInch = 96
  TextHeight = 13
  inherited Panel3: TPanel
    Top = 64
    ExplicitTop = 124
  end
  inherited pnlLabels: TPanel
    Height = 64
    ExplicitHeight = 124
    object Label1: TLabel
      Left = 16
      Top = 24
      Width = 24
      Height = 13
      Caption = #220'lke:'
    end
  end
  inherited pnlFields: TPanel
    Height = 64
    ExplicitHeight = 124
    object edtNaam: TEdit
      Left = 6
      Top = 16
      Width = 121
      Height = 21
      HelpType = htKeyword
      HelpKeyword = 'Naam'
      TabOrder = 0
    end
  end
  inherited CurrQuery: TADOQuery
    Top = 48
  end
end
