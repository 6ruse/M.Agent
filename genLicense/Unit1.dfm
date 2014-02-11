object Form1: TForm1
  Left = 537
  Top = 311
  AutoSize = True
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'GenLicense M.Agent'
  ClientHeight = 140
  ClientWidth = 425
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 0
    Top = 3
    Width = 70
    Height = 13
    Caption = #1050#1083#1102#1095' '#1082#1083#1080#1077#1085#1090#1072
  end
  object Label2: TLabel
    Left = 16
    Top = 43
    Width = 57
    Height = 13
    Caption = #1040#1082#1090#1080#1074#1077#1085' '#1076#1086
  end
  object EdtKeyKlient: TEdit
    Left = 80
    Top = 0
    Width = 345
    Height = 21
    TabOrder = 0
  end
  object DWork: TDateTimePicker
    Left = 80
    Top = 39
    Width = 186
    Height = 21
    Date = 41633.931446296290000000
    Time = 41633.931446296290000000
    TabOrder = 1
  end
  object EdtLicense: TEdit
    Left = 0
    Top = 83
    Width = 425
    Height = 21
    TabOrder = 2
  end
  object BtnGenerateLicense: TButton
    Left = 272
    Top = 35
    Width = 153
    Height = 25
    Caption = 'Generate license'
    TabOrder = 3
    OnClick = BtnGenerateLicenseClick
  end
  object BtnClose: TButton
    Left = 0
    Top = 115
    Width = 425
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100' '#1087#1088#1086#1075#1088#1072#1084#1084#1091
    TabOrder = 4
    OnClick = BtnCloseClick
  end
end
