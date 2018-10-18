object FormMain: TFormMain
  Left = 0
  Top = 0
  Caption = #1054#1073#1088#1072#1073#1086#1090#1072#1090#1100' '#1087#1088#1072#1081#1089#1099' '#1076#1083#1103' Prom.ua'
  ClientHeight = 563
  ClientWidth = 980
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    980
    563)
  PixelsPerInch = 96
  TextHeight = 13
  object MemoTxt: TMemo
    Left = 8
    Top = 8
    Width = 964
    Height = 97
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Courier New'
    Font.Style = []
    Lines.Strings = (
      'MemoTxt')
    ParentFont = False
    ParentShowHint = False
    ReadOnly = True
    ScrollBars = ssBoth
    ShowHint = True
    TabOrder = 0
    Visible = False
    WantTabs = True
    WordWrap = False
  end
  object BitBtnXLS: TBitBtn
    Left = 8
    Top = 506
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'XLS file'
    Kind = bkOK
    NumGlyphs = 2
    TabOrder = 1
    OnClick = BitBtnXLSClick
  end
  object BitBtnClose: TBitBtn
    Left = 897
    Top = 530
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Kind = bkClose
    NumGlyphs = 2
    TabOrder = 2
    OnClick = BitBtnCloseClick
  end
  object MemoLog: TMemo
    Left = 8
    Top = 120
    Width = 964
    Height = 345
    Anchors = [akLeft, akRight, akBottom]
    Lines.Strings = (
      'MemoLog')
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object BitBtnCSV: TBitBtn
    Left = 145
    Top = 506
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'CSV file'
    Kind = bkOK
    NumGlyphs = 2
    TabOrder = 4
    Visible = False
    OnClick = BitBtnCSVClick
  end
  object PB: TProgressBar
    Left = 8
    Top = 471
    Width = 964
    Height = 17
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 5
  end
  object CheckBoxZeroPrice: TCheckBox
    Left = 288
    Top = 526
    Width = 353
    Height = 17
    Anchors = [akLeft, akBottom]
    Caption = #1059#1073#1088#1072#1090#1100' '#1082#1086#1084#1084#1077#1085#1090#1072#1088#1080#1081' '#1076#1083#1103' '#1090#1086#1074#1072#1088#1086#1074' '#1089' '#1085#1091#1083#1077#1074#1086#1081' '#1094#1077#1085#1086#1081
    Checked = True
    State = cbChecked
    TabOrder = 6
  end
  object CheckBoxZeroOstatki: TCheckBox
    Left = 288
    Top = 494
    Width = 377
    Height = 26
    Anchors = [akLeft, akBottom]
    Caption = #1059#1073#1088#1072#1090#1100' '#1082#1086#1084#1084#1077#1085#1090#1072#1088#1080#1081' '#1076#1083#1103' '#1086#1090#1089#1091#1090#1089#1090#1074#1091#1102#1097#1080#1093' ('#1085#1091#1083#1077#1074#1086#1077' '#1082#1086#1083#1080#1095#1077#1089#1090#1074#1086')'
    Checked = True
    State = cbChecked
    TabOrder = 7
  end
  object FileOpenDialog1: TFileOpenDialog
    DefaultExtension = '*.xls'
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'Excel'
        FileMask = '*.xls'
      end>
    Options = [fdoPathMustExist, fdoFileMustExist, fdoShareAware]
    Title = #1042#1099#1073#1077#1088#1080#1090#1077' '#1092#1072#1081#1083' '#1089' '#1086#1089#1090#1072#1090#1082#1072#1084#1080' remonline.ua'
    Left = 160
    Top = 128
  end
  object FileOpenDialog2: TFileOpenDialog
    DefaultExtension = 'export*.xlsx'
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'Prom files'
        FileMask = 'export*.xlsx'
      end>
    Options = [fdoPathMustExist, fdoFileMustExist, fdoShareAware]
    Title = #1059#1082#1072#1078#1080#1090#1077' '#1092#1072#1081#1083' '#1089' '#1087#1086#1079#1080#1094#1080#1103#1084#1080' '#1074' prom.ua'
    Left = 272
    Top = 128
  end
end
