object FormMain: TFormMain
  Left = 0
  Top = 0
  Caption = #1054#1073#1088#1072#1073#1086#1090#1072#1090#1100' '#1087#1088#1072#1081#1089#1099' '#1076#1083#1103' Prom.ua'
  ClientHeight = 563
  ClientWidth = 833
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDblClick = FormDblClick
  DesignSize = (
    833
    563)
  PixelsPerInch = 96
  TextHeight = 13
  object MemoTxt: TMemo
    Left = 8
    Top = 8
    Width = 817
    Height = 228
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
    ScrollBars = ssBoth
    ShowHint = True
    TabOrder = 0
    Visible = False
    WantTabs = True
    WordWrap = False
  end
  object BitBtnXLS: TBitBtn
    Left = 8
    Top = 530
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
    Left = 750
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
    Top = 242
    Width = 817
    Height = 239
    Anchors = [akLeft, akRight, akBottom]
    Lines.Strings = (
      'MemoLog')
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object BitBtnCSV: TBitBtn
    Left = 208
    Top = 530
    Width = 121
    Height = 25
    Anchors = [akLeft, akBottom]
    Caption = 'CSV file'
    Kind = bkOK
    NumGlyphs = 2
    TabOrder = 4
    OnClick = BitBtnCSVClick
  end
  object PB: TProgressBar
    Left = 8
    Top = 496
    Width = 817
    Height = 17
    TabOrder = 5
  end
  object FileOpenDialog1: TFileOpenDialog
    DefaultExtension = '*.xls'
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'Excel'
        FileMask = '*.xls;*.xlsx'
      end>
    Options = [fdoPathMustExist, fdoFileMustExist, fdoShareAware]
    Left = 184
    Top = 120
  end
end