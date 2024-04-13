object MailPropsFrm: TMailPropsFrm
  Left = 0
  Top = 0
  Caption = 'Mail Props'
  ClientHeight = 494
  ClientWidth = 1221
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 453
    Width = 1221
    Height = 41
    Align = alBottom
    TabOrder = 0
    object btnViewAsPlainText: TButton
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 120
      Height = 33
      Align = alLeft
      Caption = 'View As Plain Text'
      TabOrder = 0
      OnClick = btnViewAsPlainTextClick
    end
    object btnViewAsHtmlBody: TButton
      AlignWithMargins = True
      Left = 130
      Top = 4
      Width = 120
      Height = 33
      Align = alLeft
      Caption = 'View As HTML'
      TabOrder = 1
      OnClick = btnViewAsHtmlBodyClick
    end
    object btnViewAsRTFbody: TButton
      AlignWithMargins = True
      Left = 256
      Top = 4
      Width = 120
      Height = 33
      Align = alLeft
      Caption = 'View As RTF'
      TabOrder = 2
      OnClick = btnViewAsRTFbodyClick
    end
    object btnSaveMessage: TButton
      AlignWithMargins = True
      Left = 508
      Top = 4
      Width = 150
      Height = 33
      Align = alLeft
      Caption = 'Save Message As MSG'
      TabOrder = 3
      OnClick = btnSaveMessageClick
    end
    object btnSaveAsMHTML: TButton
      AlignWithMargins = True
      Left = 664
      Top = 4
      Width = 150
      Height = 33
      Align = alLeft
      Caption = 'Save Message As MHTML'
      TabOrder = 4
      OnClick = btnSaveAsMHTMLClick
    end
    object btnViewAsMhtml: TButton
      AlignWithMargins = True
      Left = 382
      Top = 4
      Width = 120
      Height = 33
      Align = alLeft
      Caption = 'View As MHTML'
      TabOrder = 5
      OnClick = btnViewAsMhtmlClick
    end
  end
  object StoresListsGrid: TDBGrid
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 1215
    Height = 447
    Align = alClient
    DataSource = MailPropDS
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'Number'
        Width = 50
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PropName'
        Title.Caption = 'Prop Name'
        Width = 221
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PropType'
        Title.Caption = 'Prop Type'
        Width = 81
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PropValue'
        Title.Caption = 'Prop Value'
        Width = 400
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PropValueW'
        Title.Caption = 'Prop Value Wide String'
        Width = 400
        Visible = True
      end>
  end
  object MailPropDS: TDataSource
    DataSet = MailPropTbl
    Left = 120
    Top = 48
  end
  object MailPropTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Number'
        DataType = ftInteger
      end
      item
        Name = 'PropName'
        DataType = ftWideString
        Size = 150
      end
      item
        Name = 'PropType'
        DataType = ftWideString
        Size = 20
      end
      item
        Name = 'PropValue'
        DataType = ftWideString
        Size = 1000
      end
      item
        Name = 'PropValueW'
        DataType = ftWideString
        Size = 1000
      end>
    IndexDefs = <>
    IndexFieldNames = 'Number'
    Params = <>
    StoreDefs = True
    Left = 184
    Top = 73
    object MailPropTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object MailPropTblPropName: TWideStringField
      FieldName = 'PropName'
      Size = 150
    end
    object MailPropTblPropType: TWideStringField
      FieldName = 'PropType'
    end
    object MailPropTblPropValue: TWideStringField
      FieldName = 'PropValue'
      Size = 1000
    end
    object MailPropTblPropValueW: TWideStringField
      FieldName = 'PropValueW'
      Size = 1000
    end
  end
  object FileOpenDialog: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <>
    Options = []
    Left = 392
    Top = 168
  end
  object FileSaveDialog: TFileSaveDialog
    FavoriteLinks = <>
    FileTypes = <>
    Options = []
    Left = 392
    Top = 224
  end
end
