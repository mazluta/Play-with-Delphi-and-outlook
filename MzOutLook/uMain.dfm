object PlayWithOotlookFrm: TPlayWithOotlookFrm
  Left = 0
  Top = 0
  Caption = 'PlayWithOotlookFrm'
  ClientHeight = 611
  ClientWidth = 1180
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 169
    Width = 1180
    Height = 5
    Cursor = crVSplit
    Align = alTop
    Color = clGray
    ParentColor = False
  end
  object Panel1: TPanel
    Left = 0
    Top = 570
    Width = 1180
    Height = 41
    Align = alBottom
    TabOrder = 0
    object sbBuildFoldersList: TButton
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 139
      Height = 33
      Align = alLeft
      Caption = 'Build Folders List'
      TabOrder = 0
      OnClick = sbBuildFoldersListClick
    end
    object sbGetProps: TButton
      AlignWithMargins = True
      Left = 149
      Top = 4
      Width = 98
      Height = 33
      Align = alLeft
      Caption = 'Get Props'
      TabOrder = 1
      OnClick = sbGetPropsClick
    end
    object btnMoveMailItemToOtherFolder: TButton
      AlignWithMargins = True
      Left = 822
      Top = 4
      Width = 172
      Height = 33
      Align = alRight
      Caption = 'Move Item To other Folder'
      TabOrder = 2
      OnClick = btnMoveMailItemToOtherFolderClick
      ExplicitLeft = 253
    end
    object btnDeleteSelectedMailItem: TButton
      AlignWithMargins = True
      Left = 1000
      Top = 4
      Width = 176
      Height = 33
      Align = alRight
      Caption = 'Delete The Mail Item'
      TabOrder = 3
      OnClick = btnDeleteSelectedMailItemClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1180
    Height = 169
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Splitter2: TSplitter
      Left = 457
      Top = 0
      Width = 5
      Height = 169
      Color = clGray
      ParentColor = False
    end
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 457
      Height = 169
      Align = alLeft
      TabOrder = 0
      object StoresListsGrid: TDBGrid
        AlignWithMargins = True
        Left = 4
        Top = 4
        Width = 449
        Height = 120
        Align = alClient
        DataSource = StoresListDS
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 0
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
            FieldName = 'StoreName'
            Width = 264
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StoreID'
            Width = 500
            Visible = True
          end>
      end
      object Panel6: TPanel
        Left = 1
        Top = 127
        Width = 455
        Height = 41
        Align = alBottom
        BevelOuter = bvNone
        TabOrder = 1
        object btnAddMsgStore: TButton
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 120
          Height = 35
          Align = alLeft
          Caption = 'Add MsgStore'
          TabOrder = 0
          OnClick = btnAddMsgStoreClick
        end
        object btnRemoveMsgStore: TButton
          AlignWithMargins = True
          Left = 129
          Top = 3
          Width = 120
          Height = 35
          Align = alLeft
          Caption = 'Remove MsgStore'
          TabOrder = 1
          OnClick = btnRemoveMsgStoreClick
        end
      end
    end
    object Panel4: TPanel
      Left = 462
      Top = 0
      Width = 718
      Height = 169
      Align = alClient
      TabOrder = 1
      object FoldersListGrid: TDBGrid
        AlignWithMargins = True
        Left = 4
        Top = 4
        Width = 710
        Height = 161
        Align = alClient
        DataSource = FoldersListDS
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 0
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
            FieldName = 'FolderName'
            Width = 300
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'EntryId'
            Width = 300
            Visible = True
          end>
      end
    end
  end
  object Panel5: TPanel
    Left = 0
    Top = 174
    Width = 1180
    Height = 396
    Align = alClient
    TabOrder = 2
    object Label1: TLabel
      Left = 1
      Top = 382
      Width = 1178
      Height = 13
      Align = alBottom
      Caption = 'Click the column title to sort by that field'
      ExplicitWidth = 193
    end
    object MailsListGrid: TDBGrid
      Left = 1
      Top = 1
      Width = 1178
      Height = 381
      Align = alClient
      DataSource = MailsListDS
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      OnDblClick = MailsListGridDblClick
      OnTitleClick = MailsListGridTitleClick
      Columns = <
        item
          Expanded = False
          FieldName = 'Number'
          Width = 50
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'Subject'
          Width = 425
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ReciveDate'
          Width = 158
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'FromEmail'
          Width = 290
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'FromName'
          Width = 200
          Visible = True
        end>
    end
  end
  object MailsListDS: TDataSource
    DataSet = MailsListTbl
    Left = 224
    Top = 256
  end
  object MailsListTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <
      item
        Name = 'ByNumber'
        Fields = 'Number'
      end
      item
        Name = 'BySubject'
        Fields = 'Subject;Number'
      end
      item
        Name = 'ByReciveDate'
        Fields = 'ReciveDate;Number'
      end
      item
        Name = 'ByFromEmail'
        Fields = 'FromEmail;Number'
      end
      item
        Name = 'ByFromName'
        Fields = 'FromName;Number'
      end>
    IndexFieldNames = 'Number'
    Params = <>
    StoreDefs = True
    Left = 288
    Top = 281
    object MailsListTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object MailsListTblReciveDate: TDateTimeField
      FieldName = 'ReciveDate'
    end
    object MailsListTblSubject: TWideStringField
      FieldName = 'Subject'
      Size = 250
    end
    object MailsListTblFromName: TWideStringField
      FieldName = 'FromName'
      Size = 250
    end
    object MailsListTblFromEmail: TWideStringField
      FieldName = 'FromEmail'
      Size = 250
    end
    object MailsListTblCC: TWideStringField
      FieldName = 'CC'
      Size = 250
    end
    object MailsListTblBCC: TWideStringField
      FieldName = 'BCC'
      Size = 250
    end
    object MailsListTblStoreID: TWideStringField
      FieldName = 'StoreID'
      Size = 5000
    end
    object MailsListTblFolderID: TWideStringField
      FieldName = 'FolderID'
      Size = 250
    end
    object MailsListTblEntryID: TWideStringField
      FieldName = 'EntryID'
      Size = 250
    end
    object MailsListTblSearchKey: TWideStringField
      FieldName = 'SearchKey'
      Size = 250
    end
  end
  object StoresListDS: TDataSource
    DataSet = StoresListTbl
    Left = 120
    Top = 48
  end
  object StoresListTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Number'
        DataType = ftInteger
      end
      item
        Name = 'StoreName'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'StoreID'
        DataType = ftWideString
        Size = 5000
      end
      item
        Name = 'FullName'
        DataType = ftWideString
        Size = 250
      end>
    IndexDefs = <>
    IndexFieldNames = 'Number'
    Params = <>
    StoreDefs = True
    AfterScroll = StoresListTblAfterScroll
    Left = 184
    Top = 73
    object StoresListTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object StoresListTblStoreName: TWideStringField
      FieldName = 'StoreName'
      Size = 250
    end
    object StoresListTblStoreID: TWideStringField
      FieldName = 'StoreID'
      Size = 5000
    end
    object StoresListTblFullName: TWideStringField
      FieldName = 'FullName'
      Size = 250
    end
  end
  object FoldersListTbl: TClientDataSet
    Aggregates = <>
    IndexFieldNames = 'Number'
    Params = <>
    AfterScroll = FoldersListTblAfterScroll
    Left = 632
    Top = 81
    object FoldersListTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object FoldersListTblFolderName: TWideStringField
      FieldName = 'FolderName'
      Size = 250
    end
    object FoldersListTblEntryId: TWideStringField
      FieldName = 'EntryId'
      Size = 250
    end
    object FoldersListTblStoreID: TWideStringField
      FieldName = 'StoreID'
      Size = 5000
    end
    object FoldersListTblNewname: TWideStringField
      FieldName = 'Newname'
      Size = 150
    end
  end
  object FoldersListDS: TDataSource
    DataSet = FoldersListTbl
    Left = 568
    Top = 56
  end
  object FileOpenDialog: TFileOpenDialog
    DefaultExtension = 'pst'
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'PST file (*.pst)'
        FileMask = '*.pst'
      end>
    Options = [fdoPathMustExist, fdoFileMustExist]
    Left = 336
    Top = 62
  end
end
