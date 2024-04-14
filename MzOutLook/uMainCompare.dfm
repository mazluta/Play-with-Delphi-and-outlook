object MainCompareFrm: TMainCompareFrm
  Left = 0
  Top = 0
  Caption = 'MainCompareFrm'
  ClientHeight = 469
  ClientWidth = 1234
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
  OnDestroy = FormDestroy
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 265
    Width = 1234
    Height = 5
    Cursor = crVSplit
    Align = alTop
    Color = clSilver
    ParentColor = False
    ExplicitTop = 150
    ExplicitWidth = 319
  end
  object LeftPanel: TPanel
    Left = 0
    Top = 0
    Width = 1234
    Height = 265
    Align = alTop
    TabOrder = 0
    object Panel1: TPanel
      Left = 1
      Top = 1
      Width = 1232
      Height = 30
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object Label1: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 65
        Height = 24
        Align = alLeft
        Caption = 'Select Store :'
        Layout = tlCenter
        ExplicitHeight = 13
      end
      object cbTopStoreList: TComboBox
        AlignWithMargins = True
        Left = 74
        Top = 3
        Width = 350
        Height = 21
        Align = alLeft
        TabOrder = 0
        Text = 'cbTopStoreList'
        OnChange = cbTopStoreListChange
      end
      object btnTopCompare: TButton
        AlignWithMargins = True
        Left = 1129
        Top = 3
        Width = 100
        Height = 24
        Hint = 
          'The compare process check if all the message on the top area exi' +
          'sts in the bottom area.'#13#10#13#10'The compare done by checking the Sear' +
          'chKey'
        Align = alRight
        Caption = 'Compare'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnTopCompareClick
      end
      object btnGetTopStoreID: TButton
        AlignWithMargins = True
        Left = 430
        Top = 3
        Width = 107
        Height = 24
        Align = alLeft
        Caption = 'Get StoreID'
        TabOrder = 2
        OnClick = btnGetTopStoreIDClick
      end
      object btnCopyToBottom: TButton
        AlignWithMargins = True
        Left = 543
        Top = 3
        Width = 146
        Height = 24
        Align = alLeft
        Caption = 'Copy Mail To Bottom Folder'
        TabOrder = 3
        OnClick = btnCopyToBottomClick
        ExplicitLeft = 551
        ExplicitTop = 6
      end
      object btnDeleteTopSelectedMailItem: TButton
        AlignWithMargins = True
        Left = 695
        Top = 3
        Width = 200
        Height = 24
        Align = alLeft
        Caption = 'Delete Top Selected MailItem'
        TabOrder = 4
        OnClick = btnDeleteTopSelectedMailItemClick
      end
    end
    object Panel3: TPanel
      Left = 1
      Top = 31
      Width = 1232
      Height = 233
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object Splitter2: TSplitter
        Left = 356
        Top = 0
        Width = 5
        Height = 233
        Color = clSilver
        ParentColor = False
      end
      object tvTop: TTreeView
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 350
        Height = 227
        Align = alLeft
        HideSelection = False
        HotTrack = True
        Indent = 19
        ReadOnly = True
        RowSelect = True
        SortType = stText
        TabOrder = 0
        OnChange = tvTopChange
      end
      object TopMailsListGrid: TDBGrid
        Left = 361
        Top = 0
        Width = 871
        Height = 233
        Align = alClient
        DataSource = TopMailsListDS
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = TopMailsListGridDrawColumnCell
        OnDblClick = TopMailsListGridDblClick
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
          end
          item
            Expanded = False
            FieldName = 'SearchKey'
            Width = 300
            Visible = True
          end>
      end
    end
  end
  object RightPanel: TPanel
    Left = 0
    Top = 270
    Width = 1234
    Height = 199
    Align = alClient
    TabOrder = 1
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 1232
      Height = 30
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object Label2: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 65
        Height = 24
        Align = alLeft
        Caption = 'Select Store :'
        Layout = tlCenter
        ExplicitHeight = 13
      end
      object cbBottomStoreList: TComboBox
        AlignWithMargins = True
        Left = 74
        Top = 3
        Width = 350
        Height = 21
        Align = alLeft
        TabOrder = 0
        Text = 'ComboBox1'
        OnChange = cbBottomStoreListChange
      end
      object btnGetBottomStoreID: TButton
        AlignWithMargins = True
        Left = 430
        Top = 3
        Width = 107
        Height = 24
        Align = alLeft
        Caption = 'Get StoreID'
        TabOrder = 1
        OnClick = btnGetBottomStoreIDClick
      end
      object btnBottomCompare: TButton
        AlignWithMargins = True
        Left = 1129
        Top = 3
        Width = 100
        Height = 24
        Hint = 
          'The compare process check if all the message on the bottom area ' +
          'exists in the top area.'#13#10#13#10'The compare done by checking the Sear' +
          'chKey'
        Align = alRight
        Caption = 'Compare'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = btnBottomCompareClick
      end
      object btnCopyToTop: TButton
        AlignWithMargins = True
        Left = 543
        Top = 3
        Width = 146
        Height = 24
        Align = alLeft
        Caption = 'Copy Mail To Top Folder'
        TabOrder = 3
        OnClick = btnCopyToTopClick
      end
      object btnDeleteBottomSelectedMailItem: TButton
        AlignWithMargins = True
        Left = 695
        Top = 3
        Width = 200
        Height = 24
        Align = alLeft
        Caption = 'Delete Bottom Selected MailItem'
        TabOrder = 4
        OnClick = btnDeleteBottomSelectedMailItemClick
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 31
      Width = 1232
      Height = 167
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object Splitter3: TSplitter
        Left = 356
        Top = 0
        Width = 5
        Height = 167
        Color = clSilver
        ParentColor = False
        ExplicitLeft = 364
      end
      object tvBottom: TTreeView
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 350
        Height = 161
        Align = alLeft
        HideSelection = False
        HotTrack = True
        Indent = 19
        ReadOnly = True
        RowSelect = True
        SortType = stText
        TabOrder = 0
        OnChange = tvBottomChange
      end
      object BottomMailsListGrid: TDBGrid
        Left = 361
        Top = 0
        Width = 871
        Height = 167
        Align = alClient
        DataSource = BottomMailsListDS
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = BottomMailsListGridDrawColumnCell
        OnDblClick = BottomMailsListGridDblClick
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
          end
          item
            Expanded = False
            FieldName = 'SearchKey'
            Width = 300
            Visible = True
          end>
      end
    end
  end
  object TopMailsListDS: TDataSource
    DataSet = TopMailsListTbl
    Left = 496
    Top = 88
  end
  object TopMailsListTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Number'
        DataType = ftInteger
      end
      item
        Name = 'ReciveDate'
        DataType = ftDateTime
      end
      item
        Name = 'Subject'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'FromName'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'FromEmail'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'CC'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'BCC'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'StoreID'
        DataType = ftWideString
        Size = 5000
      end
      item
        Name = 'FolderID'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'EntryID'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'SearchKey'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'Error'
        DataType = ftBoolean
      end>
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
    Left = 560
    Top = 113
    object TopMailsListTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object TopMailsListTblReciveDate: TDateTimeField
      FieldName = 'ReciveDate'
    end
    object TopMailsListTblSubject: TWideStringField
      FieldName = 'Subject'
      Size = 250
    end
    object TopMailsListTblFromName: TWideStringField
      FieldName = 'FromName'
      Size = 250
    end
    object TopMailsListTblFromEmail: TWideStringField
      FieldName = 'FromEmail'
      Size = 250
    end
    object TopMailsListTblCC: TWideStringField
      FieldName = 'CC'
      Size = 250
    end
    object TopMailsListTblBCC: TWideStringField
      FieldName = 'BCC'
      Size = 250
    end
    object TopMailsListTblStoreID: TWideStringField
      FieldName = 'StoreID'
      Size = 5000
    end
    object TopMailsListTblFolderID: TWideStringField
      FieldName = 'FolderID'
      Size = 250
    end
    object TopMailsListTblEntryID: TWideStringField
      FieldName = 'EntryID'
      Size = 250
    end
    object TopMailsListTblSearchKey: TWideStringField
      FieldName = 'SearchKey'
      Size = 250
    end
    object TopMailsListTblError: TBooleanField
      FieldName = 'Error'
    end
  end
  object BottomMailsListTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Number'
        DataType = ftInteger
      end
      item
        Name = 'ReciveDate'
        DataType = ftDateTime
      end
      item
        Name = 'Subject'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'FromName'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'FromEmail'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'CC'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'BCC'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'StoreID'
        DataType = ftWideString
        Size = 5000
      end
      item
        Name = 'FolderID'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'EntryID'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'SearchKey'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'Error'
        DataType = ftBoolean
      end>
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
    Left = 504
    Top = 369
    object BottomMailsListTblNumber: TIntegerField
      FieldName = 'Number'
    end
    object BottomMailsListTblReciveDate: TDateTimeField
      FieldName = 'ReciveDate'
    end
    object BottomMailsListTblSubject: TWideStringField
      FieldName = 'Subject'
      Size = 250
    end
    object BottomMailsListTblFromName: TWideStringField
      FieldName = 'FromName'
      Size = 250
    end
    object BottomMailsListTblFromEmail: TWideStringField
      FieldName = 'FromEmail'
      Size = 250
    end
    object BottomMailsListTblCC: TWideStringField
      FieldName = 'CC'
      Size = 250
    end
    object BottomMailsListTblBCC: TWideStringField
      FieldName = 'BCC'
      Size = 250
    end
    object BottomMailsListTblStoreID: TWideStringField
      FieldName = 'StoreID'
      Size = 5000
    end
    object BottomMailsListTblFolderID: TWideStringField
      FieldName = 'FolderID'
      Size = 250
    end
    object BottomMailsListTblEntryID: TWideStringField
      FieldName = 'EntryID'
      Size = 250
    end
    object BottomMailsListTblSearchKey: TWideStringField
      FieldName = 'SearchKey'
      Size = 250
    end
    object BottomMailsListTblError: TBooleanField
      FieldName = 'Error'
    end
  end
  object BottomMailsListDS: TDataSource
    DataSet = BottomMailsListTbl
    Left = 440
    Top = 344
  end
end
