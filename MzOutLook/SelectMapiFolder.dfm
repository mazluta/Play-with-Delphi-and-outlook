object SelectMapiFolderFrm: TSelectMapiFolderFrm
  Left = 0
  Top = 0
  Caption = 'Select Mapi Folder'
  ClientHeight = 271
  ClientWidth = 444
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 444
    Height = 230
    Align = alClient
    TabOrder = 0
    object FoldersGrid: TDBGrid
      AlignWithMargins = True
      Left = 4
      Top = 34
      Width = 436
      Height = 192
      Align = alClient
      DataSource = FoldersListDS
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      OnDblClick = FoldersGridDblClick
      Columns = <
        item
          Expanded = False
          FieldName = 'FolderName'
          Title.Caption = 'Folder Name'
          Width = 400
          Visible = True
        end>
    end
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 442
      Height = 30
      Align = alTop
      TabOrder = 1
      ExplicitLeft = 2
      ExplicitTop = -2
      object lbSearchFor: TLabel
        AlignWithMargins = True
        Left = 4
        Top = 4
        Width = 52
        Height = 22
        Align = alLeft
        Caption = 'Search For'
        Layout = tlCenter
        ExplicitHeight = 13
      end
      object sbFromStart: TSpeedButton
        AlignWithMargins = True
        Left = 196
        Top = 4
        Width = 23
        Height = 22
        Align = alLeft
        Glyph.Data = {
          42020000424D4202000000000000420000002800000010000000100000000100
          1000030000000002000000000000000000000000000000000000007C0000E003
          00001F0000001F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
          1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
          1F7C2925CE391F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
          4A295A6B6B2D1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C4A29
          7B6F94521F7C1F7C1F7C1F7C1F7C1F7C1F7CB556524AEF3DEF3D524AEF3D5A6B
          734E1F7C1F7C1F7C1F7C1F7C1F7C1F7C734EF75EBD77BD7718639452EF3DEF3D
          1F7C1F7C1F7C1F7CE35DE361045A0B4AF65E1667B15E905AB25A9B7393528835
          C845E35D1F7C1F7C24662E7F4F7F724EBD77576F55735573566F586F7A6FEF3D
          D77F24661F7C1F7C45662E7F2B7F514ADE7B576F55735473566F376B7A6F1042
          B47F45661F7C1F7C866EDC7FFD7F734EDE7BBC779B779B737A6F7A6F7B6F734E
          FC7F866E1F7C1F7C866E6372026A2C56396BDE7BDD7BBD7BBD7BBD7717670956
          026A866E1F7C1F7C646E4F7F4F7F3077B25A5A6BDE7BFE7FDE7B396BB25A9677
          D77F646E1F7C1F7CA576507F2B7F2B7F0E77B0627252524A7252D25E7377B47F
          B47FA5761F7C1F7CA676DC7FFE7FFC7FDB7FDB7FDA7FDA7FDB7FDB7FFC7FFC7F
          FC7FA6761F7C1F7CA67A847A4272427242724272427242724272427242724272
          4272A67A1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C1F7C
          1F7C1F7C1F7C}
        ParentShowHint = False
        ShowHint = True
        OnClick = sbFromStartClick
        ExplicitLeft = 192
        ExplicitTop = 5
      end
      object sbFindnext: TSpeedButton
        AlignWithMargins = True
        Left = 225
        Top = 4
        Width = 23
        Height = 22
        Hint = 'Search For next Record'
        Align = alLeft
        Glyph.Data = {
          1E010000424D1E010000000000007600000028000000180000000E0000000100
          040000000000A800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777877777
          7777778777777777700877777777777877777777000087777777777787777770
          0000087777777777787777000000008777777777778770000000000877777777
          7778F00000000008F77F77778778FFFF00007FF7FFFF77778FF7777F00008777
          777F77778777777F00008777777F77778777777F00008777777F77778777777F
          00008777777F77778777777F00008777777F77778777777FFFFF7777777FFFFF
          7777}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = sbFindnextClick
        ExplicitLeft = 240
        ExplicitTop = 5
      end
      object sbFindPrior: TSpeedButton
        AlignWithMargins = True
        Left = 254
        Top = 4
        Width = 23
        Height = 22
        Hint = 'Search For prior Record'
        Align = alLeft
        Glyph.Data = {
          1E010000424D1E010000000000007600000028000000180000000E0000000100
          040000000000A800000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777788888777
          777788888777777F00008777777F77778777777F00008777777F77778777777F
          00008777777F77778777777F00008777777F77778777777F00008777777F7777
          8777788700008888788777778888F00000000008F77777777778F00000000007
          F777777777777F00000000777F777777777777F00000077777F777777777777F
          00007777777F777777777777F00777777777F777777777777F77777777777F77
          7777}
        NumGlyphs = 2
        ParentShowHint = False
        ShowHint = True
        OnClick = sbFindPriorClick
        ExplicitLeft = 280
        ExplicitTop = 5
      end
      object edSearchFor: TEdit
        AlignWithMargins = True
        Left = 62
        Top = 4
        Width = 128
        Height = 22
        Align = alLeft
        TabOrder = 0
        Text = 'edSearchFor'
        OnKeyPress = edSearchForKeyPress
        ExplicitLeft = 41
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 230
    Width = 444
    Height = 41
    Align = alBottom
    TabOrder = 1
    object btnSelect: TButton
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 75
      Height = 33
      Align = alLeft
      Caption = 'Select'
      TabOrder = 0
      OnClick = btnSelectClick
    end
    object btnReturn: TButton
      AlignWithMargins = True
      Left = 85
      Top = 4
      Width = 75
      Height = 33
      Align = alLeft
      Caption = 'Return'
      TabOrder = 1
      OnClick = btnReturnClick
    end
  end
  object FoldersListDS: TDataSource
    DataSet = FoldersListTbl
    Left = 160
    Top = 64
  end
  object FoldersListTbl: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Number'
        DataType = ftInteger
      end
      item
        Name = 'FolderName'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'EntryId'
        DataType = ftWideString
        Size = 250
      end
      item
        Name = 'StoreID'
        DataType = ftWideString
        Size = 5000
      end
      item
        Name = 'Newname'
        DataType = ftWideString
        Size = 150
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 224
    Top = 89
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
end
