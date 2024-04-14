unit uMainCompare;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.ComCtrls,
  Data.DB, Datasnap.DBClient,
  Outlook2010, ActiveX, ComObj, ole2, OleServers, Vcl.OleServer, Vcl.Grids,
  Vcl.DBGrids;

type
  PMyRec = ^TMyRec;
  TMyRec = record
    EnteryID: string;
  end;

type
  TMainCompareFrm = class(TForm)
    LeftPanel: TPanel;
    RightPanel: TPanel;
    Splitter1: TSplitter;
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    cbTopStoreList: TComboBox;
    cbBottomStoreList: TComboBox;
    Label2: TLabel;
    btnTopCompare: TButton;
    Panel3: TPanel;
    Panel4: TPanel;
    tvTop: TTreeView;
    tvBottom: TTreeView;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    TopMailsListDS: TDataSource;
    TopMailsListTbl: TClientDataSet;
    TopMailsListTblNumber: TIntegerField;
    TopMailsListTblReciveDate: TDateTimeField;
    TopMailsListTblSubject: TWideStringField;
    TopMailsListTblFromName: TWideStringField;
    TopMailsListTblFromEmail: TWideStringField;
    TopMailsListTblCC: TWideStringField;
    TopMailsListTblBCC: TWideStringField;
    TopMailsListTblStoreID: TWideStringField;
    TopMailsListTblFolderID: TWideStringField;
    TopMailsListTblEntryID: TWideStringField;
    TopMailsListTblSearchKey: TWideStringField;
    TopMailsListGrid: TDBGrid;
    BottomMailsListTbl: TClientDataSet;
    BottomMailsListTblNumber: TIntegerField;
    BottomMailsListTblReciveDate: TDateTimeField;
    BottomMailsListTblSubject: TWideStringField;
    BottomMailsListTblFromName: TWideStringField;
    BottomMailsListTblFromEmail: TWideStringField;
    BottomMailsListTblCC: TWideStringField;
    BottomMailsListTblBCC: TWideStringField;
    BottomMailsListTblStoreID: TWideStringField;
    BottomMailsListTblFolderID: TWideStringField;
    BottomMailsListTblEntryID: TWideStringField;
    BottomMailsListTblSearchKey: TWideStringField;
    BottomMailsListDS: TDataSource;
    BottomMailsListGrid: TDBGrid;
    btnGetTopStoreID: TButton;
    btnGetBottomStoreID: TButton;
    btnBottomCompare: TButton;
    TopMailsListTblError: TBooleanField;
    BottomMailsListTblError: TBooleanField;
    btnCopyToTop: TButton;
    btnCopyToBottom: TButton;
    btnDeleteTopSelectedMailItem: TButton;
    btnDeleteBottomSelectedMailItem: TButton;
    procedure FormResize(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cbTopStoreListChange(Sender: TObject);
    procedure cbBottomStoreListChange(Sender: TObject);
    procedure btnGetTopStoreIDClick(Sender: TObject);
    procedure tvTopChange(Sender: TObject; Node: TTreeNode);
    procedure TopMailsListGridDblClick(Sender: TObject);
    procedure tvBottomChange(Sender: TObject; Node: TTreeNode);
    procedure btnGetBottomStoreIDClick(Sender: TObject);
    procedure TopMailsListGridDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnTopCompareClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnBottomCompareClick(Sender: TObject);
    procedure BottomMailsListGridDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure btnCopyToTopClick(Sender: TObject);
    procedure BottomMailsListGridDblClick(Sender: TObject);
    procedure btnCopyToBottomClick(Sender: TObject);
    procedure btnDeleteTopSelectedMailItemClick(Sender: TObject);
    procedure btnDeleteBottomSelectedMailItemClick(Sender: TObject);
  private
    { Private declarations }
    OnFirstTime       : Boolean;
    AvoideBuildTree   : Boolean;
    AvoideBuildTop    : Boolean;
    AvoideBuildBottom : Boolean;
    StoreIdList       : TStringList;
    OutlookApp        : Outlook2010.TOutlookApplication;
    procedure ClearTopTables;
    procedure ClearBottomTables;
    procedure ClearAllTables;
    procedure InitStoreList;
    procedure BuildTreeList(StoreID : String; TV : TTreeView);
    procedure AddChilde(TV : TTreeView; ParentNode : TTreeNode; FL : Outlook2010._Folders);
  public
    { Public declarations }
  end;

var
  MainCompareFrm: TMainCompareFrm;

implementation
{$R *.dfm}

Uses
  uOutlookPR,
  uMailProps;

procedure TMainCompareFrm.cbBottomStoreListChange(Sender: TObject);
Var
  StoreID   : String;
begin
  if AvoideBuildTree Then
    Exit;
  if StoreIdList.Count = 0 Then
    Exit;

  AvoideBuildTree := True;
  Try
    ClearBottomTables;
    StoreID := StoreIdList.Strings[cbBottomStoreList.ItemIndex];
    BuildTreeList(StoreID, tvBottom);
    tvBottom.SortType := stNone;
    Application.ProcessMessages;
    tvBottom.SortType := stText; // this will force the treeview to be sorted
    Application.ProcessMessages;
  Finally
    AvoideBuildTree := False;
  End;
end;

procedure TMainCompareFrm.cbTopStoreListChange(Sender: TObject);
Var
  StoreID   : String;
begin
  if AvoideBuildTree Then
    Exit;
  if StoreIdList.Count = 0 Then
    Exit;

  AvoideBuildTree := True;
  Try
    ClearTopTables;
    StoreID := StoreIdList.Strings[cbTopStoreList.ItemIndex];
    BuildTreeList(StoreID, tvTop);
    tvTop.SortType := stNone;
    Application.ProcessMessages;
    tvTop.SortType := stText; // this will force the treeview to be sorted
    Application.ProcessMessages;
  Finally
    AvoideBuildTree := False;
  End;
end;

procedure TMainCompareFrm.FormActivate(Sender: TObject);
begin
  if OnFirstTime Then
  Try
    ClearAllTables;

    OutlookApp := TOutlookApplication.Create(Nil);
    OutlookApp.ConnectKind := ckRunningOrNew;
    OutlookApp.Connect;

    InitStoreList;
  Finally
    OnFirstTime := False
  End;
end;

procedure TMainCompareFrm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  OutlookApp.Disconnect;
  OutlookApp.Quit;
  FreeAndNil(OutlookApp);
end;

procedure TMainCompareFrm.FormCreate(Sender: TObject);
begin
  OnFirstTime := True;
  AvoideBuildTree   := False;
  AvoideBuildTop    := False;
  AvoideBuildBottom := False;
  StoreIdList       := TStringList.Create;
  Self.Height       := Screen.Height - 20;
  Self.Width        := Screen.Width - 20;
  Self.Left         := 10;
  Self.Top          := 10;
end;

procedure TMainCompareFrm.FormDestroy(Sender: TObject);
begin
  Try FreeAndNil(StoreIdList); Except; End;
end;

procedure TMainCompareFrm.FormResize(Sender: TObject);
begin
  leftPanel.Height := Trunc(Self.ClientHeight/2);
end;

procedure TMainCompareFrm.ClearTopTables;
begin
  TopMailsListTbl.Close;
  TopMailsListTbl.CreateDataSet;
  TopMailsListTbl.Open;
  TopMailsListTbl.EmptyDataSet;
end;

procedure TMainCompareFrm.ClearBottomTables;
begin
  BottomMailsListTbl.Close;
  BottomMailsListTbl.CreateDataSet;
  BottomMailsListTbl.Open;
  BottomMailsListTbl.EmptyDataSet;
end;

procedure TMainCompareFrm.ClearAllTables;
begin
  ClearTopTables;
  ClearBottomTables;
end;

procedure TMainCompareFrm.InitStoreList;
Var
  NS        : Outlook2010._NameSpace;
  olStores  : Outlook2010._Stores;
  olStore   : Outlook2010._Store;
  CurStore  : Integer;
  StoreID_0 : String;
  StoreID_1 : String;
begin
  StoreID_0 := '';
  StoreID_1 := '';

  StoreIdList.Clear;

  AvoideBuildTree := True;
  Try
    cbTopStoreList.Items.Clear;
    cbBottomStoreList.Items.Clear;

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      Try
        olStores := NS.Stores;
        for CurStore := 1 to olStores.Count do
        begin
          olStore := olStores.Item(CurStore);

          StoreIdList.Add(olStore.StoreID);
          cbTopStoreList.Items.Add(olStore.DisplayName);
          cbBottomStoreList.Items.Add(olStore.DisplayName);
          case CurStore of
             1 : StoreID_0 := olStore.StoreID;
             2 : StoreID_1 := olStore.StoreID;
          end;
        end;
      Finally
      End;
    Finally
      olStore := nil;
      olStores := nil;
      NS := nil;
    End;

    tvTop.Items.Clear;
    tvBottom.Items.Clear;

    cbTopStoreList.Text := '';
    cbBottomStoreList.Text := '';
    cbTopStoreList.ItemIndex := -1;
    cbBottomStoreList.ItemIndex := -1;
    if cbTopStoreList.Items.Count > 0 Then
      cbTopStoreList.ItemIndex := 0;
    if cbBottomStoreList.Items.Count > 0 Then
      cbBottomStoreList.ItemIndex := 1;

    BuildTreeList(StoreID_0, tvTop);
    if StoreIdList.Count > 1 Then
    begin
      BuildTreeList(StoreID_1, tvBottom);
    end;
  Finally
    AvoideBuildTree := False;
  End;
end;

procedure TMainCompareFrm.TopMailsListGridDblClick(Sender: TObject);
Var
  NS      : Outlook2010._NameSpace;
  StoreID : String;
  EntryID : String;
  ItemToShow : System.IDispatch;
begin
  if not TopMailsListTbl.Active then
    Exit;
  if TopMailsListTbl.RecordCount = 0 then
    Exit;

  StoreID := TopMailsListTbl.FieldByName('StoreID').AsString;
  EntryID := TopMailsListTbl.FieldByName('EntryID').AsString;

  NS := OutlookApp.GetNameSpace('MAPI');

  ItemToShow  := NS.GetItemFromID(EntryID, OleVariant(StoreID));
  if Assigned(ItemToShow) then
  begin
    Try
      MailPropsFrm := TMailPropsFrm.Create(Application);
      MailPropsFrm.MI := (ItemToShow as MailItem);
      MailPropsFrm.ShowModal;
    Finally
      FreeAndNil(MailPropsFrm);
    End;
  end;
end;

procedure TMainCompareFrm.TopMailsListGridDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if TopMailsListTbl.FieldByName('Error').AsBoolean then
  begin
    TopMailsListGrid.Canvas.Font.Color := clMaroon;
  end;
  TopMailsListGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TMainCompareFrm.tvBottomChange(Sender: TObject; Node: TTreeNode);
Var
  NS        : Outlook2010._NameSpace;
  StoreID   : String;
  FolderID  : String;
  StrA      : String;
  StrW      : WideString;
  olPA      : _PropertyAccessor;
  FO        : MAPIFolder;
  MyRecPtr  : PMyRec;
  foItems   : Outlook2010.Items;
  foItem    : Outlook2010.MailItem;
  CurItem   : Integer;
  CurCursor  : TCursor;
begin
  if AvoideBuildBottom or AvoideBuildTree Then
    Exit;

  Try
    AvoideBuildBottom := True;
    CurCursor  := Screen.Cursor;
    Screen.Cursor := crHourGlass;

    ClearBottomTables;

    StoreID := StoreIdList.Strings[cbBottomStoreList.ItemIndex];

    MyRecPtr := Node.Data;
    FolderID := MyRecPtr.EnteryID;

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      FO := NS.GetFolderFromID(FolderID, OleVariant(StoreID));
      foItems := FO.Items;
      Try
        foItems.Sort('ReceivedTime',True);
      Except;
      End;

      Try
        BottomMailsListTbl.DisableControls;
        For CurItem := 1 to foItems.Count Do
        begin
          if Supports(foItems.Item(CurItem), MailItem, foItem) then
          begin
            olPA := foItem.PropertyAccessor;
            StrA := olPA.BinaryToString(foItem.PropertyAccessor.GetProperty(PR_SEARCH_KEY));
            BottomMailsListTbl.Insert;
            BottomMailsListTbl.FieldByName('Number').AsInteger      := CurItem;
            Try
              BottomMailsListTbl.FieldByName('ReciveDate').AsDateTime := foItem.ReceivedTime;
            Except;
              BottomMailsListTbl.FieldByName('ReciveDate').AsDateTime := NULL;
            End;
            BottomMailsListTbl.FieldByName('Subject').AsString      := foItem.Subject;
            BottomMailsListTbl.FieldByName('FromName').AsString     := foItem.SenderName;
            BottomMailsListTbl.FieldByName('FromEmail').AsString    := foItem.SenderEmailAddress;
            BottomMailsListTbl.FieldByName('CC').AsString           := foItem.CC;
            BottomMailsListTbl.FieldByName('BCC').AsString          := foItem.BCC;
            BottomMailsListTbl.FieldByName('StoreID').AsString      := StoreID;
            BottomMailsListTbl.FieldByName('FolderID').AsString     := FolderID;
            BottomMailsListTbl.FieldByName('EntryID').AsString      := foItem.EntryID;
            BottomMailsListTbl.FieldByName('SearchKey').AsString    := StrA;
            BottomMailsListTbl.Post;
          end;
        end;

        if BottomMailsListTbl.RecordCount > 0 Then
        begin
          BottomMailsListTbl.IndexName := 'ByNumber';
          BottomMailsListTbl.First;
        end;
      Finally
        BottomMailsListTbl.EnableControls;
      End;
    Finally
      foItem := nil;
      foItems := nil;
      FO := nil;
      NS := nil;
    End;
  Finally
    AvoideBuildBottom := False;
    Screen.Cursor  := CurCursor;
  End;
end;

procedure TMainCompareFrm.tvTopChange(Sender: TObject; Node: TTreeNode);
Var
  NS        : Outlook2010._NameSpace;
  StoreID   : String;
  FolderID  : String;
  StrA      : String;
  StrW      : WideString;
  olPA      : _PropertyAccessor;
  FO        : MAPIFolder;
  MyRecPtr  : PMyRec;
  foItems   : Outlook2010.Items;
  foItem    : Outlook2010.MailItem;
  CurItem   : Integer;
  CurCursor  : TCursor;
begin
  if AvoideBuildTop or AvoideBuildTree Then
    Exit;

  Try
    AvoideBuildTop    := True;

    CurCursor  := Screen.Cursor;
    Screen.Cursor := crHourGlass;

    ClearTopTables;

    StoreID := StoreIdList.Strings[cbTopStoreList.ItemIndex];

    MyRecPtr := Node.Data;
    FolderID := MyRecPtr.EnteryID;

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      FO := NS.GetFolderFromID(FolderID, OleVariant(StoreID));
      foItems := FO.Items;
      Try
        foItems.Sort('ReceivedTime',True);
      Except;
      End;

      Try
        TopMailsListTbl.DisableControls;
        For CurItem := 1 to foItems.Count Do
        begin
          if Supports(foItems.Item(CurItem), MailItem, foItem) then
          begin
            olPA := foItem.PropertyAccessor;
            StrA := olPA.BinaryToString(foItem.PropertyAccessor.GetProperty(PR_SEARCH_KEY));
            TopMailsListTbl.Insert;
            TopMailsListTbl.FieldByName('Number').AsInteger      := CurItem;
            Try
              TopMailsListTbl.FieldByName('ReciveDate').AsDateTime := foItem.ReceivedTime;
            Except;
              TopMailsListTbl.FieldByName('ReciveDate').AsDateTime := NULL;
            End;
            TopMailsListTbl.FieldByName('Subject').AsString      := foItem.Subject;
            TopMailsListTbl.FieldByName('FromName').AsString     := foItem.SenderName;
            TopMailsListTbl.FieldByName('FromEmail').AsString    := foItem.SenderEmailAddress;
            TopMailsListTbl.FieldByName('CC').AsString           := foItem.CC;
            TopMailsListTbl.FieldByName('BCC').AsString          := foItem.BCC;
            TopMailsListTbl.FieldByName('StoreID').AsString      := StoreID;
            TopMailsListTbl.FieldByName('FolderID').AsString     := FolderID;
            TopMailsListTbl.FieldByName('EntryID').AsString      := foItem.EntryID;
            TopMailsListTbl.FieldByName('SearchKey').AsString    := StrA;
            TopMailsListTbl.Post;
          end;
        end;

        if TopMailsListTbl.RecordCount > 0 Then
        begin
          TopMailsListTbl.IndexName := 'ByNumber';
          TopMailsListTbl.First;
        end;
      Finally
        TopMailsListTbl.EnableControls;
      End;
    Finally
      foItem := nil;
      foItems := nil;
      FO := nil;
      NS := nil;
    End;
  Finally
    AvoideBuildTop    := False;
    Screen.Cursor  := CurCursor;
  End;
end;

procedure TMainCompareFrm.BottomMailsListGridDblClick(Sender: TObject);
Var
  NS      : Outlook2010._NameSpace;
  StoreID : String;
  EntryID : String;
  ItemToShow : System.IDispatch;
begin
  if not BottomMailsListTbl.Active then
    Exit;
  if BottomMailsListTbl.RecordCount = 0 then
    Exit;

  StoreID := BottomMailsListTbl.FieldByName('StoreID').AsString;
  EntryID := BottomMailsListTbl.FieldByName('EntryID').AsString;

  NS := OutlookApp.GetNameSpace('MAPI');

  ItemToShow  := NS.GetItemFromID(EntryID, OleVariant(StoreID));
  if Assigned(ItemToShow) then
  begin
    Try
      MailPropsFrm := TMailPropsFrm.Create(Application);
      MailPropsFrm.MI := (ItemToShow as MailItem);
      MailPropsFrm.ShowModal;
    Finally
      FreeAndNil(MailPropsFrm);
    End;
  end;
end;

procedure TMainCompareFrm.BottomMailsListGridDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if BottomMailsListTbl.FieldByName('Error').AsBoolean then
  begin
    BottomMailsListGrid.Canvas.Font.Color := clMaroon;
  end;
  BottomMailsListGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TMainCompareFrm.btnBottomCompareClick(Sender: TObject);
Var
  StoreName : String;
  StoreID_0 : String;
  StoreID_1 : String;
  EntryID   : String;
  SearchKey : String;
  RecNoB    : Integer;
  RecNoT    : Integer;
begin
  StoreID_0 := '';
  StoreID_1 := '';

  if not TopMailsListTbl.Active then
    Exit;
  if TopMailsListTbl.RecordCount = 0 then
    Exit;
  if not BottomMailsListTbl.Active then
    Exit;
  if BottomMailsListTbl.RecordCount = 0 then
    Exit;

  StoreID_0 := StoreIdList.Strings[cbTopStoreList.ItemIndex];
  StoreID_1 := StoreIdList.Strings[cbBottomStoreList.ItemIndex];

  if (StoreID_0 = StoreID_1) then
    raise System.SysUtils.Exception.Create('Please select different Store to compare');

  RecNoB := BottomMailsListTbl.RecNo;
  RecNoT := TopMailsListTbl.RecNo;
  Try
    LockWindowUpdate(Self.Handle);
    BottomMailsListTbl.DisableControls;
    TopMailsListTbl.DisableControls;

    BottomMailsListTbl.First;
    While not BottomMailsListTbl.EOF Do
    begin
      SearchKey := BottomMailsListTbl.FieldByName('SearchKey').AsString;
      if not TopMailsListTbl.Locate('SearchKey', SearchKey, []) then
      begin
        BottomMailsListTbl.Edit;
        BottomMailsListTbl.FieldByName('Error').AsBoolean := True;
        BottomMailsListTbl.Post;
      end
      else
      begin
        BottomMailsListTbl.Edit;
        BottomMailsListTbl.FieldByName('Error').AsBoolean := False;
        BottomMailsListTbl.Post;
      end;
      BottomMailsListTbl.Next;
    end;
  Finally
    TopMailsListTbl.RecNo := RecNoT;
    BottomMailsListTbl.RecNo := RecNoB;
    BottomMailsListTbl.EnableControls;
    TopMailsListTbl.EnableControls;
    LockWindowUpdate(0);
  End;
end;

procedure TMainCompareFrm.btnCopyToBottomClick(Sender: TObject);
Var
  NS         : Outlook2010._NameSpace;
  StoreID    : String;
  FolderID   : String;
  MailID     : String;
  FO         : Outlook2010.MAPIFolder;
  MyRecPtr   : PMyRec;
  StoreIDToCopy : String;
  MailItem      : System.IDispatch;
  ItemToCopy    : System.IDispatch;
begin
  if MessageDlg('Sure to copy the Selected Item to the selected Store/Folder of the bottom part?',
                   mtConfirmation, [mbYes,mbNo], 0, mbNo) = mrNo then
    exit;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    StoreID   := TopMailsListTbl.FieldByName('StoreID').AsString;
    MailID    := TopMailsListTbl.FieldByName('EntryID').AsString;

    MailItem   := NS.GetItemFromID(MailID, OleVariant(StoreID));
    ItemToCopy := (MailItem as Outlook2010.MailItem).Copy;

    StoreIDToCopy := StoreIdList.Strings[cbBottomStoreList.ItemIndex];
    MyRecPtr := tvBottom.Selected.Data;
    FolderID := MyRecPtr.EnteryID;
    FO := NS.GetFolderFromID(FolderID, OleVariant(StoreIDToCopy));
    if Assigned(ItemToCopy) and Assigned(FO) then
    begin
      (ItemToCopy as Outlook2010.MailItem).Move(FO as Outlook2010.MAPIFolder);
      tvBottomChange(tvBottom, tvBottom.Selected);
    end;
  Finally
    MailItem := nil;
    ItemToCopy := nil;
    FO := nil;
    NS := nil;
  End;
end;

procedure TMainCompareFrm.btnCopyToTopClick(Sender: TObject);
Var
  NS         : Outlook2010._NameSpace;
  StoreID    : String;
  FolderID   : String;
  MailID     : String;
  FO         : Outlook2010.MAPIFolder;
  MyRecPtr   : PMyRec;
  StoreIDToCopy : String;
  MailItem      : System.IDispatch;
  ItemToCopy    : System.IDispatch;
begin
  if MessageDlg('Sure to copy the Selected Item to the Store/Folder to the above select?',
                   mtConfirmation, [mbYes,mbNo], 0, mbNo) = mrNo then
    exit;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    StoreID   := BottomMailsListTbl.FieldByName('StoreID').AsString;
    MailID    := BottomMailsListTbl.FieldByName('EntryID').AsString;

    MailItem   := NS.GetItemFromID(MailID, OleVariant(StoreID));
    ItemToCopy := (MailItem as Outlook2010.MailItem).Copy;

    StoreIDToCopy := StoreIdList.Strings[cbTopStoreList.ItemIndex];
    MyRecPtr := tvTop.Selected.Data;
    FolderID := MyRecPtr.EnteryID;
    FO := NS.GetFolderFromID(FolderID, OleVariant(StoreIDToCopy));
    if Assigned(ItemToCopy) and Assigned(FO) then
    begin
      (ItemToCopy as Outlook2010.MailItem).Move(FO as Outlook2010.MAPIFolder);
      tvTopChange(tvTop, tvTop.Selected);
    end;
  Finally
    MailItem := nil;
    ItemToCopy := nil;
    FO := nil;
    NS := nil;
  End;
end;

procedure TMainCompareFrm.btnDeleteBottomSelectedMailItemClick(Sender: TObject);
Var
  NS         : Outlook2010._NameSpace;
  StoreID    : String;
  MailID     : String;
  MailItem      : System.IDispatch;
begin
  if MessageDlg('Sure to Delete the Selected MailItem ?',
                   mtConfirmation, [mbYes,mbNo], 0, mbNo) = mrNo then
    exit;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    StoreID   := BottomMailsListTbl.FieldByName('StoreID').AsString;
    MailID    := BottomMailsListTbl.FieldByName('EntryID').AsString;

    MailItem   := NS.GetItemFromID(MailID, OleVariant(StoreID));
    (MailItem as Outlook2010.MailItem).Delete;
    tvBottomChange(tvBottom, tvBottom.Selected);
  Finally
    MailItem := nil;
    NS := nil;
  End;
end;

procedure TMainCompareFrm.btnDeleteTopSelectedMailItemClick(Sender: TObject);
Var
  NS         : Outlook2010._NameSpace;
  StoreID    : String;
  MailID     : String;
  MailItem      : System.IDispatch;
begin
  if MessageDlg('Sure to Delete the Selected MailItem ?',
                   mtConfirmation, [mbYes,mbNo], 0, mbNo) = mrNo then
    exit;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    StoreID   := TopMailsListTbl.FieldByName('StoreID').AsString;
    MailID    := TopMailsListTbl.FieldByName('EntryID').AsString;

    MailItem   := NS.GetItemFromID(MailID, OleVariant(StoreID));
    (MailItem as Outlook2010.MailItem).Delete;
    tvTopChange(tvTop, tvTop.Selected);
  Finally
    MailItem := nil;
    NS := nil;
  End;
end;

procedure TMainCompareFrm.btnGetBottomStoreIDClick(Sender: TObject);
Var
  MyTreeNode : TTreeNode;
  MyRecPtr   : PMyRec;
  FolderID   : String;
begin
  Try
    MyTreeNode := tvBottom.Selected;
    MyRecPtr   := MyTreeNode.Data;

    FolderID   := MyRecPtr.EnteryID;
    ShowMessage(FolderID);
  Finally
    MyRecPtr := nil;
    MyTreeNode := nil;
  End;
end;

procedure TMainCompareFrm.btnGetTopStoreIDClick(Sender: TObject);
Var
  MyTreeNode : TTreeNode;
  MyRecPtr   : PMyRec;
  FolderID   : String;
begin
  Try
    MyTreeNode := tvTop.Selected;
    MyRecPtr   := MyTreeNode.Data;

    FolderID   := MyRecPtr.EnteryID;
    ShowMessage(FolderID);
  Finally
    MyRecPtr := nil;
    MyTreeNode := nil;
  End;
end;

procedure TMainCompareFrm.btnTopCompareClick(Sender: TObject);
Var
  StoreName : String;
  StoreID_0 : String;
  StoreID_1 : String;
  EntryID   : String;
  SearchKey : String;
  RecNoT    : Integer;
  RecNoB    : Integer;
begin
  StoreID_0 := '';
  StoreID_1 := '';

  if not TopMailsListTbl.Active then
    Exit;
  if TopMailsListTbl.RecordCount = 0 then
    Exit;
  if not BottomMailsListTbl.Active then
    Exit;
  if BottomMailsListTbl.RecordCount = 0 then
    Exit;

  StoreID_0 := StoreIdList.Strings[cbTopStoreList.ItemIndex];
  StoreID_1 := StoreIdList.Strings[cbBottomStoreList.ItemIndex];

  if (StoreID_0 = StoreID_1) then
    raise System.SysUtils.Exception.Create('Please select different Store to compare');

  RecNoT := TopMailsListTbl.RecNo;
  RecNoB := BottomMailsListTbl.RecNo;
  Try
    LockWindowUpdate(Self.Handle);
    TopMailsListTbl.DisableControls;
    BottomMailsListTbl.DisableControls;

    TopMailsListTbl.First;
    While not TopMailsListTbl.EOF Do
    begin
      SearchKey := TopMailsListTbl.FieldByName('SearchKey').AsString;
      if not BottomMailsListTbl.Locate('SearchKey', SearchKey, []) then
      begin
        TopMailsListTbl.Edit;
        TopMailsListTbl.FieldByName('Error').AsBoolean := True;
        TopMailsListTbl.Post;
      end
      else
      begin
        TopMailsListTbl.Edit;
        TopMailsListTbl.FieldByName('Error').AsBoolean := False;
        TopMailsListTbl.Post;
      end;
      TopMailsListTbl.Next;
    end;
  Finally
    TopMailsListTbl.RecNo := RecNoT;
    BottomMailsListTbl.RecNo := RecNoB;
    BottomMailsListTbl.EnableControls;
    TopMailsListTbl.EnableControls;
    LockWindowUpdate(0);
  End;
end;

procedure TMainCompareFrm.AddChilde(TV : TTreeView; ParentNode : TTreeNode; FL : Outlook2010._Folders);
Var
  CurFolder  : Integer;
  MyTreeNode : TTreeNode;
  MyRecPtr   : PMyRec;
  FO         : Outlook2010.Folder;
begin
  for CurFolder := 1 to FL.Count do
  begin
    FO := FL.Item(CurFolder);

    if (FO.Class_ = olFolder) And
       ((FO.DefaultItemType = olMailItem) Or
        (FO.DefaultItemType = olPostItem)) Then
    begin
      New(MyRecPtr);
      MyRecPtr.EnteryID := FO.EntryID;
      with TV.Items do
      begin
        MyTreeNode := AddChildObject(ParentNode, FO.Name, MyRecPtr); { Add a root node. FO.EntryID;}
      end;

      if FO.Folders.Count > 0 then
      begin
        AddChilde(TV, MyTreeNode, FO.Folders);
      end;

    end;
  end;
end;

procedure TMainCompareFrm.BuildTreeList(StoreID : String; TV : TTreeView);
Var
  NS         : Outlook2010._NameSpace;
  olStores   : Outlook2010._Stores;
  olStore    : Outlook2010._Store;
  CurStore   : Integer;
  CurFolder  : Integer;
  rootFolder : MAPIFolder;
  FL         : Outlook2010._Folders;
  FO         : Outlook2010.Folder;
  MyTreeNode : TTreeNode;
  MyRecPtr   : PMyRec;
  CurCursor  : TCursor;
begin
  Try
    CurCursor  := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    TV.Items.Clear;
    LockWindowUpdate(Self.Handle);

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      olStores := NS.Stores;
      for CurStore := 1 to olStores.Count do
      begin
        olStore := olStores.Item(CurStore);
        if olStore.StoreID = StoreID then
        begin

          Try
            rootFolder := olStore.GetRootFolder;
            FL := rootFolder.Folders;
            for CurFolder := 1 to FL.Count do
            begin
              FO := FL.Item(CurFolder);

              if (FO.Class_ = olFolder) And
                 ((FO.DefaultItemType = olMailItem) Or
                  (FO.DefaultItemType = olPostItem)) Then
              begin
                New(MyRecPtr);
                MyRecPtr.EnteryID := FO.EntryID;
                MyTreeNode := TV.Items.AddObject(nil, FO.Name, MyRecPtr);

                if FO.Folders.Count > 0 then
                begin
                  AddChilde(TV, MyTreeNode, FO.Folders);
                end;
              end;
            end;
          Finally
          End;

          // stop Store Loop
          Break;
        end;
      end;
    Finally
      FO := nil;
      FL := nil;
      olStore := nil;
      olStores := nil;
      NS := nil;
    End;
  Finally
    Screen.Cursor  := CurCursor;
    LockWindowUpdate(0);
  End;
end;

end.
