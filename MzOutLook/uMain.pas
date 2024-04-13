﻿unit uMain;


//****************************************************************************
//***** Example of Using Extended MAPI
//  foItems := FO.Items;
//  Try
//    // to find mail item by it's EntryID we need to use Extended Mapi
//    // so - we loop for all mail items with Subject XX and check for its EntryID
//    foItem := foItems.Find('[Subject]="' + Subject + '"') as MailItem;
//    if foItem = nil then
//      // if we get back the item as nil then we will look for sender email (always in Latin)
//     foItem := foItems.Find('[SenderEmailAddress]="' + FromEmail + '"') as MailItem;
//  Except;
//    foItem := nil;
//  End;
//****************************************************************************

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, forms, Dialogs,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Vcl.Controls, Vcl.ExtCtrls,
  Data.DB, Datasnap.DBClient,
  Outlook2010, ActiveX, ComObj, ole2, OleServers, Vcl.OleServer;

type

  TPlayWithOotlookFrm = class(TForm)
    Panel1: TPanel;
    sbGetProps: TButton;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    MailsListGrid: TDBGrid;
    FoldersListGrid: TDBGrid;
    StoresListsGrid: TDBGrid;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    FoldersListDS: TDataSource;
    FoldersListTbl: TClientDataSet;
    FoldersListTblNumber: TIntegerField;
    FoldersListTblFolderName: TWideStringField;
    FoldersListTblEntryId: TWideStringField;
    FoldersListTblStoreID: TWideStringField;
    FoldersListTblNewname: TWideStringField;
    MailsListDS: TDataSource;
    MailsListTbl: TClientDataSet;
    MailsListTblNumber: TIntegerField;
    MailsListTblReciveDate: TDateTimeField;
    MailsListTblSubject: TWideStringField;
    MailsListTblFromName: TWideStringField;
    MailsListTblFromEmail: TWideStringField;
    MailsListTblCC: TWideStringField;
    MailsListTblBCC: TWideStringField;
    MailsListTblStoreID: TWideStringField;
    MailsListTblFolderID: TWideStringField;
    MailsListTblEntryID: TWideStringField;
    MailsListTblSearchKey: TWideStringField;
    StoresListDS: TDataSource;
    StoresListTbl: TClientDataSet;
    StoresListTblNumber: TIntegerField;
    StoresListTblStoreName: TWideStringField;
    StoresListTblStoreID: TWideStringField;
    StoresListTblFullName: TWideStringField;
    Label1: TLabel;
    Panel6: TPanel;
    btnAddMsgStore: TButton;
    btnRemoveMsgStore: TButton;
    FileOpenDialog: TFileOpenDialog;
    btnMoveMailItemToOtherFolder: TButton;
    btnDeleteSelectedMailItem: TButton;
    sbBuildStoresList: TButton;
    procedure sbBuildStoresListClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure StoresListTblAfterScroll(DataSet: TDataSet);
    procedure FoldersListTblAfterScroll(DataSet: TDataSet);
    procedure sbGetPropsClick(Sender: TObject);
    procedure MailsListGridTitleClick(Column: TColumn);
    procedure MailsListGridDblClick(Sender: TObject);
    procedure btnAddMsgStoreClick(Sender: TObject);
    procedure btnRemoveMsgStoreClick(Sender: TObject);
    procedure btnMoveMailItemToOtherFolderClick(Sender: TObject);
    procedure btnDeleteSelectedMailItemClick(Sender: TObject);
  private
    { Private declarations }
    OnFirstTime   : Boolean;
    SwAvoidStoreScroll  : Boolean;
    SwAvoidFolderScroll : Boolean;
    OutlookApp    : Outlook2010.TOutlookApplication;
  public
    { Public declarations }
  end;

var
  PlayWithOotlookFrm: TPlayWithOotlookFrm;

implementation
{$R *.dfm}

Uses
  uOutlookPR,
  uMailProps,
  SelectMapiFolder;

procedure TPlayWithOotlookFrm.btnAddMsgStoreClick(Sender: TObject);
Var
  PstFile  : String;
  NS       : Outlook2010._NameSpace;
begin
  if FileOpenDialog.Execute then
  begin
    PstFile := FileOpenDialog.FileName;
    if UpperCase(ExtractFileExt(PstFile)) = UpperCase('.pst') then
    begin
      NS := OutlookApp.GetNameSpace('MAPI');
      NS.AddStore(OleVariant(PstFile));
      sbBuildStoresList.Click;
    end;
  end;
end;

procedure TPlayWithOotlookFrm.btnRemoveMsgStoreClick(Sender: TObject);
Var
  StoreName : String;
  StoreFile : String;
  StoreID   : String;
  NS        : Outlook2010._NameSpace;
  olStores  : Outlook2010._Stores;
  olStore   : Outlook2010._Store;
  FO        : MAPIFolder;
  CurStore  : Integer;
begin
  if not StoresListTbl.Active then
    exit;
  if StoresListTbl.RecordCount = 0 then
    exit;

  StoreName := StoresListTbl.FieldByName('StoreName').AsString;
  StoreFile := StoresListTbl.FieldByName('FullName').AsString;
  StoreID   := StoresListTbl.FieldByName('StoreID').AsString;

  if MessageDlg('Remove Msg Store - [' + StoreName + '] from outlook mapi?',
        mtConfirmation,[mbYes,mbNo],0, mbNo) = mrNo then
    exit;

  Try
    NS := OutlookApp.GetNameSpace('MAPI');
    olStores := NS.Stores;
    for CurStore := 1 to olStores.Count do
    begin
      olStore := olStores.Item(CurStore);
      if olStore.StoreID = StoreID Then
      begin
        FO := olStore.GetRootFolder;
        NS.RemoveStore(FO);

        sbBuildStoresList.Click;
        Break;
      end;
    end;
  Finally
    FO := nil;
    olStore := nil;
    olStores := nil;
    NS := nil;
  End;
end;

procedure TPlayWithOotlookFrm.btnMoveMailItemToOtherFolderClick(Sender: TObject);
Var
  NS        : Outlook2010._NameSpace;
  MailEntryID : WideString;
  StoreID     : String;
  SelectedID  : String;
  FolderDst   : System.IDispatch;
  ItemToMove  : System.IDispatch;
begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount = 0 then
    Exit;

  // 1. find the folder you want to move the MailItem to.
  // 2. find the MailItem as Outlook2010.MailItem object
  // 3. move the item to that folder

  SelectedID := '';
  Try
    SelectMapiFolderFrm := TSelectMapiFolderFrm.Create(Application);
    if SelectMapiFolderFrm.ShowModal = mrOk Then
      SelectedID := SelectMapiFolderFrm.SelectedEnteryID;
  Finally
    FreeAndNil(SelectMapiFolderFrm);
  End;

  if SelectedID = '' Then
    exit;

  Try
    NS := OutlookApp.GetNameSpace('MAPI');
    StoreID  := MailsListTbl.FieldByName('StoreID').AsString;
    FolderDst := NS.GetFolderFromID(SelectedID, OleVariant(StoreID));
    if Assigned(FolderDst) then
    begin
      MailEntryID := MailsListTbl.FieldByName('EntryID').AsWideString;
      ItemToMove  := NS.GetItemFromID(MailEntryID, OleVariant(StoreID));
      if Assigned(ItemToMove) then
      begin
        (ItemToMove as MailItem).Move(FolderDst as Outlook2010.MAPIFolder);
      end;
    end;
  Finally
    ItemToMove := nil;
    FolderDst := nil;
    NS := nil;
  End;

  FoldersListTblAfterScroll(FoldersListTbl);
end;

procedure TPlayWithOotlookFrm.btnDeleteSelectedMailItemClick(Sender: TObject);
Var
  NS         : Outlook2010._NameSpace;
  ItemToDel  : System.IDispatch;
  ItemID     : String;
  StoreID    : String;
  FolderID   : String;
  EntryID    : WideString;
begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount = 0 then
    Exit;

  ItemID   := MailsListTbl.FieldByName('Number').AsString;
  StoreID  := MailsListTbl.FieldByName('StoreID').AsString;
  FolderID := MailsListTbl.FieldByName('FolderID').AsString;
  EntryID  := MailsListTbl.FieldByName('EntryID').AsString;

  if MessageDlg('Ok To Delete Msg No# - [' + ItemID + '] ?',
        mtConfirmation,[mbYes,mbNo],0, mbNo) = mrNo then
    exit;

  Try
    NS := OutlookApp.GetNameSpace('MAPI');
    ItemToDel   := NS.GetItemFromID(EntryID, OleVariant(StoreID));
    if Assigned(ItemToDel) then
    begin
      (ItemToDel as MailItem).Delete;
    end;
  Finally
    ItemToDel := nil;
    NS := nil;
  End;

  FoldersListTblAfterScroll(FoldersListTbl);
end;

procedure TPlayWithOotlookFrm.FoldersListTblAfterScroll(DataSet: TDataSet);
Var
  NS       : Outlook2010._NameSpace;
  StoreID  : String;
  FolderID : String;
  FO       : MAPIFolder;
  foItems  : Outlook2010.Items;
  foItem   : Outlook2010.MailItem;
  CurItem  : Integer;
  CurCursor  : TCursor;

begin
  If SwAvoidFolderScroll Then
    Exit;

  Try
    CurCursor  := Screen.Cursor;
    Screen.Cursor := crHourGlass;

    MailsListTbl.Close;
    MailsListTbl.CreateDataSet;
    MailsListTbl.Open;
    MailsListTbl.EmptyDataSet;

    StoreID := DataSet.FieldByName('StoreID').AsString;
    FolderID := DataSet.FieldByName('EntryID').AsString;

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      FO := NS.GetFolderFromID(FolderID, OleVariant(StoreID));
      foItems := FO.Items;
      Try
        foItems.Sort('ReceivedTime',True);
      Except;
      End;

      Try
        MailsListTbl.DisableControls;
        For CurItem := 1 to foItems.Count Do
        begin
          if Supports(foItems.Item(CurItem), MailItem, foItem) then
          begin
            MailsListTbl.Insert;
            MailsListTbl.FieldByName('Number').AsInteger      := CurItem;
            Try
              MailsListTbl.FieldByName('ReciveDate').AsDateTime := foItem.ReceivedTime;
            Except;
              MailsListTbl.FieldByName('ReciveDate').AsDateTime := NULL;
            End;
            MailsListTbl.FieldByName('Subject').AsString      := foItem.Subject;
            MailsListTbl.FieldByName('FromName').AsString     := foItem.SenderName;
            MailsListTbl.FieldByName('FromEmail').AsString    := foItem.SenderEmailAddress;
            MailsListTbl.FieldByName('CC').AsString           := foItem.CC;
            MailsListTbl.FieldByName('BCC').AsString          := foItem.BCC;
            MailsListTbl.FieldByName('StoreID').AsString      := StoreID;
            MailsListTbl.FieldByName('FolderID').AsString     := FolderID;
            MailsListTbl.FieldByName('EntryID').AsString      := foItem.EntryID;
            MailsListTbl.Post;
          end;
        end;

        if MailsListTbl.RecordCount > 0 Then
        begin
          MailsListTbl.IndexName := 'ByNumber';
          MailsListTbl.First;
        end;
      Finally
        MailsListTbl.EnableControls;
      End;
    Finally
      foItem := nil;
      foItems := nil;
      FO := nil;
      NS := nil;
    End;
  Finally
    Screen.Cursor  := CurCursor;
  End;
end;

procedure TPlayWithOotlookFrm.FormActivate(Sender: TObject);
begin
  if OnFirstTime Then
  Try
    OutlookApp := TOutlookApplication.Create(Nil);
    OutlookApp.ConnectKind := ckRunningOrNew;
    OutlookApp.Connect;

  Finally
    OnFirstTime := False;
  End;
end;

procedure TPlayWithOotlookFrm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  OutlookApp.Disconnect;
  OutlookApp.Quit;
  FreeAndNil(OutlookApp);
end;

procedure TPlayWithOotlookFrm.FormCreate(Sender: TObject);
begin
  OnFirstTime   := True;
  SwAvoidStoreScroll  := False;
  SwAvoidFolderScroll := False;

  StoresListTbl.Close;
  FoldersListTbl.Close;
  MailsListTbl.Close;
end;

procedure TPlayWithOotlookFrm.MailsListGridDblClick(Sender: TObject);
begin
  if MailsListTbl.Active And
     (MailsListTbl.RecordCount > 0) Then
    sbGetProps.Click;
end;

procedure TPlayWithOotlookFrm.MailsListGridTitleClick(Column: TColumn);
begin
  if Column.FieldName = 'Number' then
  begin
    MailsListTbl.IndexName := 'ByNumber';
  end
  else
  if Column.FieldName = 'Subject' then
  begin
    MailsListTbl.IndexName := 'BySubject';
  end
  else
  if Column.FieldName = 'ReciveDate' then
  begin
    MailsListTbl.IndexName := 'ByReciveDate';
  end
  else
  if Column.FieldName = 'FromEmail' then
  begin
    MailsListTbl.IndexName := 'ByFromEmail';
  end
  else
  if Column.FieldName = 'FromName' then
  begin
    MailsListTbl.IndexName := 'ByFromName';
  end;
end;

procedure TPlayWithOotlookFrm.sbBuildStoresListClick(Sender: TObject);
Var
  NS       : Outlook2010._NameSpace;
  olStores : Outlook2010._Stores;
  olStore  : Outlook2010._Store;
  CurStore : Integer;
begin
  StoresListTbl.Close;
  StoresListTbl.CreateDataSet;
  StoresListTbl.Open;
  StoresListTbl.EmptyDataSet;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    StoresListTbl.DisableControls;
    SwAvoidStoreScroll := True; //this will avoid build folder list
    Try
      olStores := NS.Stores;
      for CurStore := 1 to olStores.Count do
      begin
        olStore := olStores.Item(CurStore);

        StoresListTbl.Insert;
        StoresListTbl.Open;
        StoresListTbl.FieldByName('Number').AsInteger   := CurStore;
        StoresListTbl.FieldByName('StoreName').AsString := olStore.DisplayName;
        StoresListTbl.FieldByName('StoreID').AsString   := olStore.StoreID;
        StoresListTbl.FieldByName('FullName').AsString  := olStore.FilePath;
        StoresListTbl.Post;
      end;
    Finally
      SwAvoidStoreScroll := False;
    End;
    StoresListTbl.First; // now the folder list will be build
  Finally
    olStore := nil;
    olStores := nil;
    NS := nil;
    StoresListTbl.EnableControls;
  End;
end;

procedure TPlayWithOotlookFrm.sbGetPropsClick(Sender: TObject);
Var
  NS        : Outlook2010._NameSpace;
  FO        : Outlook2010.MAPIFolder;
  ItemProp  : System.IDispatch;
  StoreID   : String;
  FolderID  : String;
  EntryID   : String;
begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount < 1 then
    Exit;

  StoreID   := MailsListTbl.FieldByName('StoreID').AsString;
  FolderID  := MailsListTbl.FieldByName('FolderID').AsString;
  EntryID   := MailsListTbl.FieldByName('EntryID').AsString;

  NS := OutlookApp.GetNameSpace('MAPI');
  Try
    FO := NS.GetFolderFromID(FolderID, OleVariant(StoreID));
    ItemProp  := NS.GetItemFromID(EntryID, OleVariant(StoreID));

    if Assigned(ItemProp) Then
    begin
      Try
        MailPropsFrm := TMailPropsFrm.Create(Application);
        MailPropsFrm.MI := (ItemProp as MailItem);
        MailPropsFrm.ShowModal;
      Finally
        FreeAndNil(MailPropsFrm);
      End;
    End
    else
      ShowMessage('Mail Item not found');
  Finally
    ItemProp := nil;
    FO := nil;
    NS := nil;
  End;
end;

procedure TPlayWithOotlookFrm.StoresListTblAfterScroll(DataSet: TDataSet);
Var
  StoreID    : String;
  NS         : Outlook2010._NameSpace;
  olStores   : Outlook2010._Stores;
  olStore    : Outlook2010._Store;
  CurStore   : Integer;
  CurFolder  : Integer;
  rootFolder : MAPIFolder;
  FL         : Outlook2010._Folders;
  FO         : Outlook2010.Folder;
  CurCursor  : TCursor;
begin
  If SwAvoidStoreScroll Then
    Exit;

  StoreID := DataSet.FieldByName('StoreID').AsString;

  Try
    CurCursor  := Screen.Cursor;
    Screen.Cursor := crHourGlass;
    FoldersListTbl.Close;
    FoldersListTbl.CreateDataSet;
    FoldersListTbl.Open;
    FoldersListTbl.EmptyDataSet;

    NS := OutlookApp.GetNameSpace('MAPI');
    Try
      FoldersListTbl.DisableControls;
      olStores := NS.Stores;
      for CurStore := 1 to olStores.Count do
      begin
        olStore := olStores.Item(CurStore);
        if olStore.StoreID = StoreID then
        begin
          Try
            SwAvoidFolderScroll := True;
            rootFolder := olStore.GetRootFolder;
            FL := rootFolder.Folders;
            for CurFolder := 1 to FL.Count do
            begin
              FO := FL.Item(CurFolder);

              if (FO.Class_ = olFolder) And
                 ((FO.DefaultItemType = olMailItem) Or
                  (FO.DefaultItemType = olPostItem)) Then
              begin
                FoldersListTbl.Insert;
                FoldersListTbl.Open;
                FoldersListTbl.FieldByName('Number').AsInteger    := CurFolder;
                FoldersListTbl.FieldByName('FolderName').AsString := FO.Name;
                FoldersListTbl.FieldByName('EntryID').AsString    := FO.EntryID;
                FoldersListTbl.FieldByName('StoreID').AsString    := StoreID;
                FoldersListTbl.Post;
              end;
            end;
          Finally
            SwAvoidFolderScroll := False;
          End;

          //Go to first folder
          if FoldersListTbl.RecordCount > 0 Then
            FoldersListTbl.First;

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
      FoldersListTbl.EnableControls;
    End;
  Finally
    Screen.Cursor  := CurCursor;
  End;
end;

initialization
  CoInitialize(nil);

finalization
  CoUnInitialize;

end.
