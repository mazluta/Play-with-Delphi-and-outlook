unit uMain;


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
    MailsListTblHasAttach: TBooleanField;
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
    btnSaveAttachment: TButton;
    btnSaveMessage: TButton;
    btnSaveAsMHTML: TButton;
    FileSaveDialog: TFileSaveDialog;
    FileOpenDialogDir: TFileOpenDialog;
    btnSendMail: TButton;
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
    procedure btnSaveMessageClick(Sender: TObject);
    procedure btnSaveAsMHTMLClick(Sender: TObject);
    procedure btnSaveAttachmentClick(Sender: TObject);
    procedure btnSendMailClick(Sender: TObject);
  private
    { Private declarations }
    OnFirstTime   : Boolean;
    SwAvoidStoreScroll  : Boolean;
    SwAvoidFolderScroll : Boolean;
    OutlookApp    : Outlook2010.TOutlookApplication;
    procedure ReConnectOutlookApp;
  public
    { Public declarations }
  end;

var
  PlayWithOotlookFrm: TPlayWithOotlookFrm;

implementation
{$R *.dfm}

Uses
  IOUtils,
  uOutlookPR,
  uMailProps,
  SelectMapiFolder,
  forsix.MapiMail;

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

procedure TPlayWithOotlookFrm.btnSaveAsMHTMLClick(Sender: TObject);
Var
  FileName    : String;
  NS          : Outlook2010._NameSpace;
  MailEntryID : WideString;
  StoreID     : String;
  ItemToSave  : System.IDispatch;
begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount = 0 then
    Exit;

  FileSaveDialog.DefaultExtension := 'mhtml';
  if FileSaveDialog.Execute then
  begin
    FileName := FileSaveDialog.FileName;
    If UpperCase(ExtractFileExt(FileName)) <> UpperCase('.mhtml') Then
      FileName := FileName + '.mhtml';

    Try
      NS := OutlookApp.GetNameSpace('MAPI');
      StoreID  := MailsListTbl.FieldByName('StoreID').AsString;
      MailEntryID := MailsListTbl.FieldByName('EntryID').AsWideString;
      ItemToSave  := NS.GetItemFromID(MailEntryID, OleVariant(StoreID));
      if Assigned(ItemToSave) then
      begin
        (ItemToSave as MailItem).SaveAs(FileName,olMHTML);
      end;
    Finally
      ItemToSave := nil;
      NS := nil;
    End;
    ReConnectOutlookApp;
    ShowMessage('Mail Save As : ' + FileName);
  end;
end;

procedure TPlayWithOotlookFrm.btnSaveAttachmentClick(Sender: TObject);
Var
  SaveToFolder : String;
  FileName     : String;
  FullFileName : String;
  NS           : Outlook2010._NameSpace;
  MailEntryID  : WideString;
  StoreID      : String;
  ItemToSave   : System.IDispatch;
  CurItem      : Integer;

  function  RemoveLastEnterChar(SrcString : WideString) : WideString;
  begin
    // Remove The Last $D$A From SrcString;

    Result := SrcString;
    While True Do
    begin
      IF (Copy(Result,Length(Result),1) = #13) Or
         (Copy(Result,Length(Result),1) = #10) Then
      begin
        Result := Copy(Result,1,Length(Result)-1);
      end
      else
        Break;
    end;
  end;

  function  RemoveBackSlashChar(SrcString : String) : String;
  begin
    // Remove The Last '\' From SrcString;
    Result := RemoveLastEnterChar(SrcString);
    IF Copy(SrcString,Length(SrcString),1) = '\' Then
      Result := Copy(Result,1,Length(Result)-1);
  end;

begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount = 0 then
    Exit;

  if not MailsListTbl.FieldByName('HasAttach').AsBoolean Then
    raise System.SysUtils.Exception.Create('This mail has no attachtments');

  SaveToFolder := '';
  if FileOpenDialogDir.Execute then
  begin
    SaveToFolder := FileOpenDialogDir.FileName;
    Try
      NS          := OutlookApp.GetNameSpace('MAPI');
      StoreID     := MailsListTbl.FieldByName('StoreID').AsString;
      MailEntryID := MailsListTbl.FieldByName('EntryID').AsWideString;
      ItemToSave  := NS.GetItemFromID(MailEntryID, OleVariant(StoreID));
      if Assigned(ItemToSave) then
      begin
        For CurItem := 1 to (ItemToSave as MailItem).Attachments.Count do
        begin
          FileName     := (ItemToSave as Outlook2010.MailItem).Attachments.Item(CurItem).FileName;
          FullFileName := RemoveBackSlashChar(SaveToFolder) + '\' + IntToStr(CurItem) + '_' + FileName;
          (ItemToSave as Outlook2010.MailItem).Attachments.Item(CurItem).SaveAsFile(FullFileName);
        end;
      end;
    Finally
      ItemToSave := nil;
      NS := nil;
    End;

    ShowMessage('Attachment have been saved to Folder : ' + SaveToFolder);
  end;
end;

procedure TPlayWithOotlookFrm.btnSaveMessageClick(Sender: TObject);
Var
  FileName    : String;
  NS          : Outlook2010._NameSpace;
  MailEntryID : WideString;
  StoreID     : String;
  ItemToSave  : System.IDispatch;
begin
  if not MailsListTbl.Active then
    Exit;
  if MailsListTbl.RecordCount = 0 then
    Exit;

  FileSaveDialog.DefaultExtension := 'msg';
  if FileSaveDialog.Execute then
  begin
    FileName := FileSaveDialog.FileName;
    If UpperCase(ExtractFileExt(FileName)) <> UpperCase('.msg') Then
      FileName := FileName + '.Msg';

    Try
      NS := OutlookApp.GetNameSpace('MAPI');
      StoreID  := MailsListTbl.FieldByName('StoreID').AsString;
      MailEntryID := MailsListTbl.FieldByName('EntryID').AsWideString;
      ItemToSave  := NS.GetItemFromID(MailEntryID, OleVariant(StoreID));
      if Assigned(ItemToSave) then
      begin
        (ItemToSave as MailItem).SaveAs(FileName,olMSGUnicode);
      end;
    Finally
      ItemToSave := nil;
      NS := nil;
    End;
    ReConnectOutlookApp;
    ShowMessage('Mail Save As : ' + FileName);
  end;
end;

procedure TPlayWithOotlookFrm.btnSendMailClick(Sender: TObject);
Var
  RL : TStringList;
  AL : TStringList;
  SubjectStr : String;
  BodyStr    : String;

  function SendMessageViaOLE(Subject    : WideString;
                             Body       : WideString;
                             UseHtml    : Boolean = False;
                             Recipients : TStringList = nil;
                             AttachList : TStringList = nil) : Integer;
  var
    CurItem   : Integer;
    MI        : System.IDispatch;
  begin
    Result := 0;
    Try
      Try
        MI := OutlookApp.CreateItem(olMailItem) ;
      Except
        on e : System.SysUtils.exception Do
        begin
          ShowMessage(e.Message);
        end;
      End;

      //MailItem.Recipients.Add('johndoe@hotmail.com') ;
      if Trim(Subject) <> '' then
        (MI as Outlook2010.MailItem).Subject := Subject;

      if Trim(Body) <> '' then
      begin
        if UseHtml then
          (MI as Outlook2010.MailItem).HTMLBody := Body
        else
          (MI as Outlook2010.MailItem).Body := Body;
      end;

      if Assigned(recipients) then
      begin
        for CurItem := 0 to Recipients.Count - 1 do
        begin
          if Trim(Recipients.Strings[CurItem]) <> '' then
            (MI as Outlook2010.MailItem).Recipients.Add(Recipients.Strings[CurItem]);
          // wait somw more for prevent - Call was rejected by callee.
          Sleep(110);
          Application.ProcessMessages;
        end;
      end;

      if Assigned(AttachList) then
      begin
        for CurItem := 0 to AttachList.Count - 1 do
        begin
          if (Trim(AttachList.Strings[CurItem]) <> '') and
             FileExists(AttachList.Strings[CurItem]) then
           (MI as Outlook2010.MailItem).Attachments.Add(OleVariant(AttachList.Strings[CurItem]),
                                                        OleVariant(1),
                                                        OleVariant(2),
                                                        OleVariant(ExtractFileName(AttachList.Strings[CurItem])) );
          // wait somw more for prevent - Call was rejected by callee.
          Sleep(110);
          Application.ProcessMessages;
        end;
      end;

      //this will show MailItem Outlook Dialog
      (MI as Outlook2010.MailItem).Display(True);

      //if SendMai then - this will be new parameter
      //begin
      //if you want to auto SEND then you have to create OUTLOOK_TLB from 2013 +
      //(MI as Outlook2010.MailItem).Send;
      //end
      //else
      //begin
      // (MI as Outlook2010.MailItem).Display(True);
      //end;
    Finally
      Try MI := nil; Except; End;
    End;
  end;


begin
  Try
    SubjectStr := 'Empty Subject';

    BodyStr := '<HTML>' +
               '<head><style> ' +
                 'body {color: blue;} ' +
                 'h1 {color: green;} ' +
                 'p {color: red;} ' +
               '</style></head> '+
               '<BODY> <h1>Test Send mail</h1> Body String <p>Paragraph string</p></BODY><HTML>';

    RL := TStringList.Create;
    RL.Add('mazluta@hanibaal.co.il');

    AL := TStringList.Create;
    AL.Add('c:\pics\Anny\47101.Jpg');

    SendMessageViaOLE(SubjectStr,
                      BodyStr,
                      True, {UseHtml}
                      RL,
                      AL);
  Finally
    FreeAndNil(RL);
    FreeAndNil(AL);
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
  //olPA     : _PropertyAccessor;
  StoreID  : String;
  FolderID : String;
  FO       : MAPIFolder;
  foItems  : Outlook2010.Items;
  foItem   : Outlook2010.MailItem;
  CurItem  : Integer;
  CurCursor  : TCursor;
  fHasAttach : Boolean;

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
            //olPA := foItem.PropertyAccessor;
            //PR_HASATTACH  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1B000B'; {PT_BOOLEAN}
            fHasAttach := foItem.PropertyAccessor.GetProperty(PR_HASATTACH);
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
            //MailsListTbl.FieldByName('HasAttach').AsBoolean   := (foItem.Attachments.Count > 0);
            MailsListTbl.FieldByName('HasAttach').AsBoolean   := fHasAttach;
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

procedure TPlayWithOotlookFrm.ReConnectOutlookApp;
begin
  // this proc will prevent error : The RPC server is unavailable.

  OutlookApp.Disconnect;
  OutlookApp.Quit;
  FreeAndNil(OutlookApp);

  OutlookApp := TOutlookApplication.Create(Nil);
  OutlookApp.ConnectKind := ckRunningOrNew;
  OutlookApp.Connect;
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

