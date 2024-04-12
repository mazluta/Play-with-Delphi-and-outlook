unit uMailProps;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Outlook2010, ActiveX, ComObj, ole2, OleServers, Vcl.OleServer, Data.DB,
  Datasnap.DBClient, Vcl.Grids, Vcl.DBGrids;

type
  TMailPropsFrm = class(TForm)
    Panel1: TPanel;
    btnViewAsPlainText: TButton;
    btnViewAsHtmlBody: TButton;
    btnViewAsRTFbody: TButton;
    btnSaveMessage: TButton;
    btnSaveAsMHTML: TButton;
    MailPropDS: TDataSource;
    MailPropTbl: TClientDataSet;
    MailPropTblNumber: TIntegerField;
    MailPropTblPropName: TWideStringField;
    MailPropTblPropType: TWideStringField;
    MailPropTblPropValue: TWideStringField;
    MailPropTblPropValueW: TWideStringField;
    StoresListsGrid: TDBGrid;
    FileOpenDialog: TFileOpenDialog;
    FileSaveDialog: TFileSaveDialog;
    btnViewAsMhtml: TButton;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSaveMessageClick(Sender: TObject);
    procedure btnSaveAsMHTMLClick(Sender: TObject);
    procedure btnViewAsPlainTextClick(Sender: TObject);
    procedure btnViewAsHtmlBodyClick(Sender: TObject);
    procedure btnViewAsRTFbodyClick(Sender: TObject);
    procedure btnViewAsMhtmlClick(Sender: TObject);
  private
    { Private declarations }
    OnFirstTime : Boolean;
    procedure BuildPropList;
  public
    { Public declarations }
    MI : Outlook2010.MailItem;
  end;

var
  MailPropsFrm: TMailPropsFrm;

implementation
{$R *.dfm}

Uses
  uOutlookPR,
  uMaileMsgViewer;

procedure TMailPropsFrm.btnViewAsHtmlBodyClick(Sender: TObject);
begin
  Try
    MaileMsgViewerFrm := TMaileMsgViewerFrm.Create(Application);
    MaileMsgViewerFrm.MI := MI;
    MaileMsgViewerFrm.ViewerType := olViewerTypeHtml;
    MaileMsgViewerFrm.ShowModal;
  Finally
    FreeAndNil(MaileMsgViewerFrm);
  End;
end;

procedure TMailPropsFrm.btnViewAsMhtmlClick(Sender: TObject);
begin
  Try
    MaileMsgViewerFrm := TMaileMsgViewerFrm.Create(Application);
    MaileMsgViewerFrm.MI := MI;
    MaileMsgViewerFrm.ViewerType := olViewerTypeMhtml;
    MaileMsgViewerFrm.ShowModal;
  Finally
    FreeAndNil(MaileMsgViewerFrm);
  End;
end;

procedure TMailPropsFrm.btnViewAsPlainTextClick(Sender: TObject);
begin
  Try
    MaileMsgViewerFrm := TMaileMsgViewerFrm.Create(Application);
    MaileMsgViewerFrm.MI := MI;
    MaileMsgViewerFrm.ViewerType := olViewerTypePlainText;
    MaileMsgViewerFrm.ShowModal;
  Finally
    FreeAndNil(MaileMsgViewerFrm);
  End;
end;

procedure TMailPropsFrm.btnViewAsRTFbodyClick(Sender: TObject);
begin
  Try
    MaileMsgViewerFrm := TMaileMsgViewerFrm.Create(Application);
    MaileMsgViewerFrm.MI := MI;
    MaileMsgViewerFrm.ViewerType := olViewerTypeRtf;
    MaileMsgViewerFrm.ShowModal;
  Finally
    FreeAndNil(MaileMsgViewerFrm);
  End;
end;

procedure TMailPropsFrm.btnSaveAsMHTMLClick(Sender: TObject);
Var
  FileName : String;
begin
  FileSaveDialog.DefaultExtension := 'mhtml';
  if FileSaveDialog.Execute then
  begin
    FileName := FileSaveDialog.FileName;
    If UpperCase(ExtractFileExt(FileName)) <> UpperCase('.mhtml') Then
      FileName := FileName + '.mhtml';

    MI.SaveAs(FileName, olMHTML);
    ShowMessage('Mail Save As : ' + FileName);
  end;
end;

procedure TMailPropsFrm.btnSaveMessageClick(Sender: TObject);
Var
  FileName : String;
begin
  FileSaveDialog.DefaultExtension := 'msg';
  if FileSaveDialog.Execute then
  begin
    FileName := FileSaveDialog.FileName;
    If UpperCase(ExtractFileExt(FileName)) <> UpperCase('.msg') Then
      FileName := FileName + '.Msg';

    MI.SaveAs(FileName,olMSGUnicode);
    ShowMessage('Mail Save As : ' + FileName);
  end;
end;

procedure TMailPropsFrm.FormActivate(Sender: TObject);
begin
  If OnFirstTime Then
  Try
    BuildPropList;
  Finally
    OnFirstTime := False;
  End;
end;

procedure TMailPropsFrm.FormCreate(Sender: TObject);
begin
  OnFirstTime := True;
end;

procedure TMailPropsFrm.BuildPropList;
Var
  olPA : _PropertyAccessor;
  StrA : String;
  StrW : WideString;
  CurProp : Integer;

  function GetBestFormatType(BestFormatType : Integer) : String;
  begin
    case BestFormatType of
      0 : Result := 'olUnspecified';
      1 : Result := 'olPlainText';
      2 : Result := 'olRTFCompressed';
      3 : Result := 'olHtmlBody';
      4 : Result := 'olClearSigned';
    else
      Result := 'UnKnown';
    end;
  end;

  function GetFormatType(FormatType : Integer) : String;
  begin
    case FormatType of
      0 : Result := 'olFormatUnspecified';
      1 : Result := 'olFormatPlain';
      2 : Result := 'olFormatHTML';
      3 : Result := 'olFormatRichText';
    else
      Result := 'UnKnown';
    end;
  end;

  function GetObjectType(ObjType : Integer) : String;
  begin
    case ObjType of
      1 : Result := 'Message Store';
      2 : Result := 'Address Book';
      3 : Result := 'Folder';
      4 : Result := 'Address Book Container';
      5 : Result := 'Message';
      6 : Result := 'Individual Recipient';
      7 : Result := 'Attachment';
      8 : Result := 'Distribution List Recipient';
      9 : Result := 'Profile Section';
      10 : Result := 'Status Object';
      11 : Result := 'Session';
      12 : Result := 'Form Information';
    else
      Result := 'UnKnown';
    end;
  end;

begin
  MailPropTbl.Close;
  MailPropTbl.CreateDataSet;
  MailPropTbl.Open;
  MailPropTbl.EmptyDataSet;
  CurProp := 0;

  olPA := MI.PropertyAccessor;
  Try
    MailPropTbl.DisableControls;

    //PR_SUBJECT : String = 'http://schemas.microsoft.com/mapi/proptag/0x0037001E'; {PT_TSTRING}
    StrA := Trim(MI.PropertyAccessor.GetProperty(PR_SUBJECT));
    StrW := Trim(MI.PropertyAccessor.GetProperty(PR_SUBJECT_W));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_SUBJECT';
    MailPropTbl.FieldByName('PropType').AsString := 'PR_STRING';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.FieldByName('PropValueW').AsString := StrW;
    MailPropTbl.Post;

    //PR_NORMALIZED_SUBJECT : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1D001E'; {PT_TSTRING}
    StrA := Trim(MI.PropertyAccessor.GetProperty(PR_NORMALIZED_SUBJECT));
    StrW := Trim(MI.PropertyAccessor.GetProperty(PR_NORMALIZED_SUBJECT_W));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_NORMALIZED_SUBJECT';
    MailPropTbl.FieldByName('PropType').AsString := 'PR_STRING';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.FieldByName('PropValueW').AsString := StrW;
    MailPropTbl.Post;

    //PR_MESSAGE_CLASS : String = 'http://schemas.microsoft.com/mapi/proptag/0x001A001E'; {PT_TSTRING}
    StrA := Trim(MI.PropertyAccessor.GetProperty(PR_MESSAGE_CLASS));
    StrW := Trim(MI.PropertyAccessor.GetProperty(PR_MESSAGE_CLASS_W));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_MESSAGE_CLASS';
    MailPropTbl.FieldByName('PropType').AsString := 'PR_STRING';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.FieldByName('PropValueW').AsString := StrW;
    MailPropTbl.Post;

    //PR_SENDER_NAME : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1A001E'; {PT_TSTRING}
    StrA := Trim(MI.PropertyAccessor.GetProperty(PR_SENDER_NAME));
    StrW := Trim(MI.PropertyAccessor.GetProperty(PR_SENDER_NAME_W));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_SENDER_NAME';
    MailPropTbl.FieldByName('PropType').AsString := 'PR_STRING';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.FieldByName('PropValueW').AsString := StrW;
    MailPropTbl.Post;

    //PR_SENDER_EMAIL_ADDRESS : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1F001E'; {PT_TSTRING}
    StrA := Trim(MI.PropertyAccessor.GetProperty(PR_SENDER_EMAIL_ADDRESS));
    StrW := Trim(MI.PropertyAccessor.GetProperty(PR_SENDER_EMAIL_ADDRESS_W));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_SENDER_EMAIL_ADDRESS';
    MailPropTbl.FieldByName('PropType').AsString := 'PR_STRING';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.FieldByName('PropValueW').AsString := StrW;
    MailPropTbl.Post;

    //PR_CREATION_TIME : String = 'http://schemas.microsoft.com/mapi/proptag/0x30070040'; {PT_SYSTIME}
    StrA := DateTimeToStr(MI.PropertyAccessor.GetProperty(PR_CREATION_TIME));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_CREATION_TIME';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_SYSTIME';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_MESSAGE_DELIVERY_TIME : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'; {PT_SYSTIME}
    StrA := DateTimeToStr(MI.PropertyAccessor.GetProperty(PR_MESSAGE_DELIVERY_TIME));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_MESSAGE_DELIVERY_TIME';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_SYSTIME';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_SEARCH_KEY : String = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102'; {PT_BINARY}
    StrA := olPA.BinaryToString(MI.PropertyAccessor.GetProperty(PR_SEARCH_KEY));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_SEARCH_KEY';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BINARY';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_ENTRYID : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFF0102'; {PT_BINARY}
    StrA := olPA.BinaryToString(MI.PropertyAccessor.GetProperty(PR_ENTRYID));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_ENTRYID';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BINARY';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_PARENT_ENTRYID : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E090102'; {PT_BINARY}
    StrA := olPA.BinaryToString(MI.PropertyAccessor.GetProperty(PR_PARENT_ENTRYID));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_PARENT_ENTRYID';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BINARY';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_STORE_ENTRYID : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFB0102'; {PT_BINARY}
    StrA := olPA.BinaryToString(MI.PropertyAccessor.GetProperty(PR_STORE_ENTRYID));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_STORE_ENTRYID';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BINARY';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_HASATTACH  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1B000B'; {PT_BOOLEAN}
    StrA := BoolToStr(MI.PropertyAccessor.GetProperty(PR_HASATTACH),True);
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_HASATTACH';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BOOLEAN';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_RTF_IN_SYNC : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1F000B'; {PT_BOOLEAN}
    StrA := BoolToStr(MI.PropertyAccessor.GetProperty(PR_RTF_IN_SYNC),True);
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_RTF_IN_SYNC';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_BOOLEAN';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_MSG_EDITOR_FORMAT : String = 'http://schemas.microsoft.com/mapi/proptag/0x59090003'; {PT_LONG}
    StrA := GetFormatType(MI.PropertyAccessor.GetProperty(PR_MSG_EDITOR_FORMAT));
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_MSG_EDITOR_FORMAT';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_LONG';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_NATIVE_BODY_INFO : String = 'http://schemas.microsoft.com/mapi/proptag/0x10160003'; {PT_LONG}
    Try
      StrA := GetBestFormatType(MI.PropertyAccessor.GetProperty(PR_NATIVE_BODY_INFO));
    Except;
      StrA := 'UnKnown';
    End;
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_NATIVE_BODY_INFO';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_LONG';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_MESSAGE_CODEPAGE : String = 'http://schemas.microsoft.com/mapi/proptag/0x3FFD0003'; {PT_LONG}
    Try
      StrA := IntToStr(MI.PropertyAccessor.GetProperty(PR_MESSAGE_CODEPAGE));
    Except;
      StrA := 'UnKnown';
    End;
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_MESSAGE_CODEPAGE';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_LONG';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_OBJECT_TYPE : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFE0003'; {PT_LONG}
    Try
      StrA := GetObjectType(MI.PropertyAccessor.GetProperty(PR_OBJECT_TYPE));
    Except;
      StrA := 'UnKnown';
    End;
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_OBJECT_TYPE';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_LONG';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    //PR_INTERNET_CPID : String = 'http://schemas.microsoft.com/mapi/proptag/0x3FDE0003'; {PT_LONG}
    Try
      StrA := IntToStr(MI.PropertyAccessor.GetProperty(PR_INTERNET_CPID));
    Except;
      StrA := 'UnKnown';
    End;
    CurProp := CurProp + 1;
    MailPropTbl.Insert;
    MailPropTbl.FieldByName('Number').AsInteger := CurProp;
    MailPropTbl.FieldByName('PropName').AsString := 'PR_INTERNET_CPID';
    MailPropTbl.FieldByName('PropType').AsString := 'PT_LONG';
    MailPropTbl.FieldByName('PropValue').AsString  := StrA;
    MailPropTbl.Post;

    MailPropTbl.First;
  Finally
    MailPropTbl.EnableControls;
  End;

end;

end.
