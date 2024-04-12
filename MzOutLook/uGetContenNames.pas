unit uGetContenNames;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls,
  Outlook2010, ActiveX, ComObj, ole2, OleServers, Vcl.OleServer;

type
  TForm26 = class(TForm)
    Panel1: TPanel;
    Memo1: TMemo;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form26: TForm26;

implementation
{$R *.dfm}

Uses
  ApiUtil, uFileUtil;

procedure TForm26.Button1Click(Sender: TObject);
const
  olFolderContacts = $0000000A;
  olFolderSuggestedContacts =  $0000001E;
var
  Contacts, Contact, SuggestedContactsRoot : OleVariant;
  i: Integer;
  NS : Outlook2010._NameSpace;
  olStore : Outlook2010._Store;
  OutlookApp : Outlook2010.TOutlookApplication;
  FO       : MAPIFolder;
  foItems  : Outlook2010.Items;
  foItem   : Outlook2010.MailItem;
  CurItem  : Integer;
  JsonName : String;
begin
  Memo1.Lines.Clear;
  JsonName := 'c:\a\abc.json';
  DeleteFile(JsonName);

  WriteTextFile(JsonName, '[', False, False);
  Try
    OutlookApp := TOutlookApplication.Create(Nil);
    OutlookApp.ConnectKind := ckRunningOrNew;
    OutlookApp.Connect;
    NS := OutlookApp.GetNameSpace('MAPI');

    Contacts := NS.GetDefaultFolder(olFolderContacts);
    for i := 1 to Contacts.Items.Count do
    begin
      Contact := Contacts.Items.Item(i);
      {now read property of contact. full name and email address}
      Memo1.Lines.Add(Contact.FullName + ' <' + Contact.Email1Address + '>');

      WriteTextFile(JsonName, '{', False, False);
      WriteTextFile(JsonName, '"FullName":"' + Contact.FullName + '"', False, False);
      WriteTextFile(JsonName, '"Email1Address":"' + Contact.Email1Address + '"', False, False);
      WriteTextFile(JsonName, '},', False, False);
    end;

    Try
      SuggestedContactsRoot := NS.GetDefaultFolder(olFolderSuggestedContacts);
      for i := 1 to SuggestedContactsRoot.Items.Count do
      begin
        Contact := SuggestedContactsRoot.Items.Item(i);
        Memo1.Lines.Add(Contact.FullName + ' <' + Contact.Email1Address + '>');

        WriteTextFile(JsonName, '{', False, False);
        WriteTextFile(JsonName, '"FullName":"' + Contact.FullName + '"', False, False);
        WriteTextFile(JsonName, '"Email1Address":"' + Contact.Email1Address + '"', False, False);
        WriteTextFile(JsonName, '},', False, False);
      end;
    Except;
    End;

    // now get the first 1000 mailes
    olStore := NS.DefaultStore;
    FO      := olStore.GetDefaultFolder(olFolderInbox);

    foItems := FO.Items;
    foItems.Sort('ReceivedTime',True);

    For CurItem := 1 to foItems.Count Do
    begin
      if CurItem > 1000 then
        Break;
      if Supports(foItems.Item(CurItem), MailItem, foItem) then
      begin
        Memo1.Lines.Add(foItem.SenderName + ' <' + foItem.SenderEmailAddress + '>');

        WriteTextFile(JsonName, '{', False, False);
        WriteTextFile(JsonName, '"FullName":"' + foItem.SenderName + '"', False, False);
        WriteTextFile(JsonName, '"Email1Address":"' + foItem.SenderEmailAddress + '"', False, False);
        WriteTextFile(JsonName, '},', False, False);
      end;
    end;
  Finally
    FreeAndNil(OutlookApp);
  End;

  WriteTextFile(JsonName, '{', False, False);
  WriteTextFile(JsonName, '"FullName":"end"', False, False);
  WriteTextFile(JsonName, '"Email1Address":"end"', False, False);
  WriteTextFile(JsonName, '}', False, False);

  WriteTextFile(JsonName, ']', False, False);
end;

end.
