unit uMaileMsgViewer;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Outlook2010, ActiveX, ComObj, {ole2, OleServers, Vcl.OleServer, Vcl.OleCtrls,}
  AxCtrls, IOutils, SHDocVw, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.OleCtrls, RichEditBrowser, RVScroll, RichView, RVStyle;


// Constants for enum MailFormat
type
  OlBodyFormat = TOleEnum;
const
  olFormatUnspecified = $00000000;
  olFormatPlain       = $00000001;
  olFormatHTML        = $00000002;
  olFormatRichText    = $00000003;

// Constants for enum ViewerType
type
  olViewerType = TOleEnum;
const
  olViewerTypePlainText = $00000001;
  olViewerTypeHtml      = $00000002;
  olViewerTypeRtf       = $00000003;
  olViewerTypeMhtml     = $00000004;

type
  TFN_WrapCompressedRTFStream = function(lpCompressedRTFStream: IStream; ulFlags: ULONG; out lpUncompressedRTFStream: IStream): HRESULT; stdcall;

const
  WrapCompressedRTFStream_N = 'WrapCompressedRTFStream@12';

type
  TMaileMsgViewerFrm = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    edFrom: TEdit;
    edTo: TEdit;
    edBCC: TEdit;
    edSubject: TEdit;
    edReciveDate: TEdit;
    MsgViewerPanel: TPanel;
    MailPlainText: TMemo;
    MailBrowser: TWebBrowser;
    MailRichEdit: TRichView;
    RVStyle1: TRVStyle;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    OnFirstTime : Boolean;
  public
    { Public declarations }
    ViewerType : olViewerType;
    MI : Outlook2010.MailItem;
  end;

var
  MaileMsgViewerFrm: TMaileMsgViewerFrm;

var
  MAPI32DLLNAME: string = 'MAPI32.DLL';
  MAPI32Module: HModule = 0;

implementation
{$R *.dfm}

Uses
  uOutlookPR;

procedure TMaileMsgViewerFrm.FormActivate(Sender: TObject);
Const
  _EmptyStr : String = ' ';

Var
  FileName    : String;
  fPlainText  : WideString;
  fHtmlText   : WideString;
  fRichText   : OleVariant;
  //fRichText   : String;
  RtfStr      : AnsiString;
  RtfFileName : String;
  HasRTF      : Boolean;
  olPA        : _PropertyAccessor;
  FSsrc       : TStringStream;
  FSdst       : TStringStream;
  Stream      : TStringStream;
  StreamB     : TStream;

  Rslt        : Integer;
  OleStream,
   Uncompressed : IStream;

  function WrapCompressedRTFStream(lpCompressedRTFStream: IStream; ulFlags: ULONG; out lpUncompressedRTFStream: IStream): HRESULT;
  Var
    _WrapCompressedRTFStream: TFN_WrapCompressedRTFStream;
  begin
    MAPI32Module := LoadLibrary(PWideChar(MAPI32DLLNAME));
    @_WrapCompressedRTFStream := GetProcAddress(MAPI32Module, WrapCompressedRTFStream_N);
    if @_WrapCompressedRTFStream <> nil then
      Result := _WrapCompressedRTFStream(lpCompressedRTFStream, ulFlags, lpUncompressedRTFStream)
    else
      Result := -1;
  end;

  function GetWinTempDir : String;
  var
    le:Integer;
  Const
    MAX_PATH : Dword = 1024;
  begin
    SetLength(result, MAX_PATH);
    le:=GetTempPath(MAX_PATH, PChar(result));
    SetLength(result,le);
  end;

  function  RemoveAllEnterChar(SrcString : WideString) : WideString;
  begin
    // Remove The Last $D$A From SrcString;
    Result := SrcString;
    Result := StringReplace(Result,#13+#10,' ',[rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result,#13,' ',[rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result,#10,' ',[rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result,#$D#$A,' ',[rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result,#$D,' ',[rfReplaceAll, rfIgnoreCase]);
    Result := StringReplace(Result,#$A,' ',[rfReplaceAll, rfIgnoreCase]);
  end;

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

  function GetBinaryString(Value: Variant): String;
  var
    SafeArray: PVarArray;
    RtfStr: AnsiString;
  begin
    SafeArray := VarArrayAsPSafeArray(Value);
    Assert(SafeArray.ElementSize = 1);
    SetString(RtfStr, PAnsiChar(SafeArray.Data), SafeArray.Bounds[0].ElementCount);
    Result := String(RtfStr);
  end;

//  var
//    aFileStream: TFileStream;
//    iStr: TStreamAdapter;
//    iRes , iRes1, iRes2: Largeint;
//    aStreamStat: TStatStg;
//    aStreamContent: IStream;
//  begin
//    aFileStream := TFileStream.Create('<...>', fmCreate);
//    try
//      aStreamContent := <...> as IStream;
//      aStreamContent.Seek(0, 0, iRes);
//      iStr := TStreamAdapter.Create(aFileStream, soReference);
//      aStreamContent.Stat(aStreamStat, 1);
//      aStreamContent.CopyTo(iStr, aStreamStat.cbSize , iRes1, iRes2);
//    finally
//      aFileStream.Free;
//    end;
//  end;

//  procedure TForm1.Button1Click(Sender: TObject);
//  var
//    ovOutlook: OLEVariant;
//    ovNameSpace: OLEVariant;
//    ovFolder: OLEVariant;
//    ovItem: OLEVariant;
//    RtfStr: AnsiString;
//    I: Integer;
//  begin
//    ovOutlook := CreateOleObject('Outlook.Application');
//    ovNameSpace := ovOutlook.GetNameSpace('MAPI');
//    ovNameSpace.Logon(EmptyParam, EmptyParam, False, True);
//    ovFolder := ovNameSpace.GetDefaultFolder(6); // olFolderInbox
//    for I := 1 to ovFolder.Items.Count do
//    begin
//      Memo1.Lines.Add(I.ToString);
//      if VarIsNull(ovFolder.Items[I]) or VarIsEmpty(ovFolder.Items[I]) then Continue;
//      ovItem := ovNameSpace.GetItemFromID(ovFolder.Items[I].EntryID);
//
//      if ovItem.BodyFormat = 3 then
//      begin
//        RtfStr := GetBinaryString(ovItem.RTFBody);
//        Memo1.Lines.Add(ovFolder.Items[I].Subject + '=' + RtfStr);
//      end;
//
//    end;
//
//    ovItem := Unassigned;
//    ovFolder := Unassigned;
//    ovNameSpace := Unassigned;
//    ovOutlook := Unassigned;
//
//  end;


begin
  If OnFirstTime Then
  Try
    edFrom.Text    := MI.SenderName + '[' + MI.SenderEmailAddress + ']';
    edTo.Text      := MI.CC;
    edBCC.Text     := MI.BCC;
    edSubject.Text := MI.Subject;
    edReciveDate.Text := DateTimeToStr(MI.ReceivedTime);

    Try
      case ViewerType of
        olViewerTypePlainText : begin
                                  fPlainText := Trim(MI.Body);
                                  If fPlainText <> '' Then
                                    MailPlainText.Lines.Text := fPlainText
                                  else
                                    MailPlainText.Lines.Text := 'np plain text found';

                                  MailPlainText.Align := alClient;
                                  MailPlainText.AlignWithMargins := True;
                                  MailPlainText.Visible := True;
                                end;
        olViewerTypeHtml      : begin
                                  FileName := RemoveBackSlashChar(GetWinTempDir) + '\TmpHtml' + IntToStr(GetTickCount) + '.html';
                                  fHtmlText := Trim(MI.HTMLBody);
                                  If fHtmlText <> '' Then
                                    TFile.WriteAllText(FileName, fHtmlText, TEncoding.UTF8)
                                  else
                                  begin
                                    TFile.WriteAllText('<html><body><p>no Html Text found </p></body></html>', FileName)
                                  end;

                                  MailBrowser.Align := alClient;
                                  MailBrowser.AlignWithMargins := True;
                                  MailBrowser.Navigate2(FileName);
                                  MailBrowser.Visible := True;
                                end;
        olViewerTypeRtf       : begin
                                  //fRichText := Trim(MI.RTFBody);
                                  RtfFileName := RemoveBackSlashChar(GetWinTempDir) + '\TmpRtf' + IntToStr(GetTickCount) + '.Rtf';

                                  //PR_RTF_COMPRESSED : String = 'http://schemas.microsoft.com/mapi/proptag/0x10090102'; {PT_BINARY}
                                  olPA := MI.PropertyAccessor;
                                  Try
                                    fRichText := olPA.BinaryToString(MI.PropertyAccessor.GetProperty(PR_RTF_COMPRESSED));
                                  Except;
                                    fRichText := MI.RTFBody;
                                    // this hold ansi compressed string
                                  End;

                                  HasRTF := MI.PropertyAccessor.GetProperty(PR_RTF_IN_SYNC);
                              //
                              //   FSsrc := TStringStream.Create(fRichText , TEncoding.ascii);
                              //   FSsrc.Position := 0;
                              //   With TOleStream.Create(OleStream) do
                              //   begin
                              //     CopyFrom(FSsrc, FSsrc.Size);
                              //     Position := 0;
                              //     Free;
                              //   end;
                              //
                              //   Rslt := WrapCompressedRTFStream(OleStream, 0, Uncompressed);
                              //   If Rslt = 0 Then
                              //   begin
                              //     //If fRichText <> '' Then
                              //     //  MailRichEdit.Lines.Text := fRichText
                              //     //else
                              //     //  MailRichEdit.Lines.Text := 'no RTF text found';
                              //
                              //     //Uncompressed.SaveToFile('c:\a\aa.rtf');
                              //     MailRichEdit.Lines.LoadFromFile('c:\a\aa.rtf');
                              //   end
                              //   else
                              //   begin
                              //     MailRichEdit.Lines.Text := 'no RTF text found';
                              //   end;

                                  if HasRTF then
                                  begin
                                    //MI.SaveAs(RtfFileName , olRTF);

                                    RtfStr := GetBinaryString(fRichText);
                                    TFile.WriteAllText(RtfFileName, RtfStr);
                                    MailRichEdit.LoadRTF(RtfFileName);
                                    MailRichEdit.Align := alClient;
                                    MailRichEdit.AlignWithMargins := True;
                                    MailRichEdit.Visible := True;

                                    //Stream    := TStringStream.Create(fRichText);
                                    //OleStream := TStreamAdapter.Create(Stream, soReference) as IStream;
                                    //Rslt := WrapCompressedRTFStream(OleStream, 0, Uncompressed);
                                    //If Rslt = 0 Then
                                    //begin
                                    //  StreamB := TOleStream.Create(Uncompressed);
                                    //  (StreamB as TStringStream).SaveToFile(RtfFileName);
                                    //  MailRichEdit.LoadRTF(RtfFileName);
                                    //  MailRichEdit.Align := alClient;
                                    //  MailRichEdit.AlignWithMargins := True;
                                    //  MailRichEdit.Visible := True;
                                    //end;
                                    //MailBrowser.Align := alClient;
                                    //MailBrowser.AlignWithMargins := True;
                                    //MailBrowser.Navigate2(RtfFileName);
                                    //MailBrowser.Visible := True;
                                  end
                                  else
                                  begin
                                    fPlainText := Trim(MI.Body);
                                    If fPlainText <> '' Then
                                      MailPlainText.Lines.Text := fPlainText
                                    else
                                      MailPlainText.Lines.Text := 'np plain text found';

                                    MailPlainText.Align := alClient;
                                    MailPlainText.AlignWithMargins := True;
                                    MailPlainText.Visible := True;
                                  end;
                                end;
        olViewerTypeMhtml     : begin
                                  FileName := RemoveBackSlashChar(GetWinTempDir) + '\TmpMHtml' + IntToStr(GetTickCount) + '.mhtml';
                                  MI.SaveAs(FileName, olMHTML);

                                  MailBrowser.Align := alClient;
                                  MailBrowser.AlignWithMargins := True;
                                  MailBrowser.Navigate2(FileName);
                                  MailBrowser.Visible := True;
                                end;
      end;
    Except;
       ShowMessage('Fail Load Viewer');
    End;

    MsgViewerPanel.Visible := True;

  Finally
    OnFirstTime := False;
  End;
end;

procedure TMaileMsgViewerFrm.FormCreate(Sender: TObject);
begin
  OnFirstTime := True;

  MailPlainText.Visible := False;
  MailRichEdit.Visible  := False;
  MailBrowser.Visible   := False;

  MsgViewerPanel.Visible := False;
end;

end.
