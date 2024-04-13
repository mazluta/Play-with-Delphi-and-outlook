unit SelectMapiFolder;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.DBClient, Vcl.Buttons;

type
  TSelectMapiFolderFrm = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnSelect: TButton;
    btnReturn: TButton;
    FoldersGrid: TDBGrid;
    FoldersListDS: TDataSource;
    FoldersListTbl: TClientDataSet;
    FoldersListTblNumber: TIntegerField;
    FoldersListTblFolderName: TWideStringField;
    FoldersListTblEntryId: TWideStringField;
    FoldersListTblStoreID: TWideStringField;
    FoldersListTblNewname: TWideStringField;
    Panel3: TPanel;
    lbSearchFor: TLabel;
    edSearchFor: TEdit;
    sbFromStart: TSpeedButton;
    sbFindnext: TSpeedButton;
    sbFindPrior: TSpeedButton;
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnReturnClick(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure FoldersGridDblClick(Sender: TObject);
    procedure sbFromStartClick(Sender: TObject);
    procedure sbFindnextClick(Sender: TObject);
    procedure sbFindPriorClick(Sender: TObject);
    procedure edSearchForKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    OnFirstTime : Boolean;
  public
    { Public declarations }
    SelectedEnteryID : String;
  end;

var
  SelectMapiFolderFrm: TSelectMapiFolderFrm;

implementation
{$R *.dfm}

Uses
  uMain;

procedure TSelectMapiFolderFrm.btnReturnClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TSelectMapiFolderFrm.btnSelectClick(Sender: TObject);
begin
  SelectedEnteryID := FoldersListTbl.FieldByName('EntryId').AsString;
  ModalResult := mrOk;
end;

procedure TSelectMapiFolderFrm.edSearchForKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    sbFromStart.Click;
end;

procedure TSelectMapiFolderFrm.FoldersGridDblClick(Sender: TObject);
begin
  if not FoldersListTbl.Active then
    exit;
  if FoldersListTbl.RecordCount = 0 then
    exit;
  btnSelect.Click;
end;

procedure TSelectMapiFolderFrm.FormActivate(Sender: TObject);
begin
  if OnFirstTime Then
  Try
    FoldersListTbl.CloneCursor(PlayWithOotlookFrm.FoldersListTbl,False, False);
    FoldersListTbl.IndexFieldNames := 'FolderName';
  Finally
    OnFirstTime := False;
  End;
end;

procedure TSelectMapiFolderFrm.FormCreate(Sender: TObject);
begin
  OnFirstTime := True;
  SelectedEnteryID := '';
  edSearchFor.Text := '';
end;

procedure TSelectMapiFolderFrm.sbFindnextClick(Sender: TObject);
Var
  RecNo : Integer;
  Found : Boolean;
  SearcFor : String;
begin
  if not FoldersListTbl.Active then
    exit;
  if FoldersListTbl.RecordCount = 0 then
    exit;

  SearcFor := Trim(edSearchFor.Text);
  if SearcFor = '' then
    Exit;

  Try
    FoldersListTbl.DisableControls;
    RecNo := FoldersListTbl.RecNo;
    Found := False;
    FoldersListTbl.Next;
    while not FoldersListTbl.EOF do
    begin
      if Pos(SearcFor, FoldersListTbl.FieldByName('FolderName').AsString) > 0 then
      begin
        Found := True;
        Break;
      end;
      FoldersListTbl.Next;
    end;

    if not Found then
      FoldersListTbl.RecNo := RecNo;
  Finally
    FoldersListTbl.EnableControls;
  End;
end;

procedure TSelectMapiFolderFrm.sbFindPriorClick(Sender: TObject);
Var
  RecNo : Integer;
  Found : Boolean;
  SearcFor : String;
begin
  if not FoldersListTbl.Active then
    exit;
  if FoldersListTbl.RecordCount = 0 then
    exit;

  SearcFor := Trim(edSearchFor.Text);
  if SearcFor = '' then
    Exit;

  Try
    FoldersListTbl.DisableControls;
    RecNo := FoldersListTbl.RecNo;
    Found := False;
    FoldersListTbl.Prior;
    while not FoldersListTbl.BOF do
    begin
      if Pos(SearcFor, FoldersListTbl.FieldByName('FolderName').AsString) > 0 then
      begin
        Found := True;
        Break;
      end;
      FoldersListTbl.Prior;
    end;

    if not Found then
      FoldersListTbl.RecNo := RecNo;
  Finally
    FoldersListTbl.EnableControls;
  End;
end;

procedure TSelectMapiFolderFrm.sbFromStartClick(Sender: TObject);
Var
  Found : Boolean;
  SearcFor : String;
begin
  if not FoldersListTbl.Active then
    exit;
  if FoldersListTbl.RecordCount = 0 then
    exit;

  SearcFor := Trim(edSearchFor.Text);
  if SearcFor = '' then
    Exit;

  Try
    FoldersListTbl.DisableControls;
    Found := False;
    FoldersListTbl.First;
    while not FoldersListTbl.EOF do
    begin
      if Pos(SearcFor, FoldersListTbl.FieldByName('FolderName').AsString) > 0 then
      begin
        Found := True;
        Break;
      end;
      FoldersListTbl.Next;
    end;

    if not Found then
      FoldersListTbl.First;
  Finally
    FoldersListTbl.EnableControls;
  End;
end;

end.
