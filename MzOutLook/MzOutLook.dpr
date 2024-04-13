program MzOutLook;



uses
  madExcept,
  Vcl.Forms,
  uMain in 'uMain.pas' {PlayWithOotlookFrm},
  Vcl.Themes,
  Vcl.Styles,
  uMailProps in 'uMailProps.pas' {MailPropsFrm},
  uMaileMsgViewer in 'uMaileMsgViewer.pas' {MaileMsgViewerFrm},
  SelectMapiFolder in 'SelectMapiFolder.pas' {SelectMapiFolderFrm};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TPlayWithOotlookFrm, PlayWithOotlookFrm);
  Application.Run;
end.
