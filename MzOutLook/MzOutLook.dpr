program MzOutLook;



uses
  Vcl.Forms,
  uMain in 'uMain.pas' {PlayWithOotlookFrm},
  Vcl.Themes,
  Vcl.Styles,
  uMailProps in 'uMailProps.pas' {MailPropsFrm},
  uMaileMsgViewer in 'uMaileMsgViewer.pas' {MaileMsgViewerFrm};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TPlayWithOotlookFrm, PlayWithOotlookFrm);
  Application.Run;
end.
