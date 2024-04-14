program MzComparePST;

uses
  Vcl.Forms,
  uMainCompare in 'uMainCompare.pas' {MainCompareFrm},
  uMailProps in 'uMailProps.pas' {MailPropsFrm};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainCompareFrm, MainCompareFrm);
  Application.Run;
end.
