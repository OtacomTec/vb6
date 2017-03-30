program Demo;

uses
  Forms,
  formMain in 'formMain.pas' {frmMain},
  VaUtils in '..\Source Files\VaUtils.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
