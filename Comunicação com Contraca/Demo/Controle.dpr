program Controle;

uses
  Forms,
  ControleUnit in 'ControleUnit.pas' {ControleForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TControleForm, ControleForm);
  Application.Run;
end.
