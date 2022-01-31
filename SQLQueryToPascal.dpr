program SQLQueryToPascal;

uses
  Forms,
  ufrmPrinc in 'ufrmPrinc.pas' {frmPrinc},
  Vcl.Themes,
  Vcl.Styles,
  uConfiguracoes in 'uConfiguracoes.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'SQL Query To Pascal';
  Application.CreateForm(TfrmPrinc, frmPrinc);
  Application.Run;
end.
