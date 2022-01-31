program SQLQueryToPascal;

uses
  Forms,
  ufrmPrinc in 'ufrmPrinc.pas' {frmPrinc},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'SQL Query To Pascal';
  Application.CreateForm(TfrmPrinc, frmPrinc);
  Application.Run;
end.
