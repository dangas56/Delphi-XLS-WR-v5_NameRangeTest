program prjNameRangeCheck;

uses
  Vcl.Forms,
  Main in 'Main.pas' {Form10},
  XLSTestRun4 in 'XLSTestRun4.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm10, Form10);
  Application.Run;
end.
