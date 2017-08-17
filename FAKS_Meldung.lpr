program FAKS_Meldung;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, zcomponent,
  dbflaz, runtimetypeinfocontrols, printer4lazarus, lazdbexport,
  Unit1, uniqueinstance_package, summen, Unit2;

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TFormSum, FormSum);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
