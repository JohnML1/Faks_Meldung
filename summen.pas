unit summen;

{$mode objfpc}{$H+}

interface

uses
  Classes,LCLIntf, SysUtils, FileUtil, PrintersDlgs, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, Grids, StdCtrls, Printers, IniPropStorage;

type

  { TFormSum }

  TFormSum = class(TForm)
    BtnPrint: TButton;
    BtnSave: TButton;
    IniPropStorage1: TIniPropStorage;
    Panel1: TPanel;
    Panel2: TPanel;
    PrintDialog1: TPrintDialog;
    StringGrid1: TStringGrid;
    procedure BtnPrintClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
  private

  public
    procedure PrintGrid(sGrid: TStringGrid; sTitle: string);

  end;

var
  FormSum: TFormSum;

implementation

{$R *.lfm}

{ TFormSum }

uses unit1;

procedure TFormSum.PrintGrid(sGrid: TStringGrid; sTitle: string);
var
  X1, X2: Integer;
  Y1, Y2: Integer;
  TmpI: Integer;
  F: Integer;
  TR: TRect;
begin
  if not PrintDialog1.Execute then exit;
  Printer.Title := sTitle;
  Printer.BeginDoc;
  Printer.Canvas.Pen.Color  := 0;
  Printer.Canvas.Font.Name  := 'Times New Roman';
  Printer.Canvas.Font.Size  := 12;
  Printer.Canvas.Font.Style := [fsBold, fsUnderline];
  Printer.Canvas.TextOut(0, 100, Printer.Title);
  for F := 1 to sGrid.ColCount - 1 do
  begin
    X1 := 0;
    for TmpI := 1 to (F - 1) do
      X1 := X1 + 5 * (sGrid.ColWidths[TmpI]);
    Y1 := 300;
    X2 := 0;
    for TmpI := 1 to F do
      X2 := X2 + 5 * (sGrid.ColWidths[TmpI]);
    Y2 := 450;
    TR := Rect(X1, Y1, X2 - 30, Y2);
    Printer.Canvas.Font.Style := [fsBold];
    Printer.Canvas.Font.Size := 7;
    Printer.Canvas.TextRect(TR, X1 + 50, 350, sGrid.Cells[F, 0]);
    Printer.Canvas.Font.Style := [];
    for TmpI := 1 to sGrid.RowCount - 1 do
    begin
      Y1 := 150 * TmpI + 300;
      Y2 := 150 * (TmpI + 1) + 300;
      TR := Rect(X1, Y1, X2 - 30, Y2);
      Printer.Canvas.TextRect(TR, X1 + 50, Y1 + 50, sGrid.Cells[F, TmpI]);
    end;
  end;
  Printer.EndDoc;
end;

procedure TFormSum.BtnPrintClick(Sender: TObject);
begin
   PrintGrid(StringGrid1,'Summen zu Fehlermeldung, Stand: ' + FormatDateTime('dd.mm.yy    hh:nn', now));
end;

procedure TFormSum.BtnSaveClick(Sender: TObject);
var FName : String;
    append : string;
    OldTimeSeparator : Char;
begin
  OldTimeSeparator := TimeSeparator;
  TimeSeparator := '-';
  append := 'Summen _zu Fehlermeldung_' + FormatDateTime('dd.mm.yy_hh:nn', now) + '.csv';
  TimeSeparator := OldTimeSeparator;

  if DirectoryExists(Form1.DirectoryEdit1.text) then
  begin
    FName := IncludeTrailingBackslash(Form1.DirectoryEdit1.text) + append;
  end
  else
  begin
    FName := ExePath +  append;
  end;

  StringGrid1.SaveToCSVFile(FName,';');

  if not OpenUrl(FName) then
   ShowMessage(FName + ' konnte nicht geöffnet werden.');

end;

end.

