unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  locale_de, Classes, SysUtils, db, fpSQLExport, fpstdexports, FileUtil, Forms,
  Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, CheckLst, SynEdit,
  fpdataexporter, LCLType, DbCtrls, DBGrids, Menus, ZDataset, ZSqlMetadata,
  Variants, StrUtils, LCLIntf, fpsexport;

type

  { TForm2 }

  TForm2 = class(TForm)
    BtnInsertFieldNames: TButton;
    CheckListBox1: TCheckListBox;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    FPSExport1: TFPSExport;
    Label1: TLabel;
    MnExport: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    PopupFieldProps: TPopupMenu;
    RadioGroup1: TRadioGroup;
    ZReadOnlyQuery1: TZReadOnlyQuery;
    procedure BtnInsertFieldNamesClick(Sender: TObject);
    procedure CheckListBox1Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure MnExportClick(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure ZReadOnlyQuery1AfterOpen(DataSet: TDataSet);
  private

  public

  end;



var
  Form2: TForm2;

implementation

uses unit1, my_utils;


{$R *.lfm}

{ TForm2 }


procedure TForm2.FormShow(Sender: TObject);
var x :Integer;
begin
  try
    if not Form1.QFaks.Active then
    begin
     Exception.Create('Die Datenbankabfrage ist nicht geöffnet (QFaks)');
     exit;
    end;

    if not (Screen.ActiveControl is TSynedit) then
    Exception.Create('Bitte erst die Einfügestelle im SQL-Code anklicken');

    (* damit ncht ggf. ALLE Felder angezeigt werden *)
    CheckListBox1.Items.Clear;

    (* ALLE Felder aus f2fsv oder nur die Felder aus QFAKS? *)
    if Messagedlg('Welche Felder sollen angezeigt werden? JA=Felder der eingebauten Abfrage, NEIN=ALLE Felder der Oracle Tabelle f2fsv.',mtConfirmation,[mbYes,mbNo],0)= mrYes then
    begin


      for x := 0 to Form1.QFaks.Fields.count -1 do
      begin
        (* da ja evtl. öfters OnFormShow ausgelöst wird, nicht dopplte Feldnamen eintragen *)
       if CheckListBox1.Items.IndexOf(Form1.QFaks.Fields[x].FieldName) = -1 then
         CheckListBox1.Items.add(Form1.QFaks.Fields[x].FieldName);
      end;

    end
    else
    begin
     if not Form1.ZCheckFields.Active then Form1.ZCheckFields.Active := true;

     for x := 0 to Form1.ZCheckFields.Fields.count -1 do
     begin
       (* da ja evtl. öfters OnFormShow ausgelöst wird, nicht dopplte Feldnamen eintragen *)
      if CheckListBox1.Items.IndexOf(Form1.ZCheckFields.Fields[x].FieldName) = -1 then
        CheckListBox1.Items.add(Form1.ZCheckFields.Fields[x].FieldName);
     end;

    end;

  finally
    BtnInsertFieldNames.SetFocus;

  end;
end;

procedure TForm2.MnExportClick(Sender: TObject);
var FName : String;
begin
  try
    Jei;
    FName := IncludeTrailingBackslash(Form1.DirectoryEdit1.Directory) + 'Feldeigenschaften_FAKS-Tabelle_F2FSV.xls';
    FPSExport1.FileName:=FName;

    FPSExport1.Execute;
    Nei;

    if FileExists(FName) then OpenURL(FName);

  finally
    Nei;

  end;
end;

procedure TForm2.RadioGroup1Click(Sender: TObject);
var x : integer;
begin
  case RadioGroup1.ItemIndex of
  0:
    begin
      CheckListBox1.CheckAll(cbChecked);
    end;
  1:
    begin
      CheckListBox1.CheckAll(cbUnChecked);
    end;
  2:
    begin
      for x := 0  to CheckListBox1.ITEMS.COUNT -1 do
       CheckListBox1.Checked[x] := not CheckListBox1.Checked[x];
    end;


  end;
end;

procedure TForm2.ZReadOnlyQuery1AfterOpen(DataSet: TDataSet);
begin
   DBGrid1.AutoAdjustColumns;
   DBGrid1.Columns[DBGrid1.Columns.Count -1].Width:=255;
end;

procedure TForm2.BtnInsertFieldNamesClick(Sender: TObject);
var FeldNamen, salias : string;
    x : integer;
begin

   (* FeldNamen ggf. den alias a. voranstellen *)
   if Form1.cbSQL.Checked then salias:='a.'
      else salias := '';

   (* ausgewählte FeldNamen in kommagetrennten String speichern *)
   for x := 0 to  CheckListBox1.Items.Count -1 do
   begin
     if CheckListBox1.Checked[x] then
     begin

      if FeldNamen = '' then
         FeldNamen :=  salias + CheckListBox1.Items[x]
       else
         FeldNamen := FeldNamen + ', ' + salias + CheckListBox1.Items[x];

     end;

   end;

   (* wurde überhaupt was ausgewählt? *)
   if WordCount(FeldNamen,[',']) = 0 then
   begin
     ShowMessage('Es wurde kein Feldname ausgewählt. Habe fertig :-)');
     exit;
   end;

   (* noch ein Komma anfügen, wird meist ok sein *)
   FeldNamen := FeldNamen + ', ';


   Form1.Memo1.InsertTextAtCaret(FeldNamen);

   Form2.Close;

   Application.ProcessMessages;

   ShowMessage('Achtung: Hinter dem letzten eingefügten Feldnamen steht ein Komma!!');

end;

procedure TForm2.CheckListBox1Click(Sender: TObject);
begin
   (* zu angeklicktem Feldnamen Zeile mit Details im DBGrid1 markieren *)
   ZReadOnlyQuery1.Locate('COLUMN_NAME',VarArrayOf([CheckListBox1.items[CheckListBox1.ItemIndex]]),[]);

   (* Spalte mit Feldnamen selektieren *)
   DBGrid1.SelectedField := ZReadOnlyQuery1.FieldByName('COLUMN_NAME');

end;

procedure TForm2.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
  );
begin
  if Key = VK_ESCAPE then Form2.Close;
  if Key = VK_RETURN then BtnInsertFieldNamesClick(Sender);
end;

end.

