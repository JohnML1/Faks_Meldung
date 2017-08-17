unit unit_test;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, db, dbf, FileUtil, SynEdit, SynHighlighterSQL, Forms,
  Controls, Graphics, Dialogs, ComCtrls, ExtCtrls, DbCtrls, EditBtn, StdCtrls,
  Spin, DBGrids, IniPropStorage, Menus, AsyncProcess, UniqueInstance, eventlog,
  ZConnection, ZDataset, ZSqlMonitor, ZSqlMetadata;

type

  { TForm2 }

  TForm2 = class(TForm)
    ApplicationProperties1: TApplicationProperties;
    AsyncProcess1: TAsyncProcess;
    AuswahlFilter: TMenuItem;
    BtnConnect: TButton;
    BtnConnect1: TButton;
    BtnConnect2: TButton;
    BtnConnect3: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    CbincMonth: TToggleBox;
    CbincMonth1: TToggleBox;
    CbincMonth2: TToggleBox;
    CbincMonth3: TToggleBox;
    cbKeinGruenberg: TCheckBox;
    cbKeinGruenberg1: TCheckBox;
    cbKeinGruenberg2: TCheckBox;
    cbKeinGruenberg3: TCheckBox;
    cbKeinGruenberg_kein_Wetterau: TCheckBox;
    cbKeinGruenberg_kein_Wetterau1: TCheckBox;
    cbKeinGruenberg_kein_Wetterau2: TCheckBox;
    cbKeinGruenberg_kein_Wetterau3: TCheckBox;
    CheckLinie: TMenuItem;
    DataSource2: TDataSource;
    DateEditBis: TDateEdit;
    DateEditBis1: TDateEdit;
    DateEditBis2: TDateEdit;
    DateEditBis3: TDateEdit;
    DateEditVon: TDateEdit;
    DateEditVon1: TDateEdit;
    DateEditVon2: TDateEdit;
    DateEditVon3: TDateEdit;
    DBase_export: TMenuItem;
    Dbf1: TDbf;
    Dbf1BC_ABSCHN: TFloatField;
    Dbf1BC_COUNT: TFloatField;
    Dbf1BC_ROLLE: TFloatField;
    Dbf1BUSNR: TFloatField;
    Dbf1DATUMA: TDateField;
    Dbf1DATUMV: TDateField;
    Dbf1DIENSTNR: TFloatField;
    Dbf1ERWACHSENE: TFloatField;
    Dbf1FAHRT: TFloatField;
    Dbf1FEHLDRUCK: TBooleanField;
    Dbf1GABDATUM: TDateField;
    Dbf1GABZEIT: TStringField;
    Dbf1GATTUNG: TFloatField;
    Dbf1GKLASSE: TFloatField;
    Dbf1GNR: TFloatField;
    Dbf1HALTNR: TFloatField;
    Dbf1KINDER: TFloatField;
    Dbf1KM_ANZ: TFloatField;
    Dbf1KM_AUS: TFloatField;
    Dbf1KM_EIN: TFloatField;
    Dbf1KONZESSION: TFloatField;
    Dbf1KURS: TFloatField;
    Dbf1KURZSTR: TFloatField;
    Dbf1LFDFAWNR: TFloatField;
    Dbf1LFDIENSTNR: TFloatField;
    Dbf1LFDNR: TFloatField;
    Dbf1LINIE: TFloatField;
    Dbf1LINIEALPHA: TStringField;
    Dbf1PERSNR: TFloatField;
    Dbf1PERSONEN: TFloatField;
    Dbf1PREIS: TFloatField;
    Dbf1PREIS2: TFloatField;
    Dbf1PSTUFE: TFloatField;
    Dbf1RICHTUNG: TStringField;
    Dbf1RID: TStringField;
    Dbf1ROUTE: TFloatField;
    Dbf1SHST: TFloatField;
    Dbf1SORTE: TFloatField;
    Dbf1STORNIERT: TBooleanField;
    Dbf1SZONE: TFloatField;
    Dbf1TARIFNR: TFloatField;
    Dbf1TOUR: TFloatField;
    Dbf1VBNR: TFloatField;
    Dbf1VDTNR: TFloatField;
    Dbf1VVST: TFloatField;
    Dbf1VZONE: TFloatField;
    Dbf1WABE: TFloatField;
    Dbf1WAEHRUNG: TFloatField;
    Dbf1ZAHLART: TFloatField;
    Dbf1ZDEDATUM: TDateField;
    Dbf1ZDEZEIT: TStringField;
    Dbf1ZEITA: TStringField;
    Dbf1ZEITV: TStringField;
    Dbf1ZHST: TFloatField;
    Dbf1ZIEL: TFloatField;
    Dbf1ZONE: TFloatField;
    Dbf1ZZONE: TFloatField;
    DBGridFaks: TDBGrid;
    DBGridFaks1: TDBGrid;
    DBGridFaks2: TDBGrid;
    DBGridFaks3: TDBGrid;
    DBNavigator2: TDBNavigator;
    DBNavigator3: TDBNavigator;
    DBNavigator4: TDBNavigator;
    DBNavigator5: TDBNavigator;
    DeleteSelected: TMenuItem;
    DirectoryEdit1: TDirectoryEdit;
    DirectoryEdit2: TDirectoryEdit;
    DirectoryEdit3: TDirectoryEdit;
    DirectoryEdit4: TDirectoryEdit;
    EditNoGo: TLabeledEdit;
    EditNoGo1: TLabeledEdit;
    EditNoGo2: TLabeledEdit;
    EditNoGo3: TLabeledEdit;
    EventLog1: TEventLog;
    EventLog_anzeigen: TMenuItem;
    ExportExcel: TMenuItem;
    FilterAb: TMenuItem;
    FilterBis: TMenuItem;
    FilterCombo: TComboBox;
    FilterCombo1: TComboBox;
    FilterCombo2: TComboBox;
    FilterCombo3: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    ImageList1: TImageList;
    IniPropStorage1: TIniPropStorage;
    kopieren: TMenuItem;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    lbDatabase: TLabeledEdit;
    lbDatabase1: TLabeledEdit;
    lbDatabase2: TLabeledEdit;
    lbDatabase3: TLabeledEdit;
    lbHostname: TLabeledEdit;
    lbHostname1: TLabeledEdit;
    lbHostname2: TLabeledEdit;
    lbHostname3: TLabeledEdit;
    lbPassword: TLabeledEdit;
    lbPassword1: TLabeledEdit;
    lbPassword2: TLabeledEdit;
    lbPassword3: TLabeledEdit;
    lbUserName: TLabeledEdit;
    lbUserName1: TLabeledEdit;
    lbUserName2: TLabeledEdit;
    lbUserName3: TLabeledEdit;
    ListBoxKnownLines: TListBox;
    ListBoxKnownLines1: TListBox;
    ListBoxKnownLines2: TListBox;
    ListBoxKnownLines3: TListBox;
    ListBox_tnsnames_ora: TListBox;
    ListBox_tnsnames_ora1: TListBox;
    ListBox_tnsnames_ora2: TListBox;
    ListBox_tnsnames_ora3: TListBox;
    ListFields: TMenuItem;
    LookupStornos: TMenuItem;
    Mem: TEdit;
    Mem1: TEdit;
    Mem2: TEdit;
    Mem3: TEdit;
    Memo1: TSynEdit;
    Memo2: TSynEdit;
    Memo3: TSynEdit;
    Memo4: TSynEdit;
    MenuAddLinie: TMenuItem;
    MenuItem1: TMenuItem;
    MnMarkExported: TMenuItem;
    OpenDialog1: TOpenDialog;
    OpenLogFile: TMenuItem;
    PageControl1: TPageControl;
    PageControl2: TPageControl;
    PageControl3: TPageControl;
    PageControl4: TPageControl;
    Panel1: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    Panel2: TPanel;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    Panel23: TPanel;
    Panel24: TPanel;
    Panel25: TPanel;
    Panel26: TPanel;
    Panel27: TPanel;
    Panel28: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Pm_Search: TMenuItem;
    PopupGrid: TPopupMenu;
    PopupLinien: TPopupMenu;
    PopupSQL: TPopupMenu;
    ProgressBar1: TProgressBar;
    ProgressBar2: TProgressBar;
    ProgressBar3: TProgressBar;
    ProgressBar4: TProgressBar;
    QFaks: TZReadOnlyQuery;
    QFaks2Dbase: TMenuItem;
    RemoveFilter: TMenuItem;
    RID_as_Filter: TMenuItem;
    SaveDialog1: TSaveDialog;
    SaveToFile: TMenuItem;
    SpinEdit1: TSpinEdit;
    SpinEdit2: TSpinEdit;
    SpinEdit3: TSpinEdit;
    SpinEdit4: TSpinEdit;
    SpinEditTarifversion: TSpinEdit;
    SpinEditTarifversion1: TSpinEdit;
    SpinEditTarifversion2: TSpinEdit;
    SpinEditTarifversion3: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    Splitter4: TSplitter;
    Splitter5: TSplitter;
    Splitter6: TSplitter;
    Splitter7: TSplitter;
    Splitter8: TSplitter;
    SQLLoad: TMenuItem;
    SQL_history_lode: TMenuItem;
    Sum: TMenuItem;
    SynSQLSyn1: TSynSQLSyn;
    TabConfig: TTabSheet;
    TabConfig1: TTabSheet;
    TabConfig2: TTabSheet;
    TabConfig3: TTabSheet;
    TabDaten: TTabSheet;
    TabDaten1: TTabSheet;
    TabDaten2: TTabSheet;
    TabDaten3: TTabSheet;
    UniqueInstance1: TUniqueInstance;
    ZCheckFields: TZReadOnlyQuery;
    ZConnection1: TZConnection;
    ZSQLMetadata1: TZSQLMetadata;
    ZSQLMonitor1: TZSQLMonitor;
    ZUpDateRid: TZQuery;
     procedure ApplicationProperties1Activate(Sender: TObject);
    procedure ApplicationProperties1Exception(Sender: TObject; E: Exception);
    procedure ApplicationProperties1Hint(Sender: TObject);
    procedure AuswahlFilterClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CbincMonthClick(Sender: TObject);
    procedure cbKeinGruenbergChange(Sender: TObject);
    procedure cbKeinGruenberg_kein_WetterauChange(Sender: TObject);
    procedure CheckLinieClick(Sender: TObject);
    procedure DateEditVonAcceptDate(Sender: TObject; var ADate: TDateTime;
      var AcceptDate: boolean);
    procedure DBase_exportClick(Sender: TObject);
    procedure DBGridFaksDblClick(Sender: TObject);
    procedure DBGridFaksTitleClick(Column: TColumn);
    procedure DeleteSelectedClick(Sender: TObject);
    procedure DirectoryEdit1ButtonClick(Sender: TObject);
    procedure EventLog_anzeigenClick(Sender: TObject);
    procedure ExportExcelClick(Sender: TObject);
    procedure FilterAbClick(Sender: TObject);
    procedure FilterBisClick(Sender: TObject);
    procedure FilterComboKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FilterSetzenClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure IniPropStorage1RestoreProperties(Sender: TObject);
    procedure IniPropStorage1StoredValues1Restore(Sender: TStoredValue;
      var Value: TStoredType);
    procedure IniPropStorage1StoredValues1Save(Sender: TStoredValue;
      var Value: TStoredType);
    procedure kopierenClick(Sender: TObject);
    procedure ListBoxKnownLinesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ListFieldsClick(Sender: TObject);
    procedure LookupStornosClick(Sender: TObject);
    procedure MenuAddLinieClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MnMarkExportedClick(Sender: TObject);
    procedure OpenLogFileClick(Sender: TObject);
    procedure Panel7DblClick(Sender: TObject);
    procedure Pm_SearchClick(Sender: TObject);
    procedure RemoveFilterClick(Sender: TObject);
    procedure RID_as_FilterClick(Sender: TObject);
    procedure SaveToFileClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure SpinEditTarifversionEditingDone(Sender: TObject);
    procedure SQLLoadClick(Sender: TObject);
    procedure SQL_history_lodeClick(Sender: TObject);
    procedure UniqueInstance1OtherInstance(Sender: TObject;
      ParamCount: integer; Parameters: array of string);
    procedure QFaksAfterOpen(DataSet: TDataSet);
    procedure BtnExportExcelClick(Sender: TObject);
    procedure SumClick(Sender: TObject);
    procedure QFaksBeforeOpen(DataSet: TDataSet);
    procedure QFaks2DbaseClick(Sender: TObject);


  private
    { private declarations }
  public
    { public declarations }
    procedure JEi;
    procedure NEi;
    (* hier wird er SQL-Code erzeugt *)
    function Verbinden(): boolean;
             (* wichtig: Damit das Ü in Grünberg erhalten bleibt: UTF8ToAnsi(Zeile) *)
    function ExportDatasetToExcel(JDataset: TDataset): boolean;
    function OpenExplorer(FName: string): boolean;
    procedure angezeigteDatenkopieren1Click(Sender: TObject);
    function OpenLog(FName: string): boolean;
    procedure ShortenLog(anzLines: integer;Sender: TObject);
    function IsDate(str: string): Boolean;
    function ShowFilterInfo(Warning : boolean):boolean;
    procedure BereitsGemeldetWurden(Sender : TObject);
    procedure CheckIniFile(Sender : TObject);
    function ExistInDb (TableName : String): Boolean;
    function ExistField (FieldName : string): Boolean;
    function MarkExported (DBFName : String; mark : string = '1'; IDs : TStringList = nil) : Boolean;
    FUNCTION resourceVersionInfo: STRING;









  end;

var
  Form2: TForm2;

  ExePath: string;
  jconnected : boolean = False;
  StoppIt : boolean = False;
  FirstRun: boolean = False;
  BM: TBookmark;
  (* Linien, die nicht berücksichtigt werden sollen
     siehe VOR sql Statement in  "Verbinden"
  *)
  NoGo: string;
  abort: boolean = False;
  gesucht : variant;
  RowID : string; (* soll den Datensatz bei Fehlermeldungen kennzeichnen *)
  FLastColumn: TColumn; //store last grid column we sorted on

  MaxBuchungsDatum, SavePath (* SavePath steht in der *.ini, wohin es laut Dialog eingetragen wird *) : string;
  GesamtEinnahme, NochNichtGemeldet  : Currency;
  VonDatum, BisDatum : TDateTime;
  SQLHistory : TStringList;
  SQLHistoryIndex : Integer = 1;
  (* wird in FormCreate berechnet *)
  TarifVersion : Integer = 35;

  FTableDatabase : string = 'f2fsv';

const
  NL = chr(10) + chr(13);





implementation

{ TForm2 }

{$R *.lfm}



FUNCTION TForm2.resourceVersionInfo: STRING;

(* Unlike most of AboutText (below), this takes significant activity at run-    *)
(* time to extract version/release/build numbers from resource information      *)
(* appended to the binary.                                                      *)

VAR     Stream: TResourceStream;
        vr: TVersionResource;
        fi: TVersionFixedInfo;

BEGIN
  RESULT:= '';
  TRY

(* This raises an exception if version info has not been incorporated into the  *)
(* binary (Lazarus Project -> Project Options -> Version Info -> Version        *)
(* numbering).                                                                  *)

    Stream:= TResourceStream.CreateFromID(HINSTANCE, 1, PChar(RT_VERSION));
    TRY
      vr:= TVersionResource.Create;
      TRY
        vr.SetCustomRawDataStream(Stream);
        fi:= vr.FixedInfo;
        RESULT := 'Version ' + IntToStr(fi.FileVersion[0]) + '.' + IntToStr(fi.FileVersion[1]) +
               ' Release ' + IntToStr(fi.FileVersion[2]) + ' Build ' + IntToStr(fi.FileVersion[3]);
        vr.SetCustomRawDataStream(nil)
      FINALLY
        vr.Free
      END
    FINALLY
      Stream.Free
    END
  EXCEPT
  END
END { resourceVersionInfo } ;


// testet, ob eine Tabelle schon in der DB vorhanden ist
// benutzt dazu Metadata
function TForm2.ExistInDb (TableName : String): Boolean;
var
  TableFound: Boolean;
  ZSQLMetadata: TZSQLMetadata;
  DSSQLMetadata: TDataSource;
begin
  Result := False;
  ZSQLMetadata := TZSQLMetadata.Create (self);
  DSSQLMetadata := TDataSource.Create (self);
  ZSQLMetadata.Connection :=  ZConnection1;
  DSSQLMetadata.DataSet := ZSQLMetadata;

  // Spalten aus Tabelleninfo (Metadata) auslesen
  ZSQLMetadata.MetadataType := mdTables;
  ZSQLMetadata.Catalog := LowerCase (FTableDatabase);
  ZSQLMetadata.Open;
  TableFound := False;
  while not DSSQLMetadata.DataSet.Eof and (TableFound = FALSE) do
  begin
    if LowerCase (DSSQLMetadata.DataSet.FieldByName ('TABLE_NAME').Text) = LowerCase (TableName) then
    begin
      TableFound := True;
      Result := True;
    end;
    DSSQLMetadata.DataSet.Next;
  end;

  //aufräumen
  ZSQLMetadata.Free;
  DSSQLMetadata.Free;
end;

function TForm2.ExistField(FieldName: string): Boolean;
var x : integer;
begin
  try
   Jei;
   Result := false;
   ZCheckFields.Open;

   for x := 0 to ZCheckFields.FieldCount - 1 do
   begin
      if CompareText(ZCheckFields.Fields[x].FieldName,FieldName) = 0 then
      begin
        Result := true;
        break;
      end;
   end;

  finally
   ZCheckFields.Close;
   Nei;

  end;
end;

function TForm2.MarkExported(DBFName: String; mark: string; IDs: TStringList
  ): Boolean;
var tab : TDBF;
    Betrag : currency;
    v : variant;
    ErrLIst : TstringList;

begin
  Result := false;
  if not FileExists(DBFName) then
  begin
     Result := false;
     ShowMessage(DBFName + NL + 'wurde nicht gefunden!');
     exit;
  end;
  try
  jei;

  ErrLIst := TstringList.Create;


  tab := TDBF.Create(Application.Mainform);
  tab.TableName:= DBFName;
  tab.Open;
  tab.First;

  BM := QFaks.GetBookmark;

  QFaks.DisableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=Tab.RecordCount;
  while not tab.EOF do
  begin
     v := Trim(tab.FieldByName('RID').AsString);
     ZUpDateRid.ParamByName('RID').AsString:=v;
     ZUpDateRid.ParamByName('Vertragsnr').AsString:= mark;
     ZUpDateRid.ExecSQL;
     //if not QFaks.Locate('RID',v,[]) then ErrList.Add(v);
     tab.next;
     ProgressBar1.Position:=Tab.RecNo;
     ProgressBar1.Update;
  end;

  if ErrList.Count> 1 then
  begin
    ErrList.SaveToFile(EXEPath + 'FAKS_ErrList.txt');
    ShowMessage('Fehler sind aufgetreten bei RID, siehe '+ EXEPath + 'FAKS_ErrList.txt' );
  end;

  finally
    if QFaks.BookmarkValid(BM) then QFaks.GotoBookmark(BM);
    QFaks.FreeBookmark(BM);
    tab.close;
    ProgressBar1.Position:=0;
    FreeAndNil(tab);
    FreeAndNil(ErrLIst);
    QFaks.EnableControls;
    QFaks.Refresh;
    nei;
  end;


end;

function TForm2.IsDate(str: string): Boolean;
var
  dt: TDateTime;
begin
  Result := True;
  try
    dt := StrToDate(str);
  except
    Result := False;
  end;
end;

function TForm2.ShowFilterInfo(Warning: boolean): boolean;
begin
  if ((trim(FilterCombo.text) = '') or
     (not QFaks.Filtered)) then Warning := false;

  if Warning then
  begin
    FilterCombo.Color := ClRed;
    FilterCombo.Font.Color := ClYellow;
  end
  else
  begin
    FilterCombo.Color := clDefault;
    FilterCombo.Font.Color := clDefault;
  end;

  (* jetzt wirklich anzeigen? *)
  FilterCombo.Invalidate;
  Application.ProcessMessages;
end;

procedure TForm2.BereitsGemeldetWurden(Sender: TObject);
var MyDBF : TDBF;
    OldFilterIndex, months, x, y : Integer;
    Betrag1, Betrag2 : Currency;
    aktuellerMonat, FolgeMonat, SQLFile, Filter, StrOr : string;
    MyDirSelect : TSelectDirectoryDialog;
    RIDFaks,RIDGeloescht, RIDMonat, MyFiles : TStringList;
begin
    (* bereits gemeldete Einnahmen zu Monat xy sind normalerweise in zwei Monatsdateien zu finden.
       Verkäufe vom Juni findet man also im Juni und im Juli.
       Das ist Folge von zu spät eingelesenen Druckern/Terminals
       Aus dem je betrachteten Monat darf nur der Datumsbereich von Interesse extrahiert werden,
       für den Juni 01.6.2015 bis 30.06.2015 *)

  (* ACHTUNG:
     es wird zwischen months=0 und months>0 im Code unterschieden!!!!!!!!!!!
     für months>0 siehe ungefähr Zeile 655 *)

   try

     Screen.ActiveControl.Invalidate;

     StoppIt := false;

     (* wird nur ein Monat oder mehrere Monate untersucht,
        für months>0 siehe ungefähr Zeile 635 *)
     months := MonthsBetween(DateEditVon.Date, DateEditBis.Date);

     StrOr := ' OR RID=';

     Filter := '';


     (* wird die RID-Werte der Monate aufnehmen *)
     RIDMonat := TStringList.Create;
     RIDMonat.Sorted:=true;

     (* wird die RID-Werte aus QFAKS aufnehmen *)
     RIDFaks := TStringList.Create;
     RIDFaks.Sorted:=true;

     (* welche Rids sind in den Dbase Dateien aber nicht mehr in QFaks? *)
     RIDGeloescht := TStringList.Create;
     RIDGeloescht.Sorted:=true;



     (* wird die Dateinamen der AmisData -DBFs aufnehmen *)
     MyFiles := TStringList.Create;

     (* Dbase Datei erzeugen, in die die Save Daten AmisData eingelesen werden *)
     MyDBF := TDBF.Create(Application);


     (* Dialog SelectDirectory erzeugen *)
     MyDirSelect := TSelectDirectoryDialog.Create(Application);

     if not DirectoryExists(SavePath) then
     begin

       (* einen halbwegs passenden Vorgabewert für Directory erzeugen,
          das tatsächliche Directory wird später in ini gespeichert *)
       MyDirSelect.Initialdir:=DirectoryEdit1.Directory;

       MyDirSelect.Title:='Wo stehen die bereits in Amisdata Schnittstelle BLE importierten Faks-Daten?';
       (* den Pfad der Dateien ermitteln *)
       if MyDirSelect.Execute then
         SavePath := IncludeTrailingBackslash(MyDirSelect.FileName)
       else
         exit;

       Application.ProcessMessages;
     end;

     if months = 0 then
     begin

       (* aktuellerMonat und FolgeMonat zusammensetzten *)
       //aktuellerMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy', incMonth(Date,-1)) + '.dbf';
       aktuellerMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy', DateEditVon.Date) + '.dbf';
       FolgeMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy', incMonth(DateEditVon.Date,1)) + '.dbf';




       if ((not FileExists(aktuellerMonat)) or
          (not FileExists(FolgeMonat))) then
       begin
         ShowMessage('Zumindest einer der zwei Monate:' + NL +
         ExtractFileName(aktuellerMonat) + NL +
         ExtractFileName(FolgeMonat) + NL + 'konnte in' + NL +
         SavePath + NL +' nicht gefunden werden. Machen Sie den Vergleich bitte mit Stat.exe selber!');
         exit;
       end;

     (* FolgeMonat untersuchen  *)
     Betrag1 := 0;
     MyDBF.TableName:= FolgeMonat;
     MyDBF.Open;
     MyDBF.First;
     ProgressBar1.Position:=0;
     ProgressBar1.Max:= MyDBF.RecordCount;

     StatusBar1.SimpleText:='ermittle die Einnahmen in: ''' + ExtractFileName(FolgeMonat) + '''' ;
     Application.ProcessMessages;

     //ShowMessage('Zwischen DateEditVon. und DateEditBis liegen ' + IntToStr(months) + ' Monate');  exit;

     while not MyDBF.EOF do
     begin
       jei;
       RIDMonat.Add(MyDBF.FieldByName('RID').AsString);
       (* Einnahme zusammnenzählen, falls sie in den Zeitraum fällt *)
       if ((MyDBF.FieldByName('DATUMV').AsDateTime >= DateEditVon.Date) and
          (MyDBF.FieldByName('DATUMV').AsDateTime <= DateEditBis.Date)) then
          begin
          Betrag1 := Betrag1 + MyDBF.FieldByName('PREIS').AsCurrency;
          end;

       Application.ProcessMessages;
       if StoppIt then
       begin
         Nei;
          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
         exit;
       end;


       MyDBF.Next;
       ProgressBar1.Position:= MyDBF.RecNo;


     end;
     MyDBF.Close;

     (* aktuellerMonat untersuchen  *)
     Betrag2 := 0;
     MyDBF.TableName:= aktuellerMonat;
     MyDBF.Open;
     MyDBF.First;
     ProgressBar1.Position:=0;
     ProgressBar1.Max:= MyDBF.RecordCount;

     StatusBar1.SimpleText:='ermittle die Einnahmen in: ''' + ExtractFileName(aktuellerMonat) + '''' ;
     Application.ProcessMessages;

     while not MyDBF.EOF do
     begin
       jei;
       (* RID in StringList eintragen *)
       RIDMonat.Add(MyDBF.FieldByName('RID').AsString);
       (* Einnahme zusammnenzählen, falls sie in den Zeitraum fällt *)
       if ((MyDBF.FieldByName('DATUMV').AsDateTime >= DateEditVon.Date) and
          (MyDBF.FieldByName('DATUMV').AsDateTime <= DateEditBis.Date)) then
          begin
          Betrag2 := Betrag2 + MyDBF.FieldByName('PREIS').AsCurrency;
          end;

       Application.ProcessMessages;
       if StoppIt then
       begin
         Nei;
          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
         exit;
       end;



       MyDBF.Next;
       ProgressBar1.Position:= MyDBF.RecNo;


     end;
     MyDBF.Close;



     Nei;

     (* Stimmen die ermittelten Beträge überein? *)
     if Betrag1 + Betrag2 = GesamtEinnahme then
       ShowMessage('Alles OK!' + NL +'''' + ExtractFileName(FolgeMonat) + ''' ( ' + FormatFloat('#,##0.00 EURO', Betrag1)+ ' )' +  NL +
                   '''' + ExtractFileName(aktuellerMonat) + ''' ( ' + FormatFloat('#,##0.00 EURO', Betrag2)+ ' )' +  NL
       + 'enthalten zusammen  ' + FormatFloat('#,##0.00 EURO', Betrag1 + Betrag2) + ' Einnahmen im Zeitraum.')
     else
       begin
       (* SQL-Code zur Fehlersuche vorbereiten und abspeichern
       ListSQL.Add('SELECT RID FROM "' + aktuellerMonat + '" WHERE DATUMV BETWEEN ''' + FormatDateTime('dd.mm.yyyy', DateEditVon.Date) + ''' AND ''' + FormatDateTime('dd.mm.yyyy', DateEditBis.Date) + '''' );
       ListSQL.Add('UNION ALL SELECT RID FROM "' + FolgeMonat + '" WHERE DATUMV BETWEEN ''' + FormatDateTime('dd.mm.yyyy', DateEditVon.Date) + ''' AND ''' + FormatDateTime('dd.mm.yyyy', DateEditBis.Date) + '''' );
       SQLFile := SavePath + 'FaksDaten_UNION_SQL_' + FormatDateTime('dd.mm.yyyy', DateEditVon.Date) + '.sql';
       ListSQL.SaveToFile(SQLFile);
       *)

       (* jetzt die Datensätze ermitteln, die der aktuelle Monat MEHR hat als die zwei Amisdata *.dbf
          Die zweite Möglichkeit: Datensätze sind in AmisData aber ncht mehr im aktuellen Monat untersuche ich nicht *)

       StatusBar1.SimpleText:='Prüfe auf nicht gemeldete RID''s ... bitte warten';
       Application.ProcessMessages;

       try
       Filter := '';
       BM := QFaks.GetBookmark;
       QFaks.DisableControls;
       ProgressBar1.Position:=0;
       QFaks.Last;
       ProgressBar1.Max:=QFaks.RecordCount;
       QFaks.First;
       jei;
       while not QFaks.EOF do
       begin
         if RIDMonat.IndexOf(QFaks.FieldByName('RID').AsString) = -1 then
           begin
             if Filter = '' then
               begin
                Filter :='RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
               end
             else
               begin
                 Filter := Filter + StrOr + QuotedStr(QFaks.FieldByName('RID').AsString);
               end;
           end;
         ProgressBar1.Position := QFaks.RecNo;
         QFaks.Next;

         Application.ProcessMessages;
         if StoppIt then
         begin
           Nei;
            ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
           exit;
         end;

       end;
       finally
       Nei;
       if QFaks.BookmarkValid(BM) then QFaks.GotoBookmark(BM);
       QFaks.FreeBookmark(BM);
       QFaks.EnableControls;
       QFaks.Refresh;
       ProgressBar1.Position := 0;
       end;

       mem.Clear;
       mem.Text := Filter;
       mem.SelectAll;
       mem.CopyToClipboard;

       FilterCombo.Text:= Filter;
       QFaks.Filter := '(' + Filter + ') And Vertragsnr is null';
       QFaks.Filtered := True;
       ShowFilterInfo(True);



       ShowMessage('Da stimmt was nicht! ' + NL +
       FormatFloat('#,##0.00 EURO', GesamtEinnahme) + NL + 'wurden in diesem Programm eben ermittelt!' + NL + 'Aber:' + NL +
       '''' + ExtractFileName(FolgeMonat) + ''' ( ' + FormatFloat('#,##0.00 EURO', Betrag1)+ ' )' +  NL +
                   '''' + ExtractFileName(aktuellerMonat) + ''' ( ' + FormatFloat('#,##0.00 EURO', Betrag2)+ ' )' +  NL
       + 'enthalten zusammen  ' + FormatFloat('#,##0.00 EURO', Betrag1 + Betrag2) + NL + ' Einnahmen im Zeitraum.' + NL +
       'Differenzbetrag = ' + FormatFloat('#,##0.00 EURO', Betrag1 + Betrag2 - GesamtEinnahme) + NL + NL + 'Strg + c kopiert diesen Text!' + NL + NL +
       'Die NICHT gemeldeten Daten werden jetzt angezeigt. Der verwendete Filter wurde kopiert.');

       end;
     FreeAndNil(MyDirSelect);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
     end (* months=0 *)
     else
     begin
       (* jetzt alle Monate durchlaufen, die RID's sammeln und in FAKS nachsehen ob vorhanden *)
       RIDMonat.Clear;

       Betrag2 := 0;
       Filter := '';

       for x := 0 to months do
       begin
         aktuellerMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy', incMonth(DateEDitVon.Date,x)) + '.dbf';
         if not FileExists(aktuellerMonat) then
           begin
             ShowMessage('Datei nicht gefunden, Sie müssen das ggf. händisch erledigen!' + NL + NL + aktuellerMonat );
             continue;
           end;


// months durchlaufen *******************************************
          if MyDBF.Active then MyDBF.Close;

          (* aktuellerMonat untersuchen  *)
          MyDBF.TableName:= aktuellerMonat;
          MyDBF.Open;
          MyDBF.First;
          ProgressBar1.Position:=0;
          ProgressBar1.Max:= MyDBF.RecordCount;

          StatusBar1.SimpleText:='ermittle die Einnahmen in: ''' + ExtractFileName(aktuellerMonat) + '''' ;
          Application.ProcessMessages;

          Betrag1 := 0;



          while not MyDBF.EOF do
          begin
            jei;
            (* RID in StringList eintragen *)
            RIDMonat.Add(MyDBF.FieldByName('RID').AsString);

            (* Einnahme zusammnenzählen, falls sie in den Zeitraum fällt *)
            if ((MyDBF.FieldByName('DATUMV').AsDateTime >= DateEditVon.Date) and
               (MyDBF.FieldByName('DATUMV').AsDateTime <= DateEditBis.Date)) then
               begin
               Betrag2 := Betrag2 + MyDBF.FieldByName('PREIS').AsCurrency;
               Betrag1 := Betrag1 + MyDBF.FieldByName('PREIS').AsCurrency;

               (* testen, ob Feld Betrag in QFaks den gleichen Wert hat *)
               (*
               if  QFaks.Locate('RID',VarArrayOf([MyDBF.FieldByName('RID').AsString]),[]) then
                 begin
                   if not (QFaks.FieldByName('Betrag').AsCurrency = MyDBF.FieldByName('PREIS').AsCurrency) then
                     begin
                       EventLog1.Log('QFaks.Betrag <> DBF.Preis bei RID: ' + MyDBF.FieldByName('RID').AsString);
                       ShowMessage('Betrag: ' + QFaks.FieldByName('Betrag').AsString + ' <> PREIS: '+ MyDBF.FieldByName('PREIS').AsString + ' bei RID: ' + MyDBF.FieldByName('RID').AsString);
                     end;
                 end;
               *)

               end;

            MyDBF.Next;
            ProgressBar1.Position:= MyDBF.RecNo;

            Application.ProcessMessages;
            if StoppIt then
            begin
              Nei;
               ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
              exit;
            end;



          end;
          MyDBF.Close;

          (* wird mit Dialog später angezeigt *)
          MyFiles.Add(aktuellerMonat + ' [Summe im Zeitraum: ' + FormatDateTime('dd.mm.yy', DateEditVon.Date) + ' bis ' +
                                                                FormatDateTime('dd.mm.yy', DateEditBis.Date)  + '] -> ' + FormatFloat('#,##0.00', Betrag1));


       end (* for x *);


          Nei;

          (* Stimmen die ermittelten Beträge überein? *)
          if Betrag2 = GesamtEinnahme then
            ShowMessage('Alles OK!' + NL + NL + MyFiles.Text +  NL + NL +
             'enthalten zusammen  ' + NL +  FormatFloat('#,##0.00 EURO',  Betrag2) + ' Einnahmen im Zeitraum.')
          else
            begin

            (* jetzt die Datensätze ermitteln, die der aktuelle Monat MEHR hat als die zwei Amisdata *.dbf
               Die zweite Möglichkeit: Datensätze sind in AmisData aber ncht mehr im aktuellen Monat untersuche ich nicht *)
            StatusBar1.SimpleText:='Prüfe auf nicht gemeldete RID''s ... bitte warten: ' + IntToStr(RIDMonat.Count) + ' Datensätze in den Dbase Dateien';
            Application.ProcessMessages;
            Jei;

            try
            BM := QFaks.GetBookmark;
            QFaks.DisableControls;
            ProgressBar1.Position:=0;
            QFaks.Last;
            ProgressBar1.Max:=QFaks.RecordCount;
            QFaks.First;
            StrOr := ' OR RID=';
            Filter := '';
            jei;
            Application.ProcessMessages;
            while not QFaks.EOF do
            begin

              (* RID's aus QFaks speichern *)
              RIDFaks.Add(QFaks.FieldByName('RID').AsString);

              y := RIDMonat.IndexOf(QFaks.FieldByName('RID').AsString);
              if y = -1 then
                begin
                  if Filter = '' then
                    begin
                     Filter :='RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
                    end
                  else
                    begin
                      Filter := Filter + StrOr + QuotedStr(QFaks.FieldByName('RID').AsString);
                    end;
                end;
                //else
                //(* was gefunden wurde, kann aus RIDMonat gelöscht werden *)
                //RIDMonat.Delete(y);

              ProgressBar1.Position := QFaks.RecNo;
              Application.ProcessMessages;
              if StoppIt then
              begin
                Nei;
                 ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
                exit;
              end;
              QFaks.Next;
            end;
            finally
            Nei;
            if QFaks.BookmarkValid(BM) then QFaks.GotoBookmark(BM);
            QFaks.FreeBookmark(BM);
            QFaks.EnableControls;
            QFaks.Refresh;
            ProgressBar1.Position := 0;
            end;

            jei;
            mem.Clear;
            mem.Text := Filter;
            mem.SelectAll;
            mem.CopyToClipboard;

            FilterCombo.Text:= Filter;
            StatusBar1.SimpleText:=StatusBar1.SimpleText + ' Jetzt wird der Filter gesetzt';
            QFaks.Filter := '(' + Filter + ') And Vertragsnr=''1''';
            QFaks.Filtered := True;
            ShowFilterInfo(True);

            (* nachsehen, welche Datensätze aus Faks gelöscht wurden *)
            for x := 0 to RIDMonat.Count -1 do
            begin
              if RIDFaks.IndexOf(RIDMonat[x]) = -1 then
              begin
                RIDGeloescht.Add(RIDMonat[x]);
              end;

            end;

            (* nach Spalte Vetragsnr sortieren *)
            QFaks.SortedFields:='Vertragsnr';
            QFaks.First;

            (* werden die angezeigten Daten bei der nächsten Meldung geholt? *)

            if ((GesamtEinnahme = Betrag2) OR (QFaks.RecordCount = 0)) then
               begin
                  ShowMessage('Scheint alles zu stimmen' + NL +
                  'Eben ermittelt wurden: ' + FormatFloat('#,##0.00 EURO', GesamtEinnahme) + NL +
                  'Summe in den DBase Dateien Amisdata: ' + FormatFloat('#,##0.00 EURO',  Betrag2) + NL + NL +
                  'Achtung: noch an AmisData zu meldende Daten (Vertragsnr is null) werden nicht angezeigt, sind aber in den Summen enthalten!');
               end
            else
            begin

            ShowMessage('Da stimmt was nicht! ' + NL +
            FormatFloat('#,##0.00 EURO', GesamtEinnahme) + NL + 'wurden in diesem Programm eben ermittelt!' + NL + 'Aber:' + NL +
            '' + MyFiles.Text  +  NL + 'enthalten zusammen:' + NL + FormatFloat('#,##0.00 EURO',  Betrag2) + NL + 'Einnahmen im Zeitraum.' + NL +
            'Differenzbetrag = ' + NL + FormatFloat('#,##0.00 EURO',  Betrag2 - GesamtEinnahme) + NL + NL + 'Strg + c kopiert diesen Text!' + NL + NL +
            'Die NICHT gemeldeten Daten werden jetzt angezeigt. Der verwendete Filter wurde kopiert.' + NL + NL +
            'Achtung: es kam schon vor das Daten in Faks nachträglich gelöscht wurden!!' + NL + NL +
            'Achtung auf noch nicht gemeldete Daten, siehe Spalte Vertragsnr (1=gemeldet)' + NL + NL +
            'Bitte einen möglichst grossen Zeitraum wählen: 01.01.' + IntToStr(YearOf(DateEDitBis.Date))+ ' bis ' + DateEDitBis.Text);

            if RIDGeloescht.Count > 0 then ShowMessage('Diese RID wurden nachträglich aus FAKS gelöscht: ' + sLineBreak + RIDGeloescht.Text);
            end;
        end;
// *******************************************



     end;

   finally
    Nei;
    StatusBar1.SimpleText:='';
    ProgressBar1.Position:=0;
    FreeAndNil(MyDBF);
    FreeAndNil(RIDMonat);
    FreeAndNil(MyFiles);
    FreeAndNil(RIDFaks);
    FreeAndNil(RIDGeloescht);
   end;

end;

procedure TForm2.CheckIniFile(Sender: TObject);
var ini1, ini2 : TIniFile;
    List1, List2 : TStringList;
    FName, Section : string;
begin
   (* kontrollieren, ob LastBuchungsdatum auch in
      Amisdata-FAKS_Meldung.ini steht und ggf. nachtragen  *)
   FName := 'F:\AMISdata\Monat\Monatsmld\FAKS_Meldung.ini';
   Section := 'TApplication.Form2.Edit_last_buchung_Items';

   (* nur Kontrolle wenn lokal ausgeführt wird und AmisData vorhanden ist *)
   if ((Pos('PROJEKTE',ExePath) > 0) and (FileExists(FName))) then
   begin
     try
       List1 := TStringList.Create;
       List2 := TStringList.Create;

       (* aktuelle ini Datei speichern *)
       INIPropStorage1.Save;

       (* ini der laufenden Anwendung *)
       ini1 := TInifile.Create(INIPropStorage1.IniFileName);
       (* ini der Anwendung unter AmisData *)
       ini2 := TInifile.Create(FName);

       ini1.ReadSectionValues(Section,List1);
       ini2.ReadSectionValues(Section,List2);


      if not SameText(List1.CommaText, List2.CommaText) then
      begin
          (* wenn die sections nicht gleich sind, ini der lokalen Anwendung kopieren *)
          Mem.Clear;
          Mem.Text := List1.Text;
          (* Section in eckigen Klammern oben drüber *)
          Mem.Text := '['+Section +']' + NL + Mem.Text;
          Mem.SelectAll;
          Mem.CopyToClipboard;

          ShowMessage('Die ini-Einträge sind nicht gleich:' + NL +
          ini1.FileName + NL + List1.CommaText + NL + NL +
          ini2.FileName + NL + List2.CommaText + NL + NL + 'Werte wurden kopiert! ZWEI(!!) Editoren werden geöffnet.');

          (* ini im Anwendungs-Verzeichnis im Editor öffnen *)
          shellExecute(Application.MainForm.Handle,'open',PChar(INIPropStorage1.IniFileName) ,'','',SW_NORMAL);
          (* ini in Amisdata-Verzeichnis im Editor öffnen *)
          shellExecute(Application.MainForm.Handle,'open',PChar(FName) ,'','',SW_NORMAL);


      end;



     finally
      freeAndNil(List1);
      freeAndNil(List2);
      freeAndNil(Ini1);
      freeAndNil(Ini2);


     end;

   end;

end;




procedure TForm2.angezeigteDatenkopieren1Click(Sender: TObject);
var
  {* Siehe auch:
  http://www.delphi3000.com/articles/article_2292.asp?SK=dbgrid *}
  col, row: integer;
  sline: string;
  mem1: TStringList;
  //ExcelApp: Variant;
  //Q : TZReadOnlyQuery;
begin
  try
  (* globale Stopp Variable, die nach ESCAPE-Taste auf true gesetzt wird *)
  StoppIt := false;

  mem1:= TStringList.Create;


  (* Dieser Code wird  für QFaks und QUmsetzung benutzt *)
  //Q := (screen.ActiveControl as TDBGrid).DataSource.DataSet as TZReadOnlyQuery;

  Screen.Cursor := crHourglass;
  QFaks.DisableControls;
  bm := QFaks.GetBookmark;
  QFaks.Last;
  QFaks.First;

  ProgressBar1.Position:=0;
  ProgressBar1.Max:=QFaks.RecordCount;





  // First we send the data to a memo
  // works faster than doing it directly to Excel

  //das geht bei lazarus NICHT!!!  mem := TMemo.Create(Self);
  //mem.Parent := Form2;

  Form2.mem.Visible := True;
  Form2.mem.Clear;
  sline := '';

  // add the info for the column names
  for col := 0 to QFaks.FieldCount - 1 do  sline := sline + QFaks.Fields[col].DisplayLabel + #9;

  (* letzten Tabulator entfernen *)
  sline:= copy(sline,0,length(sline) -1);

  mem1.Add(sline);

  (* gibts MultiSelect im Grid oder ganzes Grid kopieren *)
  if (screen.ActiveControl as TDBGrid).SelectedRows.Count > 1 then
  begin
    (* also Multiselect! *)
   statusbar1.SimpleText :='Das DBGrid ''' + (screen.ActiveControl as TDBGrid).Name + ''' hat ' +
   IntToStr((screen.ActiveControl as TDBGrid).SelectedRows.Count) + ' Zeilen selektiert!';

   // get the data into the memo
   for row := 0 to (screen.ActiveControl as TDBGrid).SelectedRows.Count - 1 do
   begin
     sline := '';
     QFaks.GotoBookmark(TBookMark((screen.ActiveControl as TDBGrid).SelectedRows[row]));;
     for col := 0 to QFaks.FieldCount - 1 do
       sline := sline + QFaks.Fields[col].AsString + #9;

     (* letzten Tabulator rntfernen *)
     sline:= copy(sline,0,length(sline) -1);

     mem1.Add(sline);
     //QFaks.Next;
   end;


  end
  else
  begin
    //if Messagedlg('Alle Daten sollen kopiert werden?' + NL + NL +
    //   'Durch ESC-Taste abbrechen!',mtConfirmation,[mbYes,mbNo],0)= mrNo then exit;
    (* kein Multiselect, also ganzes Grid *)
    statusbar1.SimpleText :='Das ganze Grid wird jetzt kopiert!';
    // get the data into the memo
    for row := 0 to QFaks.RecordCount - 1 do
    begin
      ProgressBar1.Position:=QFaks.RecNo;
      Application.ProcessMessages;
      if StoppIt then
      begin

       ProgressBar1.Position:=0;
       Application.ProcessMessages;

       ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
       StoppIt := false;
       exit;
      end;
      sline := '';
      for col := 0 to QFaks.FieldCount - 1 do
        sline := sline + QFaks.Fields[col].AsString + #9;

      (* letzten Tabulator entfernen *)
      sline:= copy(sline,0,length(sline) -1);


      mem1.Add(sline);
      QFaks.Next;
    end;

  end;


  // we copy the data to the clipboard
  Form2.mem.Text := mem1.Text;
  Form2.mem.SelectAll;
  Form2.mem.CopyToClipboard;
  Form2.mem.Clear;
  Form2.mem.Visible := False;



  ProgressBar1.Position:=0;
  Application.ProcessMessages;

  ShowMessage('Die Daten wurden kopiert und können z.B in Excel eingefügt werden.');

  finally
  FreeAndNil(mem1);
  QFaks.GotoBookmark(bm);
  QFaks.FreeBookmark(bm);
  QFaks.EnableControls;
  Screen.Cursor := crDefault;
  end;

end;

function TForm2.OpenLog(FName: string): boolean;
begin
  try

    (* mit Windows Notepad öffnen *)
      {$IFDEF MSWINDOWS}

    AsyncProcess1.CommandLine := 'notepad.exe "' + FName + '"';
    AsyncProcess1.Active := True;
     {$ENDIF}
    Result := True;

  finally
  end;

end;

procedure TForm2.ShortenLog(anzLines: integer; Sender: TObject);
var
  log: TStringList;
  x: integer;
begin
  try
    (* LogFile auf MaxLogLines Zeilen kürzen *)
    log := TStringList.Create;

    if not FileExists(ZSQLMonitor1.FileName) then
      exit;


    ZSQLMonitor1.Active := False;

    log.LoadFromFile(ZSQLMonitor1.FileName);


    if log.Count > anzLines then
    begin

      (* damit überhaupt eingelesen werden kann *)
      ZSQLMonitor1.Active := False;
      log.LoadFromFile(ZSQLMonitor1.FileName);


      while log.Count >= anzLines  do
        log.Delete(0);

      log.SaveToFile(ZSQLMonitor1.FileName);

      ZSQLMonitor1.Active := true;

    end;

  finally
    FreeAndNil(log);
  end;

end;


procedure TForm2.FormCreate(Sender: TObject);
var lst : TStringlist;
      x : integer;
      OldDelimiter, s : string;
begin

  ExePath := ExtractFilePath(Application.ExeName);
  INIPropStorage1.IniFileName := ChangeFileExt(Application.ExeName, '.ini');
  ZSQLMonitor1.FileName := ChangeFileExt(Application.ExeName, '.log');

  (* speichert die History der ausgeführten SQL-Befehle *)
  SQLHistory := TStringList.Create;



  (* Log für Dbase-Exporte,
     wichtig: LogType muss ltFile sein, sonst wird in Systemlog geschrieben *)
  EventLog1.FileName:= ChangeFileExt(Application.ExeName, '_DBase_Exporte.log');
  EventLog1.AppendContent:=true;
  EventLog1.LogType:=ltFile;
  EventLog1.Active:=true;

  (* gibts die tnsnames.ora, sonst Oracle Fehlermeldung *)
  if not FileExists(ExePath + 'tnsnames.ora') then
  begin
    ListBox_tnsnames_ora.Items.SaveToFile(ExePath + 'tnsnames.ora');
    ShowMessage('Die Oracle Konfigurationsdatei ''' + ExePath +'tnsnames.ora'' konnte nicht gefunden werden und wurde neu erzeugt. Einträge ggf. überprüfen!!');
  end;



end;

procedure TForm2.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
  );
begin
    if ((Key = VK_F3)  and (gesucht <> '')) then
    begin
     Pm_SearchClick(Sender);
    end
    else  if (Key = VK_ESCAPE)  then
    begin
      StoppIt := true;
      //raise exception.create('Hurrah Sie haben die ESCAPE-Taste getroffen :-)) ');
    end


end;

procedure TForm2.IniPropStorage1RestoreProperties(Sender: TObject);
begin
  TarifVersion:= SpinEditTarifversion.Value;
end;

procedure TForm2.IniPropStorage1StoredValues1Restore(Sender: TStoredValue;
  var Value: TStoredType);
begin
  SavePath := Value;
end;

procedure TForm2.IniPropStorage1StoredValues1Save(Sender: TStoredValue;
  var Value: TStoredType);
begin
  (* in ini Speichern Amisdata Pfad Save *)
   Value := IncludeTrailingBackslash(SavePath);
end;

procedure TForm2.kopierenClick(Sender: TObject);
begin
  angezeigteDatenkopieren1Click(Sender);
end;

procedure TForm2.ListBoxKnownLinesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_DELETE then
    ListBoxKnownLines.DeleteSelected;

end;

procedure TForm2.ListFieldsClick(Sender: TObject);
var x : integer;
begin
  try
   Jei;
  // ZSQLMetadata1.TableName:='f2fsv';
  // ZSQLMetadata1.Active:=true;
  // for x := 0 to ZSQLMetadata1.FieldCount -1 do
  // begin
  //   if CompareText(ZSQLMetadata1.Fields[x].FieldName,'BETRAG')= 0  then
  //    ShowMessage(ZSQLMetadata1.Fields[x].FieldName)
  //    else
  //    ShowMessage(ZSQLMetadata1.Fields[x].FieldName);
  //
  //
  // end;
  //
  //if ExistInDb('f2fsv') then  ShowMessage('f2fsv gefunden');

  if ExistField('journal') then  ShowMessage('Feld Journal wurde gefunden');

  finally
   Nei;

  end;
end;

procedure TForm2.LookupStornosClick(Sender: TObject);
var Filter : string;
      x, y : integer;
      Betrag : Currency;
begin
  try
    if Messagedlg('Diese Aktion wird länger dauern, kann aber im ersten Teil durch ESC-Taste abgebrochen werden.',mtConfirmation,[mbYes,mbNo],0)= mrNo then exit;
    Jei;
    StoppIt := false;


    BM := QFaks.GetBookmark;
    QFaks.DisableControls;


    QFaks.SortedFields:='Betrag';
    QFaks.Last;
    QFaks.First;
    y := QFaks.RecordCount;


    Filter := '';

    while  QFaks.FieldByName('Betrag').AsCurrency < 0 do
    begin

      Application.ProcessMessages;
      if StoppIt then
      begin
       ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
       StoppIt := false;
       exit;
      end;

      (* Feld BelegNr UND Bemerkung sind bei Stornos identisch *)
  //(MDEIDINTERN='5012' AND (BELEGNR='7141' OR Bemerkung='7141')) OR (MDEIDINTERN='5005' AND (BELEGNR='6637' OR Bemerkung='6637'))
      if Filter = '' then
        begin
         Filter :='(MDEIDINTERN=' + QuotedStr(QFaks.FieldByName('MDEIDINTERN').AsString) +
         ' AND (BELEGNR=' + QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + ' OR Bemerkung=' + QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + '))' ;
        end
      else
        begin
          Filter := Filter + ' OR '  + '(MDEIDINTERN=' + QuotedStr(QFaks.FieldByName('MDEIDINTERN').AsString) +
          ' AND (BELEGNR=' + QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + ' OR Bemerkung=' + QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + '))' ;
        end;



      QFaks.Next;
    end;

    mem.Clear;
    mem.Text := Filter;
    mem.SelectAll;
    mem.CopyToClipboard;
    //mem.Lines.SaveToFile(ExePath + 'MyFilter.txt');

    //ShowMessage(Filter);

    FilterCombo.Text:= Filter;
    QFaks.Filter := Filter;
    QFaks.Filtered := True;
    ShowFilterInfo(True);

    x := QFaks.RecordCount;

    QFaks.SortedFields:='Belegnr, Bemerkung';

    (* Welchen Wert hatten die stornierten Fahrscheine? *)
    QFaks.First;
    Betrag := 0;
    While not QFaks.EOF do
    begin
      If QFaks.FieldByName('Betrag').AsCurrency > 0 then
      Betrag := Betrag + QFaks.FieldByName('Betrag').AsCurrency;
      QFaks.Next;
    end;


  finally
    Nei;
    if QFaks.BookmarkValid(BM)then QFaks.GotoBookmark(BM);
    QFaks.EnableControls;

    ShowMessage('Im Zeitraum: ' + DateEditVon.Text + ' bis ' + DateEditBis.Text + ' wurden:' + NL  +
    FormatFloat('#,##0',x/2) + ' von ' + IntToStr(y) + ' Fahrscheinen storniert.' + NL +
    'Das entspricht ' + FormatFloat('#,##0.00 %',((x/2)/y) * 100) + NL +
    'Summe: ' + FormatFloat('#,##0.00',Betrag));


  end;
end;

procedure TForm2.MenuAddLinieClick(Sender: TObject);
var UserString : string;
begin
  UserString := InputBox('Neue Liniennummer erfassen:',
      'die neue Liniennummer ist:', '0');

  if ((UserString <> '0') and (ListBoxKnownLines.Items.IndexOf(UserString) = -1)) then
   ListBoxKnownLines.Items.add(UserString);

end;

procedure TForm2.MenuItem1Click(Sender: TObject);
begin
  if DirectoryExists(DirectoryEdit1.Directory) then
  begin
    OpenExplorer(DirectoryEdit1.Directory);
  end
  else
  begin
    ShowMessage('Bitte das Verzeichnis selber auswählen!');
    OpenExplorer(ExePath);
  end;
end;

(* wird nicht mehr benutzt: Januar bis August 2016 Exporte setzen
   laut 'F:\AMISdata\Monat\Monatsmld\FAKS_Auswertungen\RID_as_imported.dbf' *)
procedure TForm2.MenuItem2Click(Sender: TObject);
var tab : TDBF;
    Betrag : currency;
    v : variant;
    ErrLIst : TstringList;
begin
  try
  jei;

  ErrLIst := TstringList.Create;


  tab := TDBF.Create(Application.Mainform);
  tab.TableName:='F:\AMISdata\Monat\Monatsmld\FAKS_Auswertungen\RID_as_imported.dbf';
  tab.Open;
  tab.First;

  QFaks.DisableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Max:=Tab.RecordCount;
  while not tab.EOF do
  begin
     v := Trim(tab.FieldByName('RID').AsString);
     ZUpDateRid.ParamByName('RID').AsString:=v;
     ZUpDateRid.ParamByName('Vertragsnr').AsString:='1';
     ZUpDateRid.ExecSQL;
     //if not QFaks.Locate('RID',v,[]) then ErrList.Add(v);
     tab.next;
     ProgressBar1.Position:=Tab.RecNo;
     ProgressBar1.Update;
  end;

  if ErrList.Count> 1 then
  begin
    ErrList.SaveToFile(EXEPath + 'ErrList.txt');
     ShowMessage('Siehe '+ EXEPath + 'ErrList.txt' );
  end;

  finally
    tab.close;
    FreeAndNil(tab);
    FreeAndNil(ErrLIst);
    QFaks.EnableControls;
    QFaks.Refresh;
    nei;
  end;

end;

procedure TForm2.MnMarkExportedClick(Sender: TObject);
var FName, mark: string;
    MyDirSelect : TSelectDirectoryDialog;
begin

  if Messagedlg('Das ist sehr gefährlich!! Sie werden bereits verarbeitete Daten ändern! Wollen Sie wirklich fortfahren?',mtConfirmation,[mbYes,mbNo],0)= mrNo then  exit;

  FName := SavePath;

  if InputQuery('Bereits in AmisData importierte Dbf-Datei',
     'Pfad und Dateiname der dbf aus dem Ordner: '+ FName, FName) = false then exit
  else
  begin
    StatusBar1.SimpleText:= 'bearbeite die Exportmarkierung für: ''' + FName + '''';
    Application.ProcessMessages;
    if InputQuery('Was soll in Spalte VertragsNr eingetragen werden?',
       '1 = wurde exportiert ''''= noch nicht exportiert', mark) = false then exit;

    //ShowMessage(mark);

    (* Dialog SelectDirectory erzeugen *)
    MyDirSelect := TSelectDirectoryDialog.Create(Application);

    if not DirectoryExists(SavePath) then
    begin

      (* einen halbwegs passenden Vorgabewert für Directory erzeugen,
         das tatsächliche Directory wird später in ini gespeichert *)
      MyDirSelect.Initialdir:=DirectoryEdit1.Directory;

      MyDirSelect.Title:='Wo stehen die bereits in Amisdata Schnittstelle BLE importierten Faks-Daten?';
      (* den Pfad der Dateien ermitteln *)
      if MyDirSelect.Execute then
        SavePath := IncludeTrailingBackslash(MyDirSelect.FileName)
      else
        exit;

      Application.ProcessMessages;
    end;

    FreeAndNil(MyDirSelect);

    MarkExported(FName,mark);
  end;
end;

procedure TForm2.OpenLogFileClick(Sender: TObject);
begin
  OpenLog(ZSQLMonitor1.FileName);
end;

procedure TForm2.Panel7DblClick(Sender: TObject);
begin
  ShowMessage(FormatDateTime(ShortDateFormat+' / '+ShortTimeFormat, now));
end;

procedure TForm2.Pm_SearchClick(Sender: TObject);
var gefunden : boolean;
begin
  if Sender <> Form2 then
  begin

    if DBGridFaks.SelectedField.DataType = ftFloat then
    begin
      gesucht := DBGridFaks.SelectedField.AsFloat;
    end
    else
    begin
      gesucht := DBGridFaks.SelectedField.AsString;
    end;

    gesucht := InputBox('Welche exakte Zeichenfolge soll ab aktueller Position gesucht werden in Spalte ''' +
     DBGridFaks.SelectedField.FieldName + '''?','Suchbegriff ist (F3=Suche fortsetzen!):',gesucht);

    StatusBar1.SimpleText:='gesucht wird nach Textbestandteil ''' + gesucht + ''' in Spalte ''' + DBGridFaks.SelectedField.FieldName + '''';
  end;

  try
     Jei;
     gefunden := false;
     BM := QFaks.GetBookmark;

     if  not QFaks.EOF then QFaks.Next;

     while not QFaks.EOF do
     begin
        (* auch Teilzeichenfolge caseinsensitive finden  *)
        //if not  AnsiContainsText(DBGridFaks.SelectedField.AsString ,gesucht) then  QFaks.Next

       (* in Abhängigkeit vom FeldTyp suchen *)
       if DBGridFaks.SelectedField.DataType = ftFloat then
       begin
          if not  (DBGridFaks.SelectedField.AsFloat = gesucht)  then  QFaks.Next
          else
          begin
           gefunden := true;
           break;
          end;
       end
       else
       begin
         if not  (DBGridFaks.SelectedField.AsString = gesucht)  then  QFaks.Next
         else
         begin
          gefunden := true;
          break;
         end;
       end;
     end;

     Nei;

     if not gefunden then
     begin
      QFaks.GotoBookmark(BM);
      ShowMessage('''' + VarToStr(gesucht) +
      ''' konnte in Spalte ''' + DBGridFaks.SelectedField.Fieldname +  ''' nicht gefunden werden!' );
      StatusBar1.SimpleText := '';
     end;

  finally
    QFaks.FreeBookmark(BM);

  end;
end;

procedure TForm2.RemoveFilterClick(Sender: TObject);
begin
  QFaks.Filtered := False;
  QFaks.Filter := '';
  FilterCombo.Text := QFaks.Filter;
  StatusBar1.SimpleText := '';

  ShowFilterInfo(false);

end;

procedure TForm2.RID_as_FilterClick(Sender: TObject);
var Filter, StrOr : string;
      x : integer;
begin
   StrOr := ' OR RID=';
   try
     jei;
     QFaks.DisableControls;
     BM := QFaks.GetBookmark;

     if not QFaks.Filtered then
     if Messagedlg('Die Daten sind ungefiltert, der Filterausdruck für die RID kann riesig werden!' + NL + NL +
     'Wollen Sie wirklich fortfahren?',mtConfirmation,[mbYes,mbNo],0)= mrNo then  exit;

     QFaks.First;
     Filter :='RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
     While not QFaks.EOF do
     begin
       QFaks.Next;
       Filter := Filter + StrOr + QuotedStr(QFaks.FieldByName('RID').AsString);
     end;

     mem.Clear;
     mem.Text := Filter;
     mem.SelectAll;
     mem.CopyToClipboard;

     QFaks.GotoBookmark(BM);

     ShowMessage('Der erstellte RID-Filter wurde in die Zwischenablage kopiert und könnte anderswo per Strg + V eingefügt werden.');
   finally
     Nei;
     QFaks.EnableControls;
     QFaks.FreeBookmark(BM);


   end;



end;

procedure TForm2.SaveToFileClick(Sender: TObject);
begin
  SaveDialog1.InitialDir := ExePath;
  if SaveDialog1.Execute then
  begin
    Memo1.Lines.SaveToFile(SaveDialog1.FileName);
  end;
end;

procedure TForm2.SpinEdit1Change(Sender: TObject);
var D,M,Y : Word;

begin
  try

    VonDatum := incMonth(DateEditVon.Date,SpinEdit1.Value);
    DecodeDate(VonDatum,Y,M,D);
    BisDatum := EndOfAMonth(Y,M);

    StatusBar1.SimpleText := 'VonDatum: ' + FormatDateTime('dd.mm.yyyy',VonDatum) + ' '  +
                  'Bisdatum: ' + FormatDateTime('dd.mm.yyyy',BisDatum) ;


  finally

  end;
end;

procedure TForm2.SpinEditTarifversionEditingDone(Sender: TObject);
begin
   Memo1.SetFocus;
   ShowMessage('Bitte neu starten!');
end;

procedure TForm2.SQLLoadClick(Sender: TObject);
begin
  OpenDialog1.InitialDir := ExePath;
  if OpenDialog1.Execute then
  begin
    Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
  end;
end;

procedure TForm2.SQL_history_lodeClick(Sender: TObject);
var line : String;
begin
  if SQLHistory.Count > 0 then
  begin
    dec(SQLHistoryIndex);

    if SQLHistoryIndex > -1 then
      StrToStrings(SQLHistory[SQLHistoryIndex],'°',Memo1.Lines,true)
    else
    begin
      ShowMessage('Der erste Eintrag der History wurde bereits geladen!' + NL +
      'Jetzt wird wieder der höchste Eintrag der Liste angezeigt');
      SQLHistoryIndex := SQLHistory.Count -1;
      StrToStrings(SQLHistory[SQLHistoryIndex],'°',Memo1.Lines,true)

    end;

    StatusBar1.SimpleText:='Die SQL-History enthält: ' +
      IntToStr(SQLHistory.Count) + ' Einträge, angezeigt wird jetzt der Eintrag ' +
      IntToStr(SQLHistoryIndex +1);
  end
  else
   ShowMessage('Die SQL-History enthält noch keine Abfragen');
end;

procedure TForm2.UniqueInstance1OtherInstance(Sender: TObject;
  ParamCount: integer; Parameters: array of string);
begin
  if WindowState = wsMinimized then
    Application.Restore
  else
  begin
    BringToFront;
    SetFocus;
  end;
end;

procedure TForm2.QFaksAfterOpen(DataSet: TDataSet);
var x : integer;
    F : TFloatField;
begin
  Memo1.Lines.Assign(QFaks.SQL);
  DBGridFaks.AutoSizeColumns;

  (* dafür sorgen, dass nicht 1,0999999999 statt 1,10 EURO angezeigt werden *)
  //QFaks.FieldDefs.Update;
  for x := 0 to QFaks.FieldCount -1 do
  begin
    if QFaks.Fields[x].DataType = ftFloat then
    begin
     F := (QFaks.Fields[x] as TFloatField);
     //ShowMessage(QFaks.Fields[x].FieldName);
     F.Precision := 15;
     F.DisplayFormat:='#,##0.00';
     //QFaks.FieldDefs.Update;
    end;
  end;



end;

procedure TForm2.ApplicationProperties1Hint(Sender: TObject);
begin
  StatusBar1.SimpleText := Application.Hint;
end;

procedure TForm2.AuswahlFilterClick(Sender: TObject);
var
  Filter : string;
  row : integer;
begin
  //if QFaks.Filtered then QFaks.Filtered := False;

  (* gibts MultiSelect im Grid oder ganzes Grid kopieren *)
  if DBGridFaks.SelectedRows.Count > 1 then
  begin
    (* also Multiselect! *)
   statusbar1.SimpleText :='Das DBGrid ''' + DBGridFaks.Name + ''' hat ' +
   IntToStr(DBGridFaks.SelectedRows.Count) + ' Zeilen selektiert!';


   // die RID speichern
   for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
   begin

     QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));;

     if row = 0 then
       Filter := 'RID=''' + QFaks.FieldByName('RID').AsString + ''''
     else
       Filter := Filter + ' OR RID= ''' +  QFaks.FieldByName('RID').AsString + '''';

   end;




  end
  else
  begin


  (* Filter auf das Feld Zeit geht nicht, aber DATUMZEIT enthält ja die gleiche Information für die Zeit, deshalb: *)
  if DBGridFaks.SelectedField.FieldName = 'ZEIT' then
  begin
    DBGridFaks.SelectedField := QFaks.FieldByName('DATUMZEIT');
    ShowMessage(
      'Das Feld ZEIT läßt sich nicht filtern, alternativ nehmen wir das Feld DATUMZEIT, das aber auch das Datum berücksichtigt!');
  end;

  (* ggf. vorhandenen Filter durch ' AND ' ergänzen *)
  if Trim(QFaks.Filter) <> '' then
    Filter := Trim(QFaks.Filter) + ' AND ';

  (* den Filter zusammensetzen *)
  if QFaks.Fields[DBGridFaks.SelectedField.Index].IsNull then
    Filter := Filter + DBGridFaks.SelectedField.FieldName + ' is null'

  else
    Filter := Filter + DBGridFaks.SelectedField.FieldName + '=' +
      QuotedStr(QFaks.Fields[DBGridFaks.SelectedField.Index].AsString);




  end;

  //ShowMessage('Filter='+ NL + Filter);

  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  if FilterCombo.Items.IndexOf(Filter) = -1 then;
  FilterCombo.Items.Add(QFaks.Filter);

  FilterCombo.Text:=Filter;


  (* rotes Panel als Warnhinweis *)
  //FilterCombo.Visible := True;
  ShowFilterInfo(true);

end;

procedure TForm2.ApplicationProperties1Exception(Sender: TObject; E: Exception);
begin
  ShowMessage('Mist ein Fehler:' + NL + NL + E.Message + NL + NL + 'Faks Spalte RID hat den Wert ' + RowID);
  (* Oracle LogFile anzeigen *)
  //OpenLogFileClick(Sender);
end;

procedure TForm2.ApplicationProperties1Activate(Sender: TObject);
begin
  (* Application Version anzeigen *)
  //Application.Title:= Application.Title + ' ' + resourceVersionInfo;
  Application.MainForm.Caption:= Application.MainForm.Caption + ' [ ' + resourceVersionInfo + ' ]';

end;

procedure TForm2.Button2Click(Sender: TObject);
var line : String;
begin
  try
    jei;
    QFaks.Close;
    QFaks.SQL.Assign(Memo1.Lines);
    QFaks.Open;
    DateEditVon.Enabled := False;
    DateEditBis.Enabled := False;
    PageControl1.ActivePage := TabDaten;

    (* SQL in History speichern *)
    line := StringsToStr(Memo1.Lines,'°',true);
    if (SQLHistory.IndexOf(Line) = -1) then
        SQLHistory.Add(line);

    SQLHistoryIndex := SQLHistory.Count;


  finally
    Nei;

  end;

end;

procedure TForm2.CbincMonthClick(Sender: TObject);
begin
  if  CbincMonth.Checked then
  begin
    SpinEdit1Change(Sender);

    DateEditVon.Date := VonDatum;
    DateEditBis.Date := BisDatum;

    Verbinden();
    CbincMonth.Checked:= false;
  end;
end;

procedure TForm2.cbKeinGruenbergChange(Sender: TObject);
begin
  if cbKeinGruenberg.Checked then
  begin
    cbKeinGruenberg_kein_Wetterau.Checked := false;
    EditNoGo.Text := '70, 71, 72, 74, 77, 78, 79';
    ShowMessage('Bitte neu starten, damit die Änderungen wirksam werden!!');
  end;
end;

procedure TForm2.cbKeinGruenberg_kein_WetterauChange(Sender: TObject);
begin
  if cbKeinGruenberg_kein_Wetterau.Checked then
  begin
    cbKeinGruenberg.Checked := false;
    EditNoGo.Text := '70, 71, 72, 74, 77, 78, 79, 50, 51, 52, 53, 54, 55, 56, 57';
    ShowMessage('Bitte neu starten, damit die Änderungen wirksam werden!!' + NL +
    'Achtung: durch Linie 56 ist auch kein Kahlgrund dabei!!');
  end;
end;

procedure TForm2.CheckLinieClick(Sender: TObject);
var x, y : integer;
    lst : TStringList;
begin
  if not QFaks.Active then exit;
  (* Alle Liniennummern sammeln und anzeigen *)
  try
    Jei;
    lst := TStringList.Create;
    lst.Sorted:=true;
    lst.Duplicates:=dupIgnore;

    BM := QFaks.GetBookmark;

    QFaks.First;

    while not QFaks.EOF do
    begin
       lst.Add(QFaks.FieldByName('LINIE').AsString);
       QFaks.Next;
    end;

    (* nur unbekannte Linien zu ListBoxKnownLines hinzufügen *)
    for x := lst.Count -1 downto 0  do
    begin
      if ListBoxKnownLines.Items.IndexOf(lst[x]) > -1 then lst.Delete(x);
    end;



    Mem.Clear;
    Mem.Text := lst.Text;
    Mem.SelectAll;
    Mem.CopyToClipboard;
    QFaks.GotoBookmark(BM);

    if lst.Count > 0 then
    begin
    if Messagedlg('Sollen diese neuen Linien' + NL + lst.Text +  NL +
       'Auf der Seite ''Einstellungen'' der Liste bekannter Linien himnzugefügt werden?' ,
       mtConfirmation,[mbYes,mbNo],0)= mrYes then
    begin
       ListBoxKnownLines.Items.AddStrings(lst);
       PageControl1.ActivePage := TabConfig;
       ListBoxKnownLines.SetFocus;
    end;
    end
    else
      ShowMessage('Alle Linien sind bereits bekannt. Vorsicht Linie 901 und so''n Mist vor Import AmisData löschen!' + NL + lst.Text );




  finally
    Nei;
    QFaks.FreeBookmark(BM);
    FreeAndNil(lst);
  end;
end;

procedure TForm2.DateEditVonAcceptDate(Sender: TObject; var ADate: TDateTime;
  var AcceptDate: boolean);
begin
  try
    RemoveFilterClick(Sender);
    AcceptDate := True;
    StatusBar1.SimpleText:='Filter wurde entfernt, bitte den SQL-Code ggf. kontrollieren!!';
  finally

  end;
end;

procedure TForm2.DBase_exportClick(Sender: TObject);
var
  z, NichtGefunden, x, recs: integer;
  FName, StrLinie, Gattung: string;
  gefunden, Amis : boolean;
  ConvertErrors : TStringList;
  Einnahmen : currency;
  v : variant;
begin
  try
  try


    if Messagedlg('Sollen die Datensätze als nach AmisData exportiert gekennzeichnet werden?',mtConfirmation,[mbYes,mbNo],0)= mrYes then
     Amis := true
    else
     Amis := false;

    (* Tabelle nach Spalte RID aufsteigend sortieren *)
    QFaks.SortedFields:='RID';
    (* stAscending ist defibniert in ZAbstractRODataset *)
    QFaks.SortType:=stAscending;

    Einnahmen := 0;
    Screen.ActiveControl.Invalidate;
    Application.MainForm.Invalidate;
    Application.ProcessMessages;


    (* Fortschrittsbalken auf 0 stellen *)
    ProgressBar1.Position:=0;

    (* Liste für Konvertierungsfehler *)
    ConvertErrors := TStringList.Create;



    if not DirectoryExists(DirectoryEdit1.Directory) then
    begin
      PageControl1.ActivePage := TabConfig;
      DirectoryEdit1.SelectAll;
      ShowMessage('Das Exportverzeichnis: ' + NL + DirectoryEdit1.Directory +
        NL + 'existiert nicht, bitte einstellen!');
    end;

    (* Dateinamen zusammensetzen *)
    if ((DateEditBis.Date - DateEditVon.Date <= 31)
        and
        (MonthOf(DateEditVon.Date) = MonthOf(DateEditBis.Date)))
    then
    begin
      FName := 'FaksDaten_' + FormatDateTime('mmyy',DateEditVon.Date);
    end
    else
    begin
      FName := 'FaksDaten_' + FormatDateTime('dd.mm.yy',DateEditVon.Date) +'-' +
                              FormatDateTime('dd.mm.yy',DateEditBis.Date);
    end;

    Dbf1.TableName := IncludeTrailingBackslash(DirectoryEdit1.Directory) +
      FName + '.dbf';

    (* Dbf1 hat die FieldDefs der Elgeba *.dbf aus Tabelle gespeichert, plus ein neues Feld RID *)
    Dbf1.CreateTableEx(Dbf1.DBFFieldDefs);

    Jei;

    DBGridFaks.SetFocus;

    (* um das höchste Buchungsdatum und Buchungszeit im ersten Datensatz zu haben *)
    QFaks.SortedFields:='Buchungsdatum, BuchungsZeit';
    QFaks.SortType:=stDescending;
    QFaks.Refresh;

    QFaks.First;
    (* letztes Buchungsdatum und Buchungszeit in globaler Variable speichern *)
    //MaxBuchungsDatum := QFaks.FieldByName('Buchungsdatum').AsString;
    MaxBuchungsDatum := QFaks.FieldByName('Buchungsdatum').AsString + ' ' + QFaks.FieldByName('BuchungsZeit').AsString;

    QFaks.DisableControls;
    NichtGefunden := 0;

    ProgressBar1.Max := QFaks.RecordCount;

    recs := 0;


    Dbf1.Open;

    (* Daten von QFaks in Dbase Tabelle schreiben *)
    while not QFaks.EOF do
    begin
      try
      Application.ProcessMessages;
      (* diese RID auch in die globale Variable speichern, für Application.OnException Meldung *)
      RowID := QFaks.FieldByName('RID').AsString;

      Dbf1.Append;
      (* die Gerätenummer muss im AmisData definiert sein! *)
      //Das war die Personalnummer Dbf1.FieldByName('VDTNR').AsString :='999' + QFaks.FieldByName('PNR').AsString;
      (* Spalte MDEIDINTERN enthält die Gerätenummer *)
      Dbf1.FieldByName('VDTNR').AsString :='999' + QFaks.FieldByName('MDEIDINTERN').AsString;
      Dbf1.FieldByName('GNR').AsString :=
        QFaks.FieldByName('BELEGNR').AsString;
      Dbf1.FieldByName('ZEITV').AsString := QFaks.FieldByName('ZEIT').AsString;
      Dbf1.FieldByName('Datumv').AsDateTime :=
        QFaks.FieldByName('Datum').AsDateTime;


      Dbf1.FieldByName('GABZEIT').AsString :=
        QFaks.FieldByName('ZEIT').AsString;
      Dbf1.FieldByName('GABDatum').AsDateTime :=
        QFaks.FieldByName('DatumFahrt').AsDateTime;

      (* Nur für Butzbach exportieren wir VertriebsHSTIdent statt HSTSTARTIDENT *)
      if (QFaks.FieldByName('ID_F2MANDANT').AsInteger = 1) then
        Dbf1.FieldByName('HALTNR').AsString :=
         QFaks.FieldByName('VertriebsHSTIdent').AsString
      else
        Dbf1.FieldByName('HALTNR').AsString :=
         QFaks.FieldByName('HSTSTARTIDENT').AsString;

      Dbf1.FieldByName('SZONE').AsString :=
        QFaks.FieldByName('TZSTARTIDENT').AsString;
      Dbf1.FieldByName('ZZONE').AsString :=
        QFaks.FieldByName('TZZIELIDENT').AsString;
      Dbf1.FieldByName('VZONE').AsString :=
        QFaks.FieldByName('TZVIAIDENT').AsString;

      (* ab Dezember 2015 Sortennummer statt GAIDENT, ab TarifVersion 35  *)
      if QFaks.FieldByName('TarifVersion').AsInteger < TarifVersion then
        Dbf1.FieldByName('GATTUNG').AsString := QFaks.FieldByName('GAIDENT').AsString
      else
      begin
       (*
        if Copy(QFaks.FieldByName('Sortennummer').AsString,1,2) = '30' then GATTUNG :='3';
        if Copy(QFaks.FieldByName('Sortennummer').AsString,1,2) = '50' then GATTUNG :='5';
        if Copy(QFaks.FieldByName('Sortennummer').AsString,1,2) = '60' then GATTUNG :='6';
        if Copy(QFaks.FieldByName('Sortennummer').AsString,1,2) = '15' then GATTUNG :='17';

        Dbf1.FieldByName('GATTUNG').AsString := GATTUNG;
        *)
        (* ab Tarif 35 keine Gattung mehr, nur nch die Sortennummer *)
        Dbf1.FieldByName('GATTUNG').AsString := '-1';

      end;


      Dbf1.FieldByName('TarifNR').AsString :=
        QFaks.FieldByName('TarifVersion').AsString;

      (* Sonderfall Preisstufe 25 und 28 *)
       if QFaks.FieldByName('PV').AsString='RMV' then
       begin
         if QFaks.FieldByName('PREISSTIDENT').AsString = '205021' then
           Dbf1.FieldByName('PSTUFE').AsString :='25'
         else if QFaks.FieldByName('PREISSTIDENT').AsString = '205018' then
             Dbf1.FieldByName('PSTUFE').AsString :='28'
         else
           (* wird von AmisData nicht akzeptiert: PREISSTIDENT *)
           Dbf1.FieldByName('PSTUFE').AsString := QFaks.FieldByName('PREISSTdruck').AsString;
       end;


      Dbf1.FieldByName('PREIS').AsCurrency :=
        QFaks.FieldByName('Betrag').AsCurrency;

      (* für das Eventlog1 *)
      Einnahmen := Einnahmen + QFaks.FieldByName('Betrag').AsCurrency;

      (* wichtig für den nächsten Monat *)
      Dbf1.FieldByName('ZDEDATUM').AsDateTime :=
        QFaks.FieldByName('Buchungsdatum').AsDateTime;

      Dbf1.FieldByName('ZDEZEIT').AsString :=
        QFaks.FieldByName('BuchungsZeit').AsString;


      Dbf1.FieldByName('PREIS2').AsCurrency := 0;

      (* um den Ursprung des Datensatzes ggf. zu identifizieren *)
      Dbf1.FieldByName('RID').AsString :=
        QFaks.FieldByName('RID').AsString;

      (* idiotische Linie 31/32 korrigieren *)
      StrLinie := QFaks.FieldByName('Linie').AsString;

      if Pos('/', StrLinie) <> 0 then
      StrLinie := StringReplace(StrLinie,'/','',[rfReplaceAll]);

      if Pos('\', StrLinie) <> 0 then
      StrLinie := StringReplace(StrLinie,'\','',[rfReplaceAll]);

      if Pos('-', StrLinie) <> 0 then
      StrLinie := StringReplace(StrLinie,'-','',[rfReplaceAll]);

      Dbf1.FieldByName('LINIE').AsString := StrLinie;

      (* Nur für Mandant Butzbach: Linie 56 --> 1056 *)
      if ((QFaks.FieldByName('ID_F2MANDANT').AsInteger = 1) and
        (StrToInt(StrLinie) = 56)) then
      begin
        Dbf1.FieldByName('LINIE').AsInteger := 1056;
      end;

      (* Nur für Mandant Butzbach: Linie 52 --> 1052 *)
      if ((QFaks.FieldByName('ID_F2MANDANT').AsInteger = 1) and
        (StrToInt(StrLinie) = 52)) then
      begin
        Dbf1.FieldByName('LINIE').AsInteger := 1052;
      end;



      (* Sortennummer *)
      if ((QFaks.FieldByName('PV').AsString='RMV') And (QFaks.FieldByName('TarifVersion').AsInteger < TarifVersion  )) then
      begin
        //if QFaks.FieldByName('SORTENNUMMER').IsNull then
        Dbf1.FieldByName('SORTE').AsString :=
        FormatFloat('00', QFaks.FieldByName('GAIDENT').AsFloat) +
        FormatFloat('00', QFaks.FieldByName('PreisStDruck').AsFloat);
      end
      else
      begin
        Dbf1.FieldByName('SORTE').AsString := QFaks.FieldByName('SORTENNUMMER').AsString;

        (* Sonderfall Hessenticket *)
        if ((QFaks.FieldByName('SORTENNUMMER').IsNull ) AND
            (QFaks.FieldByName('PreisStDruck').AsString = '50')) then
            Dbf1.FieldByName('SORTE').AsString := '1750';
        (*
        Copy(QFaks.FieldByName('SORTENNUMMER').AsString,1,2) +
        FormatFloat('00', QFaks.FieldByName('PreisStDruck').AsFloat);
        *)

      end;


      //end;


      except
        on E:EConvertError do
        begin
         ShowMessage('Es ist ein Fehler aufgetreten bei ROWID. ' + RowId + NL + NL + E.Message + NL + NL + 'Fortsetzung erfolgt mit nächstem Datensatz!!' );
         QFaks.next;
         ConvertErrors.Add('ROWID=' + RowId);
         continue;
        end;
      end;





      Dbf1.FieldByName('STORNIERT').AsBoolean := False;

      if QFaks.FieldByName('PV').AsString='RMV' then
      Dbf1.FieldByName('WABE').AsString := '1';


      Dbf1.FieldByName('ZAHLART').AsString :=
        QFaks.FieldByName('ZAHLART').AsString;

      Dbf1.Post;

      (* Datensatzzähler erhöhen und in Prgressbar anzeigen *)
      inc(recs);
      ProgressBar1.Position:=recs;
      ProgressBar1.Invalidate;

      (* Datensatz als exportiert kennzeichen *)
      if Amis then
      begin
        v := Trim(QFaks.FieldByName('RID').AsString);
        ZUpDateRid.ParamByName('RID').AsString:=v;
        ZUpDateRid.ParamByName('Vertragsnr').AsInteger:=1;
        ZUpDateRid.ExecSQL;
      end;


      QFaks.Next;
    end;
    z := Dbf1.RecordCount;
    Dbf1.Close;

    FName := Dbf1.FilePathFull + Dbf1.TableName;

    if FileExists(FName) then
    begin

      (* Explorer *)
      if Messagedlg(IntToStr(z) + ' Datensätze wurden geschrieben in: ' +
        NL + NL + FName + NL + NL +
        'Soll die Datei im Dateimanager Explorer angezeigt werden?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        OpenExplorer(FName);
      end;

      (* Eventlog schreiben *)
      EventLog1.Log(FormatFloat('#,##0',z) + ' Datensätze und ' + FormatCurr('#,##0.00 EURO (Summe der Spalte PREIS)',Einnahmen) +  ' wurden geschrieben in ' + FName );
      EventLog1.Log('SQL-Code der Abfrage war: ' + StringsToStr(Memo1.Lines,' ',true));
    end;
  finally
    nei;
    QFaks.EnableControls;

    if ConvertErrors.Count > 0 then
    begin
      ConvertErrors.Insert(0,'Bei diesen Datensaetzen gabs Konvertierungsfehler, sie sind wahrscheinlich unvollstaendig in der DBase Datei: ' + FName);
      ConvertErrors.SaveToFile(ChangeFileExt(FName,'_Error.txt'));
      Shellexecute(Application.MainForm.Handle,'open',PChar(ChangeFileExt(FName,'_Error.txt')),'','',SW_NORMAL);

    end;

    FreeAndNil(ConvertErrors);

  end;

  except
    on E: Exception do
    begin
      mem.Visible := True;
      Mem.Clear;
      Mem.Text := RowID;
      Mem.SelectAll;
      Mem.CopyToClipboard;

       ShowMessage('Bei dem FAKS-Datensatz mit der  RID: ' + NL  + RowID + NL +
       ' ist ein Fehler aufgetreten. Die RID wurde kopiert!' + NL + 'Die Fehlermeldung lautet:' + NL + E.Message);

    end;
  end;

  QFaks.Refresh;

  (* Progressbar wieder auf 0 setzen *)
  ProgressBar1.Position:=0;
  ProgressBar1.Invalidate;
end;

procedure TForm2.DBGridFaksDblClick(Sender: TObject);
begin
    (* nur zu Fehlerdokumentation  *)
   ShowMessage('AsString: ' + QFaks.FieldByName('Betrag').AsString + NL +
               //'Value: ' +QFaks.FieldByName('Betrag').Value + NL +
               'AsFloat: ' + FormatFloat('#.000000',QFaks.FieldByName('Betrag').AsFloat));
end;

(* Anzeige von up und down Arrows laut:
http://wiki.freepascal.org/Grids_Reference_Page#Sorting_columns_or_rows_in_DBGrid_with_sort_arrows_in_column_header *)
procedure TForm2.DBGridFaksTitleClick(Column: TColumn);
const
  ImageArrowUp=0; //should match image in imagelist
  ImageArrowDown=1; //should match image in imagelist

begin
  try
    Jei;

    QFaks.DisableControls;
    BM := QFaks.GetBookmark;

    (* auf- bzw. absteigend sortieren *)
    DBGridFaks.SelectedField := QFaks.FieldByName(Column.Title.Caption);

    // Use the column tag to toggle ASC/DESC
    column.tag := not column.tag;


    if boolean(column.tag) then
    begin
      QFaks.SortedFields := Column.Title.Caption;
      (* stAscending ist defibniert in ZAbstractRODataset *)
      QFaks.SortType:=stAscending;
      Column.Title.ImageIndex:=ImageArrowUp;

    end
    else
    begin
      QFaks.SortedFields := Column.Title.Caption;
      QFaks.SortType:=stDescending;
      Column.Title.ImageIndex:=ImageArrowDown;
    end;

    QFaks.First;
    //QFaks.GotoBookmark(BM);

  finally
    QFaks.FreeBookmark(BM);

    QFaks.EnableControls;

    // Remove the sort arrow from the previous column we sorted
    if (FLastColumn <> nil) and (FlastColumn <> Column) then
      FLastColumn.Title.ImageIndex:=-1;

    FLastColumn:=column;

    Nei;
  end;
end;

procedure TForm2.DeleteSelectedClick(Sender: TObject);
var x : Integer;
begin
    x := ListBoxKnownLines.ItemIndex;
    if x > -1 then
    ListBoxKnownLines.DeleteSelected;
end;

procedure TForm2.DirectoryEdit1ButtonClick(Sender: TObject);
begin
  DirectoryEdit1.RootDir := ExePath;
end;

procedure TForm2.EventLog_anzeigenClick(Sender: TObject);
begin
  (* Eventlog anzeigen *)
  if FileExists(EventLog1.FileName) then
   ShellExecute(Application.MainForm.Handle,'open',PChar(EventLog1.FileName),'','',SW_Normal)
  else
   ShowMessage(EventLog1.FileName + NL + NL + 'wurde nicht gefunden!');
end;

procedure TForm2.ExportExcelClick(Sender: TObject);
begin
  try
    jei;
    ExportDatasetToExcel(QFaks);
  finally
    nei;

  end;
end;

procedure TForm2.FilterAbClick(Sender: TObject);
var Filter : string;
begin

  (* ggf. vorhandenen Filter durch ' AND ' ergänzen *)
  if Trim(QFaks.Filter) <> '' then
    Filter := Trim(QFaks.Filter) + ' AND ';

  (* den Filter zusammensetzen *)
  Filter := Filter + DBGridFaks.SelectedField.FieldName + '>=' +
    QuotedStr(QFaks.Fields[DBGridFaks.SelectedField.Index].AsString);


  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  if FilterCombo.Items.IndexOf(Filter) = -1 then;
  FilterCombo.Items.Add(QFaks.Filter);

  FilterCombo.Text:=Filter;

  ShowFilterInfo(true);

  (* zum ersten Datensatz springen *)
  QFaks.First;


end;

procedure TForm2.FilterBisClick(Sender: TObject);
var Filter : string;
begin

  (* ggf. vorhandenen Filter durch ' AND ' ergänzen *)
  if Trim(QFaks.Filter) <> '' then
    Filter := Trim(QFaks.Filter) + ' AND ';

  (* den Filter zusammensetzen *)
  Filter := Filter + DBGridFaks.SelectedField.FieldName + '<=' +
    QuotedStr(QFaks.Fields[DBGridFaks.SelectedField.Index].AsString);


  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  if FilterCombo.Items.IndexOf(Filter) = -1 then;
  FilterCombo.Items.Add(QFaks.Filter);

  FilterCombo.Text:=Filter;

  ShowFilterInfo(true);

  (* zum letzten Datensatz springen *)
  QFaks.Last;



end;

procedure TForm2.FilterComboKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    IF Key = VK_RETURN then
  begin
    if QFaks.Filtered then QFaks.Filtered := False;

    FilterCombo.Text := trim(FilterCombo.Text);

    QFaks.Filter:=FilterCombo.Text;

    if QFaks.Filter > '' then
    begin
      QFaks.Filtered := True;

      if ((FilterCombo.Items.IndexOf(FilterCombo.Text) = -1) and (FilterCombo.Text > '')) then;
      FilterCombo.Items.Add(FilterCombo.Text);
    end;

    ShowFilterInfo(true);

  end;

end;

procedure TForm2.FilterSetzenClick(Sender: TObject);
var
  Filter: string;
begin
  if QFaks.Filtered then QFaks.Filtered := False;


  (* den Filter zusammensetzen *)
  //Filter := QFaks.FieldByName('RID').FieldName + '=''' +
  //          ListBox_nicht_gefunden.Items[ListBox_nicht_gefunden.ItemIndex] + '''' ;


  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  if FilterCombo.Items.IndexOf(Filter) = -1 then;
  FilterCombo.Items.Add(QFaks.Filter);

  FilterCombo.Text:=Filter;

  if QFaks.RecordCount = 1 then
    PageControl1.ActivePage :=   TabDaten ;
  //else
  //   ShowMessage('Die RID ''' + ListBox_nicht_gefunden.Items[ListBox_nicht_gefunden.ItemIndex] +
  //     ''' steht nicht in den aktuellen Daten. Evtl aus einem Vormonat?');
end;

procedure TForm2.FormActivate(Sender: TObject);
var
  warning: boolean;
  msg : string;
begin
  (* das in FormCreate führt zu Access Violation!! *)
  Application.MainForm.Caption := ApplicationProperties1.Title;

  warning := False;

  if not DirectoryExists(DirectoryEdit1.Directory) then
  begin
    warning := True;
    msg := msg + ' Das Exportverzeichnis existiert nicht.';
  end;
  if trim(lbDatabase.Text) = '' then
  begin
    warning := True;
    msg := msg + ' Das Database-Name existiert nicht.';
  end;
  if trim(lbHostname.Text) = '' then
  begin
    warning := True;
    msg := msg + ' Der Hostname existiert nicht.';
  end;
  if trim(lbPassword.Text) = '' then
  begin
    warning := True;
    msg := msg + ' Das Passwort existiert nicht.';
  end;
  if trim(lbUserName.Text) = '' then
  begin
    warning := True;
    msg := msg + ' Der User-Name existiert nicht.';
  end;


  if warning then
  begin
    PageControl1.ActivePage := TabConfig;
    lbHostname.SetFocus;
    ShowMessage('Bitte erst die Einstellungen durchführen und dann neu starten!' + NL + NL + msg);
    exit;
  end;



  if not DirectoryExists(DirectoryEdit1.Directory) then
  begin
    PageControl1.ActivePage := TabConfig;
    DirectoryEdit1.SetFocus;
    DirectoryEdit1.SelectAll;
    ShowMessage('Das Export-Verzeichnis existiert nicht/wurde noch nicht eingegeben:' +
      NL + NL + 'Verzeichnis: ' + DirectoryEdit1.Directory);

  end;

  if Verbinden then
    jconnected := True
  else
    jconnected := False;

  (* ZSQLMonitor LogFile auf 3000 Zeilen kürzen *)
  ShortenLog(3000, Sender);
end;

procedure TForm2.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  QFaks.Close;
  //SQLTransaction2.Active := False;
  ZConnection1.Connected := False;

end;

procedure TForm2.FormCloseQuery(Sender: TObject; var CanClose: boolean);
var x : integer;
begin

  (* leere Einträge aus FilterCombo löschen *)
  for x := FilterCombo.Items.Count -1 downto 0 do
  begin
    if trim(FilterCombo.Items[x]) = '' then
      FilterCombo.Items.Delete(x);
  end;

  (* FilterCombo-Einträge ggf. löschen *)
  while FilterCombo.Items.Count >= 20 do
  FilterCombo.Items.Delete(0);


  (* Lokale ini mit ini in Monatsmeld vergleichen *)
  CheckIniFile(Sender);

  FreeAndNil(SQLHistory);

end;

procedure TForm2.JEi;
begin
  screen.cursor := crHourglass;
end;

procedure TForm2.NEi;
begin
  screen.cursor := crDefault;
end;

function TForm2.Verbinden: boolean;
var
  d, m, y, d1, m1, y1: word;
  (* Vormonat und VorVormonat!!!! *)
  DatumVon, DatumBis, DatumVonVormonat, DatumBisVormonat: TDateTime;
  sql, quotedNOGO, line, JBuchungsDatum, JBuchungsZeit : string;
  x : integer;
  F : TFloatField;
  dbf : TDbf;
  list : TStringList;
begin
  (* CAST(Betrag AS  NUMBER(8,2)) AS BETRAG *)

  //ShowMessage('Verbinden!!');
  {$IFDEF WINDOWS}
  (* Datenbank öffnen *)
  try
    ZConnection1.Connected := False;

    //SQLTransaction2.Active := False;
    QFaks.Close;

    if ((trim(lbHostName.Text) <> '') or (trim(lbDatabase.Text) <> '') or
      (trim(lbUserName.Text) <> '') or (trim(lbPassword.Text) <> '') ) then
    begin
      ZConnection1.HostName := lbHostName.Text;
      ZConnection1.Database := lbDatabase.Text;
      ZConnection1.User := lbUserName.Text;
      ZConnection1.Password := lbPassword.Text;
    end
    else
    begin
      PageControl1.ActivePage := TabConfig;
      lbHostName.SetFocus;
      lbHostName.SelectAll;
      exit;
    end;

    if trim(EditNoGo.Text) <> '' then
    begin
         if Messagedlg('Diese Linien werden NICHT(!!) berücksichtigt (ggf. Einstellungen bearbeiten)'
            + NL + NL + EditNoGo.Text, mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         exit;

    end;

    jei;
    (* Vormonat einstellen *)
    DatumVon := incMonth(date(), -1);
    DecodeDate(DatumVon, y, m, d);
    (* ersten des Vormonats *)
    d := 1;
    DatumVon := EncodeDate(y, m, d);

    (* DatumVon "01.12.2014" um einen Monat erhöhen "01.01.2015"
       und dann einen Tag abziehen, um den "31.12.2014" zu erhalten *)
    DatumBis := incMonth(DatumVon, 1) - 1;

    (* jetzt die gefundenen Datumswerte in die DateEdit eintragen *)
    if not FirstRun then
    begin
      DateEditVon.Date := DatumVon;
      DateEditBis.Date := DatumBis;
    end;

    (* In Januar NICHT mehr den Dezember mit altem Tarif dazunehmen *)
    if m <> 1 then
    begin
      (* ersten und letzten des Vorvor(!!)monats ermitteln *)
      DatumVonVormonat := incMonth(DatumVon, -1);
      DecodeDate(DatumVonVormonat, y1, m1, d1);
      (* ersten des VorVormonats *)
      d1 := 1;
      DatumVonVormonat := EncodeDate(y1, m1, d1);

      (* DatumVon "01.12.2014" um einen Monat erhöhen "01.01.2015"
         und dann einen Tag abziehen, um den "31.12.2014" zu erhalten *)
      DatumBisVormonat := incMonth(DatumVonVormonat, 1) - 1;
    end
    else
    begin
      (* für den Januar die bereits oben ermittelten Suchkriterien nehmen *)
      DatumVonVormonat := DatumVon;
      DatumBisVormonat := DatumBis;
    end;



    FirstRun := True;
    (*
    dbf :=  TDbf.Create(Application.MainForm);
    list := TStringList.Create;

    dbf.FilePath := 'F:\AMISdata\Save\BLE\';
    dbf.TableName:= 'Faksdaten_' + FormatDateTime('mmyy',DatumVonVormonat) + '.dbf';
    dbf.Open;
    dbf.First;

    StatusBar1.SimpleText:= 'Lese RID aus: ''' + dbf.TableName + '''';

    While not dbf.EOF do
    begin
      list.Add(QuotedStr(dbf.FieldByName('Rid').AsString));
      dbf.Next;
    end;


    dbf.Close;

    *)

    (* welche Linien sollen NICHT berücksichtigt werden? *)

    if trim(EditNoGo.Text) = '' then
    begin
      NoGo := ' AND ';
    end
    else
    begin
      (* da es so bescheuerte Liniennummern wie 31/32 gibt,
         die Liniennummer in Quotes setzen:

         Achtung: WordCount zählt 1-basiert!!!
         *)
      for x := 1 to WordCount(EditNoGo.Text,[','])  do
      begin
        if x = 1 then
          quotedNOGO:= QuotedStr(trim(ExtractWord(x,EditNoGo.Text,[','])))
        else
          quotedNOGO:= quotedNOGO + ',' + QuotedStr(Trim(ExtractWord(x,EditNoGo.Text,[','])))
      end;

      NoGo := ' AND LINIE NOT IN(' + quotedNOGO + ') AND ';
      //ShowMessage('Nogo = ' + Nogo);
    end;




    (* sehr tückisch ist die Klammersetzung für die Logik der SQL Verarbeitung!!!!!!! *)



    sql :=
      'SELECT RID, VID, ID_F2MANDANT, DATUM, TO_CHAR(ZEIT, ''HH24:MI:SS'') AS ZEIT, DATUMFAHRT, Buchungsdatum, '
       + NL +' TO_CHAR(BUCHUNGSZEIT, ''HH24:MI:SS'') AS BUCHUNGSZEIT ,  JOURNAL, MDEIDINTERN, BELEGNR, Bemerkung, PNR, '
      + NL + 'LINIE, FKART, anzahl, Einzelpreis, ' + NL
      +'  BETRAG, GAIDENT, GATTUNGSART, PreisStDruck, PREISSTIDENT, '
      + NL +
      'Zahlart, Storno, DatumZeit, TarifVersion, Netz, ORTStart, OrtZiel, PV, LfdNrPV, '
      + NL +
      'Storniert, Sortennummer, TZSTARTIDENT, TZZIELIDENT, TZVIAIDENT, HSTSTARTIDENT, HSTZIELIDENT, VERTRIEBSHSTIDENT, Vertragsnr '
      + NL + 'FROM F2FSV  WHERE ID_F2MANDANT IN(0,1) AND PV<>''HLB'' AND (PV=''RMV'') AND TarifVersion >=' + IntToStr(TarifVersion) + '  '
      + NOGO + ' ((DATUM BETWEEN ' + NL + '''' +
      DateTimeToStr(DateEditVon.Date) + ''' AND ' +
      NL + '''' + DateTimeToStr(DateEditBis.Date) + ''')' + NL +
      (* jetzt mit Vertragsnr is null *)
    //' OR (' + JBUCHUNGSDATUM + ' AND DATUM < ' + NL + '''' +
    //DateTimeToStr(DateEditBis.Date) + '''))';

    ' OR ( Vertragsnr is null AND DATUM <= ' + NL + '''' +
    DateTimeToStr(DateEditBis.Date) + ''' AND DATUM >=''01.01.' + IntToStr(Yearof(DateEditVon.Date)) +  '''))';







     QFaks.SQL.Text := sql;

    //ShowMessage(QFaks.SQL.Text);

    //exit;

    //ShowMessage(sql);

    //exit;

    ZConnection1.Connected := True;
    QFaks.Active := True;
    PageControl1.ActivePage := TabDaten;



    (* SQL in History speichern *)
    line := StringsToStr(Memo1.Lines,'°',true);
    //if (SQLHistory.IndexOf(Line) = -1) then
        SQLHistory.Add(line);

    SQLHistoryIndex := SQLHistory.Count;




  finally
    nei;
    Result := True;
    //FreeAndNil(dbf);
    //FreeAndNil(List);
  end;
  {$ENDIF}
  ;

end;

procedure TForm2.BtnExportExcelClick(Sender: TObject);
begin
  DBGridFaks.SetFocus;
  if not (activeControl is TCustomDBGrid) then
  begin
    ShowMessage('Klicken Sie bitte erst in die zu exportierende Tabelle!');
    exit;
  end
  else
  begin
    BM := QFaks.GetBookmark;
    ExportDatasetToExcel(QFaks);
    QFaks.GotoBookmark(BM);
    QFaks.FreeBookmark(BM);
  end;

end;

procedure TForm2.SumClick(Sender: TObject);
var
  Preis, PreisDelta: currency;
  row : integer;
  VorVorMonat : TDateTime;
begin
  try

    if not QFaks.Active then
      exit;

    Screen.ActiveControl.Invalidate;


    (* Tabelle nach Spalte RID aufsteigend sortieren *)
    QFaks.SortedFields:='RID';
    (* stAscending ist defibniert in ZAbstractRODataset *)
    QFaks.SortType:=stAscending;

    BM := QFaks.GetBookmark;
    QFaks.DisableControls;
    QFaks.First;
    Preis := 0;
    PreisDelta := 0;
    row := 0;
    Jei;

    if DBGridFaks.SelectedRows.Count > 1 then
    begin
      for row := 0 to DBGridFaks.SelectedRows.Count -1 do
      begin
        QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));
        Preis := Preis + QFaks.FieldByName('Betrag').AsCurrency;
      end
    end
    else
    begin
      while not QFaks.EOF do
      begin
        Preis := Preis + QFaks.FieldByName('Betrag').AsCurrency;
        QFaks.Next;
      end;
    end;
    nei;

    if row > 1 then
    (* nur für gewählte Zellen *)
    begin
      ShowMessage('Anzahl der Datensätze: ' + IntToStr(row +1) +
        NL + NL + 'Summe Spalte ''Betrag'': ' + FormatFloat('#,##0.00 EURO', Preis));
    end
    else
    begin
    (*  für die ganze Tabelle *)
      ShowMessage('Anzahl der Datensätze: ' + IntToStr(QFaks.RecordCount) +
        NL + NL + 'Summe Spalte ''Betrag'': ' + FormatFloat('#,##0.00 EURO', Preis));

      (* welcher Monat liegt zwei Monate zurück? *)
      VorVorMonat := incMonth(Date, -2);

      (* wird der richtige Monat betrachtet: ist er im Zeitraum oder kleiner?
         und liegen Beginn und Ende im selben Monat? *)
      If  (((VorVorMonat >= DateEditVon.Date) and (VorVorMonat <= DateEditBis.Date) )
          Or
          (VorVorMonat > DateEditVon.Date) )
          (* Beginn und Ende im selben Monat *)
          (* and (FormatDateTime('mmyyyy',DateEditVon.Date) = FormatDateTime('mmyyyy',DateEditBis.Date)) *)  then
      begin
        (* Summe der Einnahmen laut Preis in globaler Variablen speichern *)
        GesamtEinnahme:=Preis;

        (* Kontroll-Vergleich mit bereits in Amisdata importierten Daten *)
        if Messagedlg('Soll der gerade ermittelte Betrag: ' + NL + FormatFloat('#,##0.00 EURO', Preis) + NL +
        'mit bereits in Amsidata für den Zeitraum ' +  FormatDateTime('dd.mm.yyyy', DateEditVon.Date)  +
        ' bis ' + FormatDateTime('dd.mm.yyyy', DateEditBis.Date) +
        ' gemeldeten Monatsmeldungen verglichen werden?' + NL + NL +
        'Die SEHR langwierige Aktion kann duch ESC-Taste abgebrochen werden!',mtConfirmation,[mbYes,mbNo],0)= mrYes then
        BereitsGemeldetWurden(Sender);

      end;

    end;


  finally
    if  QFaks.BookmarkValid(BM) then QFaks.GotoBookmark(BM);
    QFaks.EnableControls;

  end;
end;

procedure TForm2.QFaksBeforeOpen(DataSet: TDataSet);
begin
  if not ZSQLMonitor1.Active then
    ZSQLMonitor1.Active := True;
end;

                (* wichtig: Damit das Ü in Grünberg erhalten bleibt: UTF8ToAnsi(Zeile) *)
function TForm2.ExportDatasetToExcel(JDataset: TDataset): boolean;
var
  x, row, recs: integer;
  F: TextFile;
  Zeile: string;
  FName_csv, FName_xls, command: variant;
  V, VApp: olevariant;
  ExcelHandle : HWND;


  function ExcludeParam(FName: string): string;
  var
    x: integer;
  begin
    {* ShellFindExecutable liefert z.B. für Excel '\Pfad\Excel.exe /e'
       /e wird aus FName entfernt *}
    Result := FName;
    x := Pos('/', FName);
    if x > 0 then
      Result := Trim(Copy(FName, 1, x - 1));
  end;

begin

  if not DirectoryExists(DirectoryEdit1.Directory) then
  begin
    PageControl1.ActivePage := TabConfig;
    DirectoryEdit1.SetFocus;
    ShowMessage('Bitte erst das Exportverzeichnis einstellen/korrigieren!');
    exit;
  end;

  Result := False;

  ProgressBar1.Position:=0;
  ProgressBar1.Max:=JDataset.RecordCount;
  recs := 0 ;

  with JDataset do
  begin
    try
      //if recordcount > 15000 then  showmessage('ACHTUNG: Je nach Excel-Version werden ggf. nicht alle Daten angezeigt, da die Zeilenzahl einer Tabelle beschränkt ist. MS-Access ist eine Alternative.');
      screen.cursor := crHourglass;
      JDataset.DisableControls;
      AssignFile(F, ExePath + Name + '.csv');
      Rewrite(F);

      First;
      Zeile := '';
      {FeldNamen}
      for x := 0 to Fieldcount - 1 do
      begin
        if x = 0 then
          Zeile := '"' + Fields[x].FieldName + '"'
        else
          Zeile := Zeile + ';' + '"' + Fields[x].FieldName + '"';
      end;
      Writeln(F, UTF8ToAnsi(Zeile));

      (* Mit und ohne Multiselect im Grid: *)
      if DBGridFaks.SelectedRows.Count > 1 then
      begin
         (* also mit Multiselect *)
         ProgressBar1.Position:=0;
         ProgressBar1.Max:=DBGridFaks.SelectedRows.Count;
         recs := 0 ;

         Zeile := '';
         for row := 0 to DBGridFaks.SelectedRows.Count -1 do
         begin
           GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));;

           for x := 0 to Fieldcount - 1 do
           begin
             if x = 0 then
             begin
               if Fields[x].isnull then
                 Zeile := 'leer'
               else
                 Zeile := Fields[x].AsString;
             end
             else
             begin
               if Fields[x].isnull then
                 Zeile := Zeile + ';' + 'leer'
               else
                 Zeile := Zeile + ';' + Fields[x].AsString;
             end;
           end;
           Writeln(F, UTF8ToAnsi(Zeile));
           inc(recs);
           ProgressBar1.Position:=recs;
           //Next;
           Zeile := '';
         end;
         Flush(F);
         CloseFile(F);



      end
      else
      begin
      (* also ohne Multiselect, ganzes Grid *)
      Zeile := '';
        while not EOF do
        begin
          for x := 0 to Fieldcount - 1 do
          begin
            if x = 0 then
            begin
              if Fields[x].isnull then
                Zeile := 'leer'
              else
                Zeile := Fields[x].AsString;
            end
            else
            begin
              if Fields[x].isnull then
                Zeile := Zeile + ';' + 'leer'
              else
                Zeile := Zeile + ';' + Fields[x].AsString;
            end;
          end;
          Writeln(F, UTF8ToAnsi(Zeile));
          inc(recs);
          ProgressBar1.Position:=recs;
          Next;
          Zeile := '';
        end;
        Flush(F);
        CloseFile(F);

    end;

      (* csv Datei in Excel öffnen und als *.xls abspeichern
         Achtung: FName_csv, FName_xls : Variant (!!!) nicht string
      *)

      FName_csv := ExePath + Name + '.csv';

      if DirectoryExists(DirectoryEdit1.Directory) then
        FName_xls := includeTrailingBackslash(DirectoryEdit1.Directory) +
          'Daten_' + DateEditVon.Text + '-' + DateEditBis.Text + '.xls'
      else
      begin
        FName_xls := ExePath + 'Daten_' + DateEditVon.Text + '-' +
          DateEditBis.Text + '.xls';
        ShowMessage('Achtung, die Daten stehen jetzt in' + NL + NL + FName_xls);
      end;

      if recs > 65500 then  ShowMessage('Achtung, die Daten haben ' + IntToStr(recs) + ' Zeilen, evtl. wird in Excel nicht alles angezeigt.' + NL + NL + 'Ggf. Rechtsklick, kopieren!!');

      (* Excel neu starten oder laufendes verwenden *)
      try
        V := GetActiveOleObject('Excel.Application')
      except
        V := CreateOleObject('Excel.Application');
      end;
      //FName := 'FileName := ' + FName_csv + ', local=true';

      VApp := V.Application;

      (* csv Datei mit Excel öffnen *)
      V.Workbooks.OpenText(FName_csv);

      //OpenText(FileName, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local)
      //V.Workbooks.OpenText(FName_csv, , , , , , , , , , , , , , , , , true);
      (*
      Set shFirstQtr = Workbooks(1).Worksheets(1)
      Set qtQtrResults = shFirstQtr.QueryTables.Add( _
    Connection := "TEXT;C:\My Documents\19980331.txt",
    Destination := shFirstQtr.Cells(1,1))


      *)


      (* Excel anzeigen *)
      V.Visible := True;

      (* Excel in den Vordergrund bringen, endlich in DelphiPraxis die Lösung gefunden *)
      ExcelHandle := V.Hwnd; {* Handle wird ermittelt *}
      v.WindowState := SW_SHOWNORMAL; {* Fenster wird aus der Taskleiste geholt *}
      SetForegroundWindow(ExcelHandle); {* Fenster wird in den Vordergrund geholt *}

      V.ActiveWorkbook.Activate;

      (* Nachfolgendes war sehr mühsam:
         CurrentRegion auswählen und mit Namen benennen
      *)
      V.Cells[1, 1].Select;
      v.Selection.CurrentRegion.Select;
      V.Selection.Name := 'JDaten';
      V.Selection.AutoFilter;
      command := DateEditVon.Text + '-' + DateEditBis.Text;
      V.ActiveSheet.Name := command;
      V.Range['A2'].Select;
      V.ActiveWindow.FreezePanes := True;

      (* Textfeld mit Hinweis für Linie 56 zu 1056 *)
      V.ActiveSheet.Shapes.AddTextbox(1, 303.75, 38.25, 285, 69).Select;
      if QFaks.Filter = '' then
      begin
        command := 'Achtung: Bei Wetterau Linie 56 in 1056 umbenennen!!!';
        V.Selection.Characters.Text := command;
        V.Selection.Font.Bold := True;
      end
      else
      begin
         command := 'Achtung: Bei Wetterau Linie 56 in 1056 umbenennen!!!' + NL + 'Datenfilter ist: ' + QFaks.Filter;
         V.Selection.Characters.Text := command;
      end;

      V.Range['A2'].Select;


      (* jetzt als richtige Excel Datei speichern *)
      V.ActiveSheet.SaveAs(FName_xls, 1);
      (* die csv-Quelldatei kann weg *)
      if recs < 65500 then
        SysUtils.DeleteFile(ExePath + Name + '.csv')
      else
         ShowMessage('Siehe auch die Daten in' + NL + NL + ExePath + Name + '.csv' + NL + 'mit ' + IntToStr(recs) +' Zeilen');

      Result := True;
      (*
          Range("C10").Select
          ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:= "'Daten_31.12.2014-31.12.2014.xls'!JDaten"


      *)
      (* Das will nicht!!
      command := 'F:\AMISdata\Monat\Monatsmld\Mon_2015\FAKS_Auswertungen\Daten_Auswertung.xls';
      V.Workbooks.Open(command);
      V.Range['C10'].Select;
      //command := ' xlDatabase, ''F:\AMISdata\Monat\Monatsmld\Mon_2015\FAKS_Auswertungen\Daten_31.12.2014-31.12.2014.xls''!JDaten';
      command := ' xlDatabase, ''F:\AMISdata\Monat\Monatsmld\Mon_2015\FAKS_Auswertungen\Daten_31.12.2014-31.12.2014.xls''!JDaten';
      V.ActiveSheet.PivotTableWizard( command);
      *)


      Application.Minimize;

    finally
      screen.cursor := crDefault;
      JDataset.EnableControls;
      V := Unassigned;
      ProgressBar1.Position:=0;


    end;
  end;
end;

function TForm2.OpenExplorer(FName: string): boolean;
begin
  try
    (* mit Windows Explorer öffnen *)
      {$IFDEF MSWINDOWS}
    AsyncProcess1.CommandLine := 'Explorer.exe /n,/e,/select,"' + FName + '"';
    AsyncProcess1.Active := True;
     {$ENDIF}
    Result := True;

  finally
  end;

end;

procedure TForm2.QFaks2DbaseClick(Sender: TObject);
var x : integer;
    Dbase : TDBF;
begin
  try
    Jei;
    Dbase := TDBF.Create(Application.Mainform);
    Dbase.TableLevel:=4;
    Dbase.FieldDefs.Assign(QFaks.FieldDefs);
    Dbase.TableName:=Exepath + 'FaksExport.dbf';
    Dbase.CreateTable;
    Dbase.Open;
    BM := QFaks.GetBookmark;

    QFaks.DisableControls;

    QFaks.First;


    while not QFaks.EOF do
    begin

      Dbase.Append;

      for x := 0 to QFaks.FieldCount -1 do
      begin
        //if ((QFaks.Fields[x].FieldName <> 'BETRAG') and (QFaks.Fields[x].FieldName <> 'EINZELPREIS')) then
        //ShowMessage(QFaks.Fields[x].FieldName + ' ' + QFaks.Fields[x].AsString);
        Dbase.Fields[x].AsVariant:= QFaks.Fields[x].AsVariant;

      end;

      Dbase.Post;

      QFaks.Next;

    end;

    ShowMessage('Die Daten stehen jetzt in: ' + NL + NL +
    EXEPATH + Dbase.TableName + NL + NL +'... der Explorer wird geöffnet!' + NL + NL +
    'Achtung: die Feldnamen haben nur 10 Buchstaben und es gibt keine Umlauts!!');

    OpenExplorer(EXEPATH + Dbase.TableName);


    //Paradox1.AppendRecord(VarArrayOf([QFaks.Fields[0]]));
    //Paradox1.AppendRecord(VarArrayOf(['1']));

    Dbase.Close;



  finally
    Nei;
    FreeAndNil(Dbase);


    if QFaks.BookmarkValid(BM) then
    QFaks.GotoBookmark(BM);

    QFaks.FreeBookmark(BM);

    QFaks.EnableControls;


  end;

end;



initialization

end.
