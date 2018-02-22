unit Unit1;


(*

C:\oracle\product\10.2.0\client_1\oci.dll

  Datenbanknamen auf hlbst02
  VASP = produktiv
  VAST = Test
  VASTS = Tarifserver

  user =  sysadm
  password = sysadm

  ID_F2MANDANT = 1 ( Butzbach )
  ID_F2MANDANT = 0 ( Giessen )

  Eintrag in tnsnames.ora, die in Exepath stehen muss:

  VASP =
  (DESCRIPTION =
    (ADDRESS_LIST =
      (ADDRESS = (PROTOCOL = TCP)(HOST = hlbst02.hlb.local)(PORT = 1521))
    )
    (CONNECT_DATA =
      (SERVICE_NAME = VASP)
    )
  )

Felder in FAKS F2FSV ( Stand: 07/2016  !! )

RID
VID
ID_F2MANDANT
DATUM
ZEIT
DATUMFAHRT
Buchungsdatum
BUCHUNGSZEIT
JOURNAL
MDEIDINTERN
BELEGNR
Bemerkung
PNR
LINIE
FKART
anzahl
Einzelpreis
Betrag
GAIDENT
GATTUNGSART
PreisStDruck
PREISSTIDENT
Zahlart
Storno
DatumZeit
TarifVersion
Netz
ORTStart
OrtZiel
PV
LfdNrPV
Storniert
Sortennummer
TZSTARTIDENT
TZZIELIDENT
TZVIAIDENT
HSTSTARTIDENT
HSTZIELIDENT
VERTRIEBSHSTIDENT
VERTRAGSNR


Wohin damit?
alter session set nls_numeric_characters =',.'

hier wird das Formatproblem scheints auch schon besprochen
http://forum.lazarus.freepascal.org/index.php/topic,20305.msg117063.html#msg117063

*)

{$mode objfpc}{$H+}

interface

uses
  locale_de, my_utils, LCLIntf, UniqueInstance, Classes, SysUtils, DB, dbf,
  FileUtil, SynHighlighterSQL, SynEdit, RTTIGrids, LResources, Forms, Controls,
  Graphics, Dialogs, ComCtrls, Menus, IniPropStorage, ExtCtrls, DBCtrls,
  DBGrids, StdCtrls, EditBtn, AsyncProcess, ZConnection, ZDataset, ZSqlMonitor,
  ZSqlMetadata, Interfaces, dateutils, comobj, variants, LCLType,
  ZAbstractRODataset, JwaWindows, ShellApi, StrUtils, PropertyStorage, Spin,
  Grids, ExtDlgs, FileCtrl, fpsexport, INIFiles, eventlog, resource,
  versiontypes, versionresource, fpDBExport, sqldb;

type
  //Letters = array ['A'..'Z']  of String;


  { TForm1 }

  TForm1 = class(TForm)
    ApplicationProperties1: TApplicationProperties;
    AsyncProcess1: TAsyncProcess;
    BtnAddLine1: TButton;
    BtnConnect: TButton;
    BtnAddLine: TButton;
    BtnDelLine1: TButton;
    BtnSearch: TButton;
    BtnSearch1: TButton;
    Button2: TButton;
    BtnDelLine: TButton;
    CbincMonth: TToggleBox;
    cbSQL: TCheckBox;
    DataSource2: TDataSource;
    DateEditBis: TDateEdit;
    DateEditVon: TDateEdit;
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
    DBGridUmsetzung: TDBGrid;
    DBGridFaks: TDBGrid;
    DBNavigator2: TDBNavigator;
    DirectoryEdit1: TDirectoryEdit;
    EventLog1: TEventLog;
    FilterCombo: TComboBox;
    GridReplaceGattungsart: TStringGrid;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    ImageList1: TImageList;
    IniPropStorage1: TIniPropStorage;
    Label1: TLabel;
    Label2: TLabel;
    DBase_export: TMenuItem;
    kopieren: TMenuItem;
    AuswahlFilter: TMenuItem;
    FilterAb: TMenuItem;
    FilterBis: TMenuItem;
    Label3: TLabel;
    Label4: TLabel;
    Edit_Max_HSTIDENT: TLabeledEdit;
    Label5: TLabel;
    lbDatabase: TLabeledEdit;
    lbPassword: TLabeledEdit;
    lbRecordCount: TLabel;
    lbUserName: TLabeledEdit;
    ListBoxKnownLines: TListBox;
    Mem: TEdit;
    MenuItem1: TMenuItem;
    CheckLinie: TMenuItem;
    DeleteSelected: TMenuItem;
    MenuAddLinie: TMenuItem;
    LookupStornos: TMenuItem;
    MnCheckTarifPreise: TMenuItem;
    MnJumpToNewValue: TMenuItem;
    MnSortBy: TMenuItem;
    MnOraFixError: TMenuItem;
    mnDelFilter: TMenuItem;
    mnLoadFilter: TMenuItem;
    mnSaveFilter: TMenuItem;
    MnSetAST_TarifVersion: TMenuItem;
    MnShowRecordCount: TMenuItem;
    MnManuelleBuchungen: TMenuItem;
    MNCheckLinieJeMandant: TMenuItem;
    MnGroupValues: TMenuItem;
    MnMore: TMenuItem;
    MnFixColumn: TMenuItem;
    MnAnrufsammeltaxis: TMenuItem;
    MnFieldlist: TMenuItem;
    MnSucheLinien: TMenuItem;
    MnSearch: TMenuItem;
    Panel9: TPanel;
    PopupGridReplace: TPopupMenu;
    PopupFilterCombo: TPopupMenu;
    Shape1: TShape;
    SpinEditTarifversion: TSpinEdit;
    UniqueInstance1: TUniqueInstance;
    UpdateElgebaLinienNr: TMenuItem;
    MnSumColumn: TMenuItem;
    MnMarkExported: TMenuItem;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    PopupLinien: TPopupMenu;
    RID_as_Filter: TMenuItem;
    Pm_Search: TMenuItem;
    ProgressBar1: TProgressBar;
    RemoveFilter: TMenuItem;
    OpenLogFile: TMenuItem;
    SpinEdit1: TSpinEdit;
    Splitter2: TSplitter;
    GridReplaceLineNumber: TStringGrid;
    Sum: TMenuItem;
    Splitter1: TSplitter;
    SQLLoad: TMenuItem;
    EventLog_anzeigen: TMenuItem;
    ExportExcel: TMenuItem;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    SaveToFile: TMenuItem;
    Panel1: TPanel;
    Memo1: TSynEdit;
    PopupSQL: TPopupMenu;
    SynSQLSyn1: TSynSQLSyn;
    TabConfig: TTabSheet;
    PageControl1: TPageControl;
    Panel2: TPanel;
    PopupGrid: TPopupMenu;
    StatusBar1: TStatusBar;
    TabDaten: TTabSheet;
    UpDown1: TUpDown;
    ZConnection1: TZConnection;
    QFaks: TZReadOnlyQuery;
    ZCheckFields: TZReadOnlyQuery;
    ZConnection2: TZConnection;
    QPreisliste: TZReadOnlyQuery;
    ZUpDateRid: TZQuery;
    ZSQLMetadata1: TZSQLMetadata;
    ZSQLMonitor1: TZSQLMonitor;
    procedure ApplicationProperties1Exception(Sender: TObject; E: Exception);
    procedure ApplicationProperties1Hint(Sender: TObject);
    procedure AuswahlFilterClick(Sender: TObject);
    procedure BtnAddLine1Click(Sender: TObject);
    procedure BtnAddLineClick(Sender: TObject);
    procedure BtnDelLine1Click(Sender: TObject);
    procedure BtnDelLineClick(Sender: TObject);
    procedure BtnSearchClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CbincMonthClick(Sender: TObject);
    procedure cbSQLChange(Sender: TObject);
    procedure CheckLinieClick(Sender: TObject);
    procedure DateEditVonAcceptDate(Sender: TObject; var ADate: TDateTime;
      var AcceptDate: boolean);
    procedure DBase_exportClick(Sender: TObject);
    procedure DBGridFaksDblClick(Sender: TObject);
    procedure DBGridFaksPrepareCanvas(sender: TObject; DataCol: Integer;
      Column: TColumn; AState: TGridDrawState);
    procedure DBGridFaksTitleClick(Column: TColumn);
    procedure DeleteSelectedClick(Sender: TObject);
    procedure DirectoryEdit1ButtonClick(Sender: TObject);
    procedure EventLog_anzeigenClick(Sender: TObject);
    procedure ExportExcelClick(Sender: TObject);
    procedure FilterAbClick(Sender: TObject);
    procedure FilterBisClick(Sender: TObject);
    procedure FilterComboKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure FilterSetzenClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure IniPropStorage1RestoreProperties(Sender: TObject);
    procedure IniPropStorage1StoredValues1Restore(Sender: TStoredValue;
      var Value: TStoredType);
    procedure IniPropStorage1StoredValues1Save(Sender: TStoredValue;
      var Value: TStoredType);
    procedure IniPropStorage1StoredValues2Restore(Sender: TStoredValue;
      var Value: TStoredType);
    procedure kopierenClick(Sender: TObject);
    procedure ListBoxKnownLinesKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure ListFieldsClick(Sender: TObject);
    procedure LookupStornosClick(Sender: TObject);
    procedure MenuAddLinieClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MnAnrufsammeltaxisClick(Sender: TObject);
    procedure MNCheckLinieJeMandantClick(Sender: TObject);
    procedure MnCheckTarifPreiseClick(Sender: TObject);
    procedure mnDelFilterClick(Sender: TObject);
    procedure MnFieldlistClick(Sender: TObject);
    procedure MnFixColumnClick(Sender: TObject);
    procedure MnGroupValuesClick(Sender: TObject);
    procedure MnJumpToNewValueClick(Sender: TObject);
    procedure mnLoadFilterClick(Sender: TObject);
    procedure MnManuelleBuchungenClick(Sender: TObject);
    procedure MnMarkExportedClick(Sender: TObject);
    procedure MnOraFixErrorClick(Sender: TObject);
    procedure mnSaveFilterClick(Sender: TObject);
    procedure MnSearchClick(Sender: TObject);
    procedure MnSetAST_TarifVersionClick(Sender: TObject);
    procedure MnShowRecordCountClick(Sender: TObject);
    procedure MnSortByClick(Sender: TObject);
    procedure MnSucheLinienClick(Sender: TObject);
    procedure MnSumColumnClick(Sender: TObject);
    procedure OpenLogFileClick(Sender: TObject);
    procedure Pm_SearchClick(Sender: TObject);
    procedure PopupGridClose(Sender: TObject);
    procedure PopupGridPopup(Sender: TObject);
    procedure RemoveFilterClick(Sender: TObject);
    procedure RID_as_FilterClick(Sender: TObject);
    procedure SaveToFileClick(Sender: TObject);
    procedure SpinEdit1Change(Sender: TObject);
    procedure SpinEditTarifversionEditingDone(Sender: TObject);
    procedure SQLLoadClick(Sender: TObject);
    procedure UniqueInstance1OtherInstance(Sender: TObject;
      ParamCount: integer; Parameters: array of string);
    procedure QFaksAfterOpen(DataSet: TDataSet);
    procedure BtnExportExcelClick(Sender: TObject);
    procedure SumClick(Sender: TObject);
    procedure QFaksBeforeOpen(DataSet: TDataSet);
    procedure QFaks2DbaseClick(Sender: TObject);
    procedure UpdateElgebaLinienNrClick(Sender: TObject);
    procedure UpDown1Click(Sender: TObject; Button: TUDBtnType);
  private
    { private declarations }
  public
    { public declarations }
    procedure JEi;
    procedure NEi;
    (* hier wird er SQL-Code erzeugt *)
    function Verbinden(): boolean;
    (* wichtig: Damit das ü in Grünberg erhalten bleibt: PWideChar(UTF8Decode(Zeile)) *)
    function ExportDatasetToExcel(JDataset: TDataset): boolean;
    function OpenExplorer(FName: string): boolean;
    procedure angezeigteDatenkopieren1Click(Sender: TObject);
    function OpenLog(FName: string): boolean;
    procedure ShortenLog(anzLines: integer; Sender: TObject);
    function IsDate(str: string): boolean;
    function ShowFilterInfo(Warning: boolean): boolean;
    procedure BereitsGemeldetWurden(Sender: TObject);
    procedure CheckIniFile(Sender: TObject);
    function ExistInDb(TableName: string): boolean;
    function ExistField(FieldName: string): boolean;
    function MarkExported(DBFName: string; mark: string = '1';
      IDs: TStringList = nil): boolean;
    function resourceVersionInfo: string;
    procedure ShowVersionInfo;
    function JNochZuMelden(Sender: TObject): double;
    (* wurde ersetzt durch TStringList "RIDS_in_Selection", das ist schneller *)
    function IsRIDinSelection(RID: string): boolean;

    (* Einnahmen zu Fehlermeldungen anzeigen *)
    function ShowFormSum(var Titel, Betrag: TStringList): boolean;
    (* Spalten nur mit Zeit kann ich nicht filtern *)
    function CheckFilterPossible(FieldName : String):Boolean;

    function LookUpStringGrid(Grid: TStringGrid;GesuchterMandant, ZuErsetzendeLinienNr: string):string;

    (* sucht die LinienNummern je Mandant *)
    procedure LinieJeMandant ( Sender: TObject);

    (* Filter zur History-Liste der FilterCombo hinzufügen *)
    function AddFilterHistory(item : String):boolean;

    procedure ShowRecordCount(Sender: TObject);

    (* Versuch den Fehler ungültige Zahl ORA-01722 zu fixen *)
    function AnalyzeTable(Table: string): boolean;




  end;

var
  Form1: TForm1;

  ExePath: string;
  jconnected: boolean = False;
  StoppIt: boolean = False;
  FirstRun: boolean = False;
  ShowAgain: boolean = true;
  BM: TBookmark;
  abort: boolean = False;
  gesucht: variant;
  RowID: string; (* soll den Datensatz bei Fehlermeldungen kennzeichnen *)
  FLastColumn: TColumn; //store last grid column we sorted on

  SavePath
  (* SavePath steht in der *.ini, wohin es laut Dialog eingetragen wird *): string;
  GesamtEinnahme, NochNichtGemeldet: currency;
  VonDatum, BisDatum: TDateTime;
  SQLHistory, RIDS_in_Selection: TStringList;
  SQLHistoryIndex: integer;
  (* wird in FormCreate berechnet *)
  TarifVersion: integer = 35;

  FTableDatabase: string = 'f2fsv';

const
  NL = chr(10) + chr(13);
  YearPrefix = '02.01.1899 ';




implementation

{ TForm1 }

{$R *.lfm}

uses summen, Unit2;

function TForm1.resourceVersionInfo: string;

  (* Unlike most of AboutText (below), this takes significant activity at run-    *)
  (* time to extract version/release/build numbers from resource information      *)
  (* appended to the binary.                                                      *)

var
  Stream: TResourceStream;
  vr: TVersionResource;
  fi: TVersionFixedInfo;

begin
  Result := '';
  try

    (* This raises an exception if version info has not been incorporated into the  *)
    (* binary (Lazarus Project -> Project Options -> Version Info -> Version        *)
    (* numbering).                                                                  *)

    Stream := TResourceStream.CreateFromID(HINSTANCE, 1, PChar(RT_VERSION));
    try
      vr := TVersionResource.Create;
      try
        vr.SetCustomRawDataStream(Stream);
        fi := vr.FixedInfo;
        Result := 'Version ' + IntToStr(fi.FileVersion[0]) + '.' +
          IntToStr(fi.FileVersion[1]) + ' Release ' +
          IntToStr(fi.FileVersion[2]) + ' Build ' + IntToStr(fi.FileVersion[3]);
        vr.SetCustomRawDataStream(nil)
      finally
        vr.Free
      end
    finally
      Stream.Free
    end
  except
  end;
end { resourceVersionInfo };

procedure TForm1.ShowVersionInfo;
begin

// nur so zum Test wie Git funktioniert

  (* Application Version anzeigen *)
  Application.Mainform.Caption :=
    'FAKS2 Fahrscheinverkäufe ' + ' [ ' + resourceVersionInfo + ' ]';
  Application.Title := Application.Mainform.Caption;

end;

function TForm1.JNochZuMelden(Sender: TObject): double;
var row : integer;
    OldSortedField : String;
    BM1 : TBookmark;
begin
  try
    Jei;
    BM1 := QFaks.GetBookmark;
    QFaks.DisableControls;
    OldSortedField := QFaks.SortedFields;
    QFaks.SortedFields := 'VertragsNr';
    QFaks.First;
    Result := 0;
    row := 0;

    (* gibts einzeln ausgewählte Zeilen? *)
    if DBGridFaks.SelectedRows.Count > 1 then
    begin
      for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
      begin
        QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));
        if QFaks.FieldByName('VertragsNr').IsNull   then
         Result := Result + QFaks.FieldByName('Betrag').AsCurrency;
      end;
    end
    else
    begin
      (* na gut, dann die ganze Tabelle *)
      while ((QFaks.FieldByName('VertragsNr').IsNull) and (not QFaks.EOF))  do
      begin
        Result := Result + QFaks.FieldByName('Betrag').AsCurrency;
        QFaks.Next;
      end;
    end;

  finally
    Nei;
    QFaks.SortedFields := OldSortedField;

    if QFaks.BookmarkValid(BM1) then QFaks.GotoBookmark(BM1);

    QFaks.FreeBookmark(BM1);
    QFaks.EnableControls;
  end;
end;

function TForm1.IsRIDinSelection(RID: string): boolean;
var
  row: integer;
begin
  try


    for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
    begin
      QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));
      Result := QFaks.FieldByName('RID').AsString = RID;
      if Result then
        exit;
    end;


  finally
  end;

end;

function TForm1.ShowFormSum(var Titel, Betrag: TStringList): boolean;
var
  x: integer;
begin
  try

    (* StringGrid auf zwei Zeilen reduzieren *)
    FormSum.StringGrid1.RowCount := 2;

    (* zweite Zeile Inhalt löschen *)
    FormSum.StringGrid1.Rows[1].Clear;

    (* Anzahl Zeilen einstellen *)
    FormSum.StringGrid1.RowCount := Titel.Count + 1;

    (* Spalte Titel füllen *)
    for x := 0 to Titel.Count - 1 do
    begin
      FormSum.StringGrid1.Cells[1, x + 1] := Titel[x];
    end;

    (* Spalte Betrag füllen *)
    for x := 0 to Betrag.Count - 1 do
    begin
      FormSum.StringGrid1.Cells[2, x + 1] := Betrag[x];
    end;


    FormSum.StringGrid1.AutoSizeColumns;

    (* Spalte etwas breiter *)
    FormSum.StringGrid1.ColWidths[1] := FormSum.StringGrid1.ColWidths[1] + 15;
    FormSum.StringGrid1.ColWidths[2] := FormSum.StringGrid1.ColWidths[2] + 15;


    FormSum.ShowModal;

  finally

  end;
end;

function TForm1.CheckFilterPossible(FieldName: String): Boolean;
begin
   (* Filter auf ein Feld nur mit Zeit geht nicht *)
    if ((POS('ZEIT',FieldName) > 0) And (FieldName <> 'DATUMZEIT'))then
    begin
      Result := false;

      DBGridFaks.SelectedField := QFaks.FieldByName('DATUMZEIT');
      Application.ProcessMessages;

      ShowMessage(
        'Das Feld ''' + FieldName + ''' läßt sich nicht filtern. Tipp: alles nach Excel und dort filtern!' + NL + NL +
        'Siehe auch den auskommentierten Teil im WHERE Ausdruck auf Seite Einstellungen' + NL +
        'Alternativ die Spalte ''DATUMZEIT'' nutzen, die jetzt ausgwählt wurde.');
    end
    else
    Result := true;



end;

function TForm1.LookUpStringGrid(Grid: TStringGrid; GesuchterMandant, ZuErsetzendeLinienNr: string): string;

  var r,c : Integer;
begin
  Result := ZuErsetzendeLinienNr;

  GesuchterMandant := trim (GesuchterMandant);
  ZuErsetzendeLinienNr := trim (ZuErsetzendeLinienNr);

  for r := 1 to Grid.RowCount -1 do
  begin
  //ShowMessage('Gesucht: GesuchterMandant: ' + GesuchterMandant + NL + 'ZuErsetzendeLinienNr: ' + ZuErsetzendeLinienNr + NL +  'in Zeichenfolge: ' + Grid.Rows[r].Text);
    if ((Grid.Cells[0,r] = GesuchterMandant) AND (Grid.Cells[1,r] = ZuErsetzendeLinienNr)) then
    begin
     Result := Grid.Cells[2,r];
    end;

  end;

end;

procedure TForm1.LinieJeMandant(Sender: TObject);
var list, DupLinie : TStringList;
    FName, StrLinie, item, item1 : String;
    Grid : TStringGrid;
    x, y, jcount : integer;

begin
  try
    Jei;
    FName := IncludeTrailingBackSlash(DirectoryEdit1.Directory) + 'Linien_je_Mandant.csv';

    (* wird Mandant und LinienNr aufnehmen *)
    list := TStringList.Create;

    (* wird die doppelten LinienNummern aufnehmen *)
    DupLinie := TStringList.Create;
    DupLinie.Sorted:=true;
    DupLinie.Duplicates:=dupIgnore;

    (* Notbehelf: damit nachher nach Spalte 2 sortiert werden kann *)
    Grid := TStringGrid.Create(Application.MainForm);


    ProgressBar1.Position:=0;
    ProgressBar1.Max:=QFaks.RecordCount;
    BM := QFaks.GetBookmark;

    QFaks.DisableControls;

    QFaks.First;

    (* List mit Mandant, Liniennummer füllen *)
    while not QFaks.EOF do
    begin

      (* idiotische Linie 31/32 korrigieren *)
      StrLinie := ExtractNumbers(QFaks.FieldByName('Linie').AsString);

      //assert(StrLinie='403','Linie ist 403');

      (* zu ersetzende LinienNummer nachschlagen, wenn nicht zu ersetzen als evtl Problemfall in list eintragen *)
      if LookUpStringGrid(GridReplaceLineNumber,QFaks.FieldByName('ID_F2MANDANT').AsString,StrLinie) =  StrLinie then
      begin
        item := QFaks.FieldByName('ID_F2MANDANT').AsString + ',' + QFaks.FieldByName('LINIE').AsString;
        if list.IndexOf(item) = -1 then list.Add(item);
      end;
      QFaks.Next;
      ProgressBar1.Position := QFaks.RecNo;
    end;

   (* Test ob Duplikat gefunden wird *)
   //list.insert(0,'0,90');
   //list.insert(0,'0,90');
   //list.insert(0,'0,95');

   (* Liste Sortieren *)
   list.Sort;


   list.insert(0,'Mandant,Linie');

   if list.Count = 1 then exit;

   (* jetzt list in Grid übernehmen, um nach Mandant zu sortieren.
      In Spalte Liniennummer sind dann Duplikate zu entdecken *)

   Grid.RowCount:=List.Count;

   for x := 0 to list.Count -1 do
   begin
     Grid.Rows[x].CommaText := list[x];
   end;

   Grid.SortColRow(true,1);

   (* doppelte Einträge suchen *)
   for x := 0 to Grid.RowCount -1 do
   begin
     jcount := 0;
     (* gesucht wird ... *)
     item := Grid.Cells[1,x];
     (* ... suche in *)
     for y := 0 to Grid.RowCount -1 do
     begin
       (* doppelt? *)
      if Grid.Cells[1,y] = item then
      begin
        inc(jcount);

        if jcount > 1 then
        DupLinie.Add('Linie: ' + item);
     end;


     end;



   end;


   if DupLinie.Count = 0 then
      ShowMessage('Keine doppelten LinienNummern, aber prüfen Sie selbst!!')
   else
      ShowMessage('Diese Linien sind doppelt (wird auch gleich in Excel angezeigt): ' + NL + NL + DupLinie.Text);

   (* In csv-Datei schreiben und öffnen *)
   Grid.SaveToCSVFile(FName,';');
   Nei;
   OpenURL(FName);


  finally
    Nei;
    FreeAndNil(List);
    FreeAndNil(Grid);
    FreeAndNil(DupLinie);

    if QFaks.BookmarkValid(BM) then
        QFaks.GotoBookmark(BM);

    QFaks.FreeBookmark(BM);

    QFaks.EnableControls;

    ProgressBar1.Position :=0;
    Application.ProcessMessages;


  end;

end;

function TForm1.AddFilterHistory(item: String): boolean;
begin
  Result := false;
  if FilterCombo.Items.IndexOf(Item) = -1 then
  begin
     FilterCombo.Items.Insert(0,item);
     FilterCombo.ItemIndex := 0;
     FilterCombo.Refresh;
     Application.ProcessMessages;
     Result := true;
  end
  else
  FilterCombo.ItemIndex := FilterCombo.Items.IndexOf(Item);
end;

procedure TForm1.ShowRecordCount(Sender: TObject);
begin
  (* Anzahl der Datensätze anzeigen *)
  lbRecordCount.Caption := FormatFloat('#,##0',QFaks.RecordCount) + ' Datensätze';

end;

function TForm1.AnalyzeTable(Table: string): boolean;
var OldSQL : string;
begin
  try
   (* Fehler ORA-01722 fixen *)
   Jei;
   (* alte SQL sichern *)
   OldSQL := ZUpDateRid.SQL.Text;
   ZUpDateRid.SQL.TEXT := 'analyze table ' + Table + ' VALIDATE STRUCTURE';
   ZUpDateRid.ExecSQL;
   EventLog1.Log('Fehler ORA-01722 fixen mit: ' + ZUpDateRid.SQL.TEXT);
  finally
    Nei;
    (* SQL rücksichern *)
    ZUpDateRid.SQL.Text := OldSQL;
    QFaks.Refresh;

  end
end;


// testet, ob eine Tabelle schon in der DB vorhanden ist
// benutzt dazu Metadata
function TForm1.ExistInDb(TableName: string): boolean;
var
  TableFound: boolean;
  ZSQLMetadata: TZSQLMetadata;
  DSSQLMetadata: TDataSource;
begin
  Result := False;
  ZSQLMetadata := TZSQLMetadata.Create(self);
  DSSQLMetadata := TDataSource.Create(self);
  ZSQLMetadata.Connection := ZConnection1;
  DSSQLMetadata.DataSet := ZSQLMetadata;

  // Spalten aus Tabelleninfo (Metadata) auslesen
  ZSQLMetadata.MetadataType := mdTables;
  ZSQLMetadata.Catalog := LowerCase(FTableDatabase);
  ZSQLMetadata.Open;
  TableFound := False;
  while not DSSQLMetadata.DataSet.EOF and (TableFound = False) do
  begin
    if LowerCase(DSSQLMetadata.DataSet.FieldByName('TABLE_NAME').Text) =
      LowerCase(TableName) then
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

function TForm1.ExistField(FieldName: string): boolean;
var
  x: integer;
begin
  try
    Jei;
    Result := False;
    ZCheckFields.Open;

    for x := 0 to ZCheckFields.FieldCount - 1 do
    begin
      if CompareText(ZCheckFields.Fields[x].FieldName, FieldName) = 0 then
      begin
        Result := True;
        break;
      end;
    end;

  finally
    ZCheckFields.Close;
    Nei;

  end;
end;

function TForm1.MarkExported(DBFName: string; mark: string; IDs: TStringList): boolean;
var
  tab: TDBF;
  Betrag: currency;
  v: variant;
  ErrLIst: TStringList;

begin
  Result := False;
  if not FileExists(DBFName) then
  begin
    Result := False;
    ShowMessage(DBFName + NL + 'wurde nicht gefunden!');
    exit;
  end;
  try
    jei;

    ErrLIst := TStringList.Create;


    tab := TDBF.Create(Application.Mainform);
    tab.TableName := DBFName;
    tab.Open;
    tab.First;

    BM := QFaks.GetBookmark;

    QFaks.DisableControls;
    ProgressBar1.Position := 0;
    ProgressBar1.Max := Tab.RecordCount;
    while not tab.EOF do
    begin
      v := Trim(tab.FieldByName('RID').AsString);
      ZUpDateRid.ParamByName('RID').AsString := v;

      (* NULL korrekt eintragen? *)
      if CompareText(mark,'NULL') <> 0 then
        ZUpDateRid.ParamByName('Vertragsnr').AsString := mark
      else
        ZUpDateRid.ParamByName('Vertragsnr').Clear;

      ZUpDateRid.ExecSQL;
      //if not QFaks.Locate('RID',v,[]) then ErrList.Add(v);
      tab.Next;
      ProgressBar1.Position := Tab.RecNo;
      ProgressBar1.Update;
    end;

    if ErrList.Count > 1 then
    begin
      ErrList.SaveToFile(EXEPath + 'FAKS_ErrList.txt');
      ShowMessage('Fehler sind aufgetreten bei RID, siehe ' + EXEPath +
        'FAKS_ErrList.txt');
    end;

  finally
    if QFaks.BookmarkValid(BM) then
      QFaks.GotoBookmark(BM);
    QFaks.FreeBookmark(BM);
    tab.Close;
    ProgressBar1.Position := 0;
    FreeAndNil(tab);
    FreeAndNil(ErrLIst);
    QFaks.EnableControls;
    QFaks.Refresh;
    nei;
  end;

end;

function TForm1.IsDate(str: string): boolean;
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

function TForm1.ShowFilterInfo(Warning: boolean): boolean;
begin
  //if ((trim(FilterCombo.Text) = '') or (not QFaks.Filtered)) then Warning := False;


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

  ShowRecordCount(nil);

  (* jetzt wirklich anzeigen? *)
  FilterCombo.Invalidate;
  FilterCombo.Refresh;
  Application.ProcessMessages;
  Application.MainForm.Invalidate;
  Application.ProcessMessages;
end;

procedure TForm1.BereitsGemeldetWurden(Sender: TObject);
var
  MyDBF: TDBF;
  OldFilterIndex, months, x, y, v: integer;
  SumZeitraum, SummeDBF, NochZuMelden: currency;
  aktuellerMonat, FolgeMonat, SQLFile, Filter, StrOr, FName: string;
  MyDirSelect: TSelectDirectoryDialog;
  RIDFaks, RIDGeloescht, RIDMonat, MyFiles, Titel, Betrag, AmisDataMissing: TStringList;
begin
    (* ja, schrecklich unübersichtlicher Code, ich kanns eben nicht besser *)

    (* bereits gemeldete Einnahmen zu Monat xy sind normalerweise in zwei Monatsdateien zu finden.
       Verkäufe vom Juni findet man also im Juni und im Juli.
       Das ist Folge von zu spät eingelesenen Druckern/Terminals
       Aus dem je betrachteten Monat darf nur der Datumsbereich von Interesse extrahiert werden,
       für den Juni 01.6.2015 bis 30.06.2015 *)

  (* ACHTUNG:
     es wird zwischen months=0 und months>0 im Code unterschieden!!!!!!!!!!!
     für months>0 siehe ungefähr Zeile 1090

     Ganz unten werden nachträgliche Stornos ( fälschlich VertragsNr=1 ) korrigiert
  *)

  try
    (* wie hoch sind die noch nicht an AmisData gemldeten Einnahmen: VertragsNr.IsNull *)
    try
      (* wird die Summen de Einnahmen in FormSum anzeigen *)
      Titel := TStringList.Create;
      Betrag := TStringList.Create;

      (* Nachträgliche Buchungen in FAKS mit falscher VertragsNr=1 , also nachzubuchen *)
      AmisDataMissing := TStringList.Create;


      QFaks.DisableControls;
      QFaks.Last;
      QFaks.First;
      ProgressBar1.Max := QFaks.RecordCount;

      StrOr := ' OR RID=';


      NochZuMelden := 0;

      StatusBar1.SimpleText := 'ermittle die noch an AmisData zu meldenden Einnahmen';

      while not QFaks.EOF do
      begin
        if QFaks.FieldByName('VertragsNr').IsNull then
          NochZuMelden := NochZuMelden + QFaks.FieldByName('Betrag').AsCurrency;

        ProgressBar1.Position := QFaks.RecNo;
        Application.ProcessMessages;
        if StoppIt then
        begin
          Nei;
          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
          StoppIt:= false;
          exit;
        end;

        QFaks.Next;

      end;


    finally

    end;


    Screen.ActiveControl.Invalidate;

    StoppIt := False;

     (* wird nur ein Monat oder mehrere Monate untersucht,
        für months>0 siehe ungefähr Zeile 835 *)
    months := MonthsBetween(DateEditVon.Date, DateEditBis.Date);


    Filter := '';


    (* wird die RID-Werte der Monate aufnehmen *)
    RIDMonat := TStringList.Create;
    RIDMonat.Sorted := True;

    (* wird die RID-Werte aus QFAKS aufnehmen *)
    RIDFaks := TStringList.Create;
    RIDFaks.Sorted := True;

    (* welche Rids sind in den Dbase Dateien aber nicht mehr in QFaks? *)
    RIDGeloescht := TStringList.Create;
    RIDGeloescht.Sorted := True;



    (* wird die Dateinamen der AmisData-DBFs aufnehmen *)
    MyFiles := TStringList.Create;

    (* Dbase Datei erzeugen, in die die Save Daten AmisData eingelesen werden *)
    MyDBF := TDBF.Create(Application);


    (* Dialog SelectDirectory erzeugen *)
    MyDirSelect := TSelectDirectoryDialog.Create(Application);

    SavePath := IncludeTrailingBackSlash(SavePath);

    if ((not DirectoryExists(SavePath)) or (trim(SavePath) = '\')) then
    begin

       (* einen halbwegs passenden Vorgabewert für Directory erzeugen,
          das tatsächliche Directory wird später in ini gespeichert *)
      MyDirSelect.Initialdir := DirectoryEdit1.Directory;

      MyDirSelect.Title :=
        'Wo stehen die bereits in Amisdata Schnittstelle BLE importierten Faks-Daten?';
      (* den Pfad der Dateien ermitteln *)
      if MyDirSelect.Execute then
      begin
        if DirectoryExists( MyDirSelect.FileName ) then
        SavePath := IncludeTrailingBackslash(MyDirSelect.FileName)
        else
        begin
           ShowMessage('Das Verzeichnis ''' + MyDirSelect.FileName + ''' existiert nicht!' + NL +
           'Bitte wiederholen Sie die letzte Aktion, damit der Pfad erneut abgefragt wird!');
           Nei;
           exit;
        end
      end
      else
        exit;

      Application.ProcessMessages;
    end;

    if months = 0 then
    begin

      (* aktuellerMonat und FolgeMonat zusammensetzten *)
      //aktuellerMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy', incMonth(Date,-1)) + '.dbf';
      aktuellerMonat := SavePath + 'FaksDaten_' +
        FormatDateTime('mmyy', DateEditVon.Date) + '.dbf';
      FolgeMonat := SavePath + 'FaksDaten_' + FormatDateTime('mmyy',
        incMonth(DateEditVon.Date, 1)) + '.dbf';




      if ((not FileExists(aktuellerMonat)) or
        (not FileExists(FolgeMonat))) then
      begin
        ShowMessage('Zumindest einer der zwei Monate:' + NL +
          ExtractFileName(aktuellerMonat) + NL + ExtractFileName(FolgeMonat) +
          NL + 'konnte in' + NL + SavePath +
          NL + ' nicht gefunden werden. Machen Sie den Vergleich bitte mit Stat.exe selber!' + NL +
          'Die Variable SavePath: ' + SavePath + ' steht in der ini-Datei dieser Anwendung!');
        exit;
      end;

      (* FolgeMonat untersuchen  *)
      SumZeitraum := 0;
      MyDBF.TableName := FolgeMonat;
      MyDBF.Open;
      MyDBF.First;
      ProgressBar1.Position := 0;
      ProgressBar1.Max := MyDBF.RecordCount;

      StatusBar1.SimpleText := 'ermittle die Einnahmen in: ''' +
        ExtractFileName(FolgeMonat) + '''';
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
          SumZeitraum := SumZeitraum + MyDBF.FieldByName('PREIS').AsCurrency;
        end;

        Application.ProcessMessages;
        if StoppIt then
        begin
          Nei;
          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
          StoppIt:= false;
          exit;
        end;


        MyDBF.Next;
        ProgressBar1.Position := MyDBF.RecNo;

      end;
      MyDBF.Close;

      (* aktuellerMonat untersuchen  *)
      SummeDBF := 0;
      MyDBF.TableName := aktuellerMonat;
      MyDBF.Open;
      MyDBF.First;
      ProgressBar1.Position := 0;
      ProgressBar1.Max := MyDBF.RecordCount;

      StatusBar1.SimpleText := 'ermittle die Einnahmen in: ''' +
        ExtractFileName(aktuellerMonat) + '''';
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
          SummeDBF := SummeDBF + MyDBF.FieldByName('PREIS').AsCurrency;
        end;

        Application.ProcessMessages;
        if StoppIt then
        begin
          Nei;
          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
          StoppIt:= false;
          exit;
        end;



        MyDBF.Next;
        ProgressBar1.Position := MyDBF.RecNo;

      end;
      MyDBF.Close;



      Nei;

      (* Stimmen die ermittelten Beträge überein? *)
      if SumZeitraum + SummeDBF + NochZuMelden = GesamtEinnahme then
      begin
        ShowMessage('Alles OK!' + NL + NL + '''' + ExtractFileName(FolgeMonat) +
          ''' ( ' + FormatFloat('#,##0.00', SumZeitraum) + ' )' + NL +
          '''' + ExtractFileName(aktuellerMonat) + ''' ( ' +
          FormatFloat('#,##0.00', SummeDBF) + ' )' + NL +
          'Davon noch zu melden:' + NL +
          FormatFloat('#,##0.00', NochZuMelden) + NL +
          'enthalten zusammen  ' + NL + FormatFloat('#,##0.00', SumZeitraum +
          SummeDBF + NochZuMelden) +  NL + 'Einnahmen im Zeitraum.');
      end
      else
      begin

       (* jetzt die Datensätze ermitteln, die der aktuelle Monat MEHR hat als die zwei Amisdata *.dbf
          Die zweite Möglichkeit: Datensätze sind in AmisData aber ncht mehr im aktuellen Monat untersuche ich nicht *)

        StatusBar1.SimpleText := 'Prüfe auf nicht gemeldete RID''s ... bitte warten';
        Application.ProcessMessages;

        try
          Filter := '';
          BM := QFaks.GetBookmark;
          QFaks.DisableControls;
          ProgressBar1.Position := 0;
          QFaks.Last;
          ProgressBar1.Max := QFaks.RecordCount;
          QFaks.First;
          jei;
          while not QFaks.EOF do
          begin
            if RIDMonat.IndexOf(QFaks.FieldByName('RID').AsString) = -1 then
            begin
              if Filter = '' then
              begin
                Filter := 'RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
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
              StoppIt:= false;
              exit;
            end;

          end;
        finally
          Nei;
          if QFaks.BookmarkValid(BM) then
            QFaks.GotoBookmark(BM);
          QFaks.FreeBookmark(BM);
          QFaks.EnableControls;
          QFaks.Refresh;
          ProgressBar1.Position := 0;
        end;

        mem.Clear;
        mem.Text := Filter;
        mem.SelectAll;
        mem.CopyToClipboard;

        FilterCombo.Text := Filter;
        QFaks.Filter := '(' + Filter + ') And Vertragsnr is null';
        QFaks.Filtered := True;
        ShowFilterInfo(True);

        (* zur Liste der Filter hinzufügen *)
        AddFilterHistory(QFaks.Filter);


        Titel.Clear;
        Betrag.Clear;

        Titel.Add('Gesamteinnahme in Faks:');
        Betrag.Add(FormatFloat('#,##0.00', GesamtEinnahme));

        Titel.Add('ABER:');
        Betrag.Add('');

        Titel.Add(ExtractFileName(FolgeMonat));
        Betrag.Add(FormatFloat('#,##0.00', SumZeitraum));

        Titel.Add(ExtractFileName(aktuellerMonat));
        Betrag.Add(FormatFloat('#,##0.00', SummeDBF));

        Titel.Add('ergibt als Summe');
        Betrag.Add(FormatFloat('#,##0.00', SumZeitraum + SummeDBF));

        Titel.Add('Die NICHT gemeldeten Daten werden jetzt angezeigt.');
        Betrag.Add('');

        ShowFormSum(Titel, Betrag);

      end;
      FreeAndNil(MyDirSelect);
      /////////////////////////////////////////////// ENDE ein Monat //////////////////////////////////////////////////////////////
    end (* months=0 *)
    else
    begin
      (* jetzt alle Monate durchlaufen, die RID's sammeln und in FAKS nachsehen ob vorhanden *)
      RIDMonat.Clear;



      SummeDBF := 0;
      Filter := '';

      for x := 0 to months do
      begin
        aktuellerMonat := SavePath + 'FaksDaten_' +
          FormatDateTime('mmyy', incMonth(DateEDitVon.Date, x)) + '.dbf';
        if not FileExists(aktuellerMonat) then
        begin
          ShowMessage('Datei nicht gefunden, Sie müssen das ggf. händisch erledigen!'
            + NL + NL + aktuellerMonat);
          continue;
        end;


        // months durchlaufen *******************************************
        if MyDBF.Active then
          MyDBF.Close;

        (* aktuellerMonat untersuchen  *)
        MyDBF.TableName := aktuellerMonat;
        MyDBF.Open;
        MyDBF.First;
        ProgressBar1.Position := 0;
        ProgressBar1.Max := MyDBF.RecordCount;

        StatusBar1.SimpleText := 'ermittle die Einnahmen in: ''' +
          ExtractFileName(aktuellerMonat) + '''';
        Application.ProcessMessages;

        SumZeitraum := 0;



        while not MyDBF.EOF do
        begin
          jei;
          (* RID in StringList eintragen *)
          RIDMonat.Add(MyDBF.FieldByName('RID').AsString);

          (* Einnahme zusammnenzählen, falls sie in den Zeitraum fällt *)
          if ((MyDBF.FieldByName('DATUMV').AsDateTime >= DateEditVon.Date) and
            (MyDBF.FieldByName('DATUMV').AsDateTime <= DateEditBis.Date)) then
          begin
            SummeDBF := SummeDBF + MyDBF.FieldByName('PREIS').AsCurrency;
            SumZeitraum := SumZeitraum + MyDBF.FieldByName('PREIS').AsCurrency;

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
          ProgressBar1.Position := MyDBF.RecNo;

          Application.ProcessMessages;
          if StoppIt then
          begin
            Nei;
            ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
            StoppIt:= false;

            exit;
          end;

        end;
        MyDBF.Close;

        (* wird mit Dialog später angezeigt *)
        MyFiles.Add(aktuellerMonat + ' [Summe im Zeitraum: ' +
          FormatDateTime('dd.mm.yy', DateEditVon.Date) + ' bis ' +
          FormatDateTime(
          'dd.mm.yy', DateEditBis.Date) + '] = ' + FormatFloat('#,##0.00', SumZeitraum));

      end (* for x *);


      Nei;

      (* Stimmen die ermittelten Beträge überein? *)
      if SummeDBF + NochZuMelden = GesamtEinnahme then
        ShowMessage('Alles OK!' + NL + NL + 'Details:' +
          NL + 'Hier in Faks angezeigt wurden:' + NL +
          FormatFloat('#,##0.00', GesamtEinnahme) + NL +
          'Nach AmisData exportiert wurden:' + NL + MyFiles.Text +
          NL + '... die enthalten zusammen:' + NL +
          FormatFloat('#,##0.00', SummeDBF) + NL +
          'zuzüglich noch der an AmisData zu meldenden Einnahmen ( Spalte VerragsNr is null) also,'
          +
          NL + FormatFloat('#,##0.00', NochZuMelden) + NL +
          'ergibt' + NL + FormatFloat('#,##0.00', GesamtEinnahme) +
          NL + 'Differenzbetrag:' + NL + FormatFloat(
          '#,##0.00', GesamtEinnahme - SummeDBF - NochZuMelden))



      else
      begin

            (* jetzt die Datensätze ermitteln, die der aktuelle Monat MEHR hat als die zwei Amisdata *.dbf
               Die zweite Möglichkeit: Datensätze sind in AmisData aber ncht mehr im aktuellen Monat untersuche ich nicht *)
        StatusBar1.SimpleText :=
          'Prüfe auf nicht gemeldete RID''s ... bitte warten: ' + IntToStr(RIDMonat.Count) +
          ' Datensätze in den Dbase Dateien';
        Application.ProcessMessages;
        Jei;

        try
          BM := QFaks.GetBookmark;
          QFaks.DisableControls;
          ProgressBar1.Position := 0;
          QFaks.Last;
          ProgressBar1.Max := QFaks.RecordCount;
          QFaks.First;
          StrOr := ' OR RID=';
          Filter := '';
          jei;
          Application.ProcessMessages;
          while not QFaks.EOF do
          begin

            (* RID's aus QFaks speichern *)
            RIDFaks.Add(QFaks.FieldByName('RID').AsString);

            (* nachsehen ob RID aus AmisData DBase jetzt in FAKS vorhanden ist
               grübel: dieser Filter ist doch Unsinn *)
            (*
            y := RIDMonat.IndexOf(QFaks.FieldByName('RID').AsString);
            if y = -1 then
            begin
              if Filter = '' then
              begin
                Filter := 'RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
              end
              else
              begin
                Filter :=
                  Filter + StrOr + QuotedStr(QFaks.FieldByName('RID').AsString);
              end;
            end;
            *)

            ProgressBar1.Position := QFaks.RecNo;
            Application.ProcessMessages;
            if StoppIt then
            begin
              Nei;
              ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
              StoppIt:= false;
              exit;
            end;
            QFaks.Next;
          end;
        finally
          Nei;
          if QFaks.BookmarkValid(BM) then
            QFaks.GotoBookmark(BM);
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

        if Filter <> '' then
        begin
          FilterCombo.Text := Filter;
          StatusBar1.SimpleText := StatusBar1.SimpleText +
          ' Jetzt wird der Filter gesetzt';
          Application.ProcessMessages;
          QFaks.Filter := '(' + Filter + ') And Vertragsnr=''1''';
          QFaks.Filtered := True;
          ShowFilterInfo(True);
          (* zur Liste der Filter hinzufügen *)
          AddFilterHistory(QFaks.Filter);

        end;

        (* nachsehen, welche Datensätze aus Faks gelöscht wurden *)
        for x := 0 to RIDMonat.Count - 1 do
        begin
          if RIDFaks.IndexOf(RIDMonat[x]) = -1 then
          begin
            if RIDGeloescht.Count = 0 then
              RIDGeloescht.Add('RID=' + QuotedStr(RIDMonat[x]))
            else
              RIDGeloescht.Add(StrOr + QuotedStr(RIDMonat[x]));

          end;

        end;

        (* nach Spalte Vetragsnr sortieren *)
        QFaks.SortedFields := 'Vertragsnr';
        QFaks.First;

        (* werden die angezeigten Daten bei der nächsten Meldung geholt? *)

        if ((GesamtEinnahme = SummeDBF +
          NochZuMelden) (*OR (QFaks.RecordCount = 0) *)) then
        begin
          ShowMessage('Scheint alles zu stimmen' + NL +
            'Eben ermittelt wurden: ' + NL +
            FormatFloat('#,##0.00', GesamtEinnahme) + NL +
            'Summe in den DBase Dateien Amisdata: ' + NL +
            FormatFloat('#,##0.00', SummeDBF) + NL +
            'Differenz:' + NL + FormatFloat(
            '#,##0.00', GesamtEinnahme - SummeDBF - NochZuMelden) + NL +
            'Achtung: noch an AmisData zu meldende Daten (Vertragsnr is null), d.h. '
            + FormatFloat('#,##0.00', NochZuMelden) +
            ' werden nicht angezeigt, sind aber in der ersten Summe enthalten!' +
            NL +
            'Es gab auch schon den verrückten Fall: es wurden in Faks nachträglich Daten gelöscht!!');
        end
        else
        begin

          Titel.Clear;
          Betrag.Clear;

          Titel.Add('Da stimmt was nicht:');
          Betrag.Add('');

          Titel.Add('Gesamteinnahme');
          Betrag.Add(FormatFloat('#,##0.00', GesamtEinnahme));

          Titel.Add('ABER:');
          Betrag.Add('');

          Titel.Add('');
          Betrag.Add('');

          (* Dateinamen und Betrag aus MyFiles extrahieren *)
          for v := 0 to MyFiles.Count - 1 do
          begin
            Titel.Add(ExtractWord(1, MyFiles[v], ['=']));
            Betrag.Add(ExtractWord(2, MyFiles[v], ['=']));
          end;

          Titel.Add('enthalten zusammen:');
          Betrag.Add(FormatFloat('#,##0.00', SummeDBF));

          Titel.Add('Differenzbetrag:');
          Betrag.Add(FormatFloat('#,##0.00', GesamtEinnahme - SummeDBF));

          Titel.Add('Davon noch nicht an AmisData gemeldet:');
          Betrag.Add(FormatFloat('#,##0.00', NochZuMelden));

          Titel.Add('bleiben ungeklärt, z.B. alter Tarif:');
          Betrag.Add(FormatFloat('#,##0.00',
            GesamtEinnahme - SummeDBF - NochZuMelden));

          Titel.Add('');
          Betrag.Add('');

          Titel.Add(
            'Es kam schon vor das Daten in Faks nachtraeglich geloescht wurden!!');
          Betrag.Add('');

          Titel.Add(
            'Die RID''s in FAKS nicht mehr vorhandener Daten werden ggf. gleich angezeigt,');
          Betrag.Add('');

          Titel.Add('erfahrungsgemaess ist das dann alter Tarif aus Vorjahr,');
          Betrag.Add('');

          Titel.Add('oder der Datumsbereich ist jetzt ein anderer (->nachtraeglich in AmisData eingelesen),');
          Betrag.Add('');

          Titel.Add('also Verkaeufe aus Januar im AmisData März');
          Betrag.Add('');


          Titel.Add('Bitte einen moeglichst grossen Zeitraum waehlen: 1.1.' +
            IntToStr(YearOf(DateEDitBis.Date)) + ' bis ' + DateEDitBis.Text);
          Betrag.Add('');

          Titel.Add('');
          Betrag.Add('');

          Titel.Add('Achtung: TarifNummer ist auch eine böse Falle!! Aktuell ist sie: ' + IntToStr(SpinEditTarifversion.Value));
          Betrag.Add('');



          ShowFormSum(Titel, Betrag);

          if RIDGeloescht.Count > 0 then
          begin
            (* Damit RID= der erste Eintrag der Liste wird *)
            RIDGeloescht.Sorted := False;
            RIDGeloescht.Move(RIDGeloescht.Count - 1, 0);

            (* als Textdatei speichern *)
            if DirectoryExists(Form1.DirectoryEdit1.Text) then
            begin
              FName := IncludeTrailingBackslash(Form1.DirectoryEdit1.Text) +
                'RIDGeloescht.txt';
            end
            else
            begin
              FName := ExePath + 'RIDGeloescht.txt';
            end;

            (* Kommentar und Datum in erste Zeile *)
            RIDGeloescht.Insert(0, 'Diese RID fehlen jetzt in FAKS, Stand: ' + FormatDateTime('dd.mm.yy hh:nn', now));
            RIDGeloescht.Insert(1, 'Zeitraum: ' + DateEditVon.Text + ' bis ' + DateEditBis.Text);
            RIDGeloescht.Insert(2, '');

            RIDGeloescht.SaveToFile(FName);


            if not OpenUrl(FName) then
              ShowMessage('Diese RID fehlen jetzt in FAKS: kopieren = Strg + c '
                + sLineBreak + RIDGeloescht.Text);
          end;

        end;

// XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX VertragsNr = 1, also nachträgliche Stornos

        (* Versuch: was ist noch nicht in AmisData, hat aber VertragsNR=1, also z.B. nachträglich Stornos *)
        StatusBar1.SimpleText := 'was ist noch nicht in AmisData, hat aber VertragsNR=1, also z.B. nachträglich Stornos';
        Application.ProcessMessages;

        (* Alle RID von QFaks durchlaufen ... *)
        for x := 0 to RIDFaks.Count -1 do
        begin

           ProgressBar1.Position:=x;
           Application.ProcessMessages;

           (* ... nachsehen, ob die RID in DBF vorhanden *)
           if RIDMonat.IndexOf(RIDFaks[x]) = -1 then
           begin
             (* In QFaks nachschlagen ob VertragsNr=1 *)
             if QFaks.Lookup('RID',VarArrayOf([RIDFaks[x]]),'VERTRAGSNR')='1' then
                AmisDataMissing.Add(RIDFaks[x]);
           end;


        end;

        (* Filter zusammenstellen *)
        if AmisDataMissing.Count > 0 then
        begin
         if QFaks.Filtered then QFaks.Filtered:=false;

         Filter := '';

         for x := 0 to AmisDataMissing.Count -1 do
         begin

           if Filter = '' then
           begin
             Filter := 'RID=' + QuotedStr(AmisDataMissing[x]);
           end
           else
           begin
             Filter :=
               Filter + StrOr + QuotedStr(AmisDataMissing[x]);
           end;

        end;

         FilterCombo.Text := Filter;
         QFaks.Filter:=Filter;
         QFaks.Filtered:=true;
         ShowFilterInfo(True);
         (* zur Liste der Filter hinzufügen *)
         AddFilterHistory(QFaks.Filter);





         ShowMessage('Summe Betrag: ' + FormatFloat('#,##.00',GesamtEinnahme - SummeDBF - NochZuMelden) + NL + NL +
         'Diese Daten  fehlen in Amisdata (Filter wurde gesetzt!):'  + NL + NL + Filter);

         if Messagedlg('Vorsicht!!!' + NL + NL +
            'Soll die VertagsNr für die eben angezeigten Datensätze zurückgesetzt werden, um sie neu in AmisData zu importieren?',mtConfirmation,[mbYes,mbNo],0)= mrYes then
            begin
              (* Id's der nachträglichen Stornos in Datei speichern *)
              if not FileExists(ExePath + 'AmisDataMissing.txt') then
                AmisDataMissing.SaveToFile(ExePath + 'AmisDataMissing.txt')
              else
              begin
                Titel.LoadFromFile(ExePath + 'AmisDataMissing.txt');
                Titel.AddStrings(AmisDataMissing);
                Titel.SaveToFile(ExePath + 'AmisDataMissing.txt')
              end;

              for x := 0 to AmisDataMissing.Count -1 do
              begin
                StatusBar1.SimpleText:='aktualisiere: ' + AmisDataMissing[x];
                Application.ProcessMessages;
                ZUpDateRid.ParamByName('RID').AsString := AmisDataMissing[x];
                ZUpDateRid.ParamByName('VertragsNr').AsString := '';
                ZUpDateRid.ExecSQL;

              end;


              StatusBar1.SimpleText:='Daten neu einlesen ...';
              Application.ProcessMessages;

               QFaks.Refresh;
               DBGRidFaks.SelectedField := QFaks.FieldByName('VertragsNr');
               ShowMessage('Bei den angezeigten ''' + IntToStr(AmisDataMissing.Count) + ''' Datensätzen wurde die VertragsNr zurückgesetzt!' + NL +
               'Siehe auch ( gesamt!! ):' + ExePath + 'AmisDataMissing.txt');
            end;
        end;
      end;

    end;

  finally
    Nei;
    StatusBar1.SimpleText := '';
    ProgressBar1.Position := 0;
    FreeAndNil(MyDBF);
    FreeAndNil(RIDMonat);
    FreeAndNil(MyFiles);
    FreeAndNil(RIDFaks);
    FreeAndNil(RIDGeloescht);
    FreeAndNil(Titel);
    FreeAndNil(Betrag);
    FreeAndNil(AmisDataMissing);
  end;

end;

procedure TForm1.CheckIniFile(Sender: TObject);
var
  ini1, ini2: TIniFile;
  List1, List2: TStringList;
  FName, Section: string;
begin
   (* kontrollieren, ob LastBuchungsdatum auch in
      Amisdata-FAKS_Meldung.ini steht und ggf. nachtragen  *)
  FName := 'F:\AMISdata\Monat\Monatsmld\FAKS_Meldung.ini';
  Section := 'TApplication.Form1.Edit_last_buchung_Items';

  (* nur Kontrolle wenn lokal ausgeführt wird und AmisData vorhanden ist *)
  if ((Pos('PROJEKTE', ExePath) > 0) and (FileExists(FName))) then
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

      ini1.ReadSectionValues(Section, List1);
      ini2.ReadSectionValues(Section, List2);


      if not SameText(List1.CommaText, List2.CommaText) then
      begin
        (* wenn die sections nicht gleich sind, ini der lokalen Anwendung kopieren *)
        Mem.Clear;
        Mem.Text := List1.Text;
        (* Section in eckigen Klammern oben drüber *)
        Mem.Text := '[' + Section + ']' + NL + Mem.Text;
        Mem.SelectAll;
        Mem.CopyToClipboard;

        ShowMessage('Die ini-Einträge sind nicht gleich:' + NL +
          ini1.FileName + NL + List1.CommaText + NL + NL + ini2.FileName +
          NL + List2.CommaText + NL + NL +
          'Werte wurden kopiert! ZWEI(!!) Editoren werden geöffnet.');

        (* ini im Anwendungs-Verzeichnis im Editor öffnen *)
        shellExecute(Application.MainForm.Handle, 'open', PChar(
          INIPropStorage1.IniFileName), '', '', SW_NORMAL);
        (* ini in Amisdata-Verzeichnis im Editor öffnen *)
        shellExecute(Application.MainForm.Handle, 'open', PChar(FName), '', '', SW_NORMAL);

      end;



    finally
      FreeAndNil(List1);
      FreeAndNil(List2);
      FreeAndNil(Ini1);
      FreeAndNil(Ini2);

    end;

  end;

end;




procedure TForm1.angezeigteDatenkopieren1Click(Sender: TObject);
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
    StoppIt := False;

    mem1 := TStringList.Create;


    (* Dieser Code wird  für QFaks und QUmsetzung benutzt *)
    //Q := (screen.ActiveControl as TDBGrid).DataSource.DataSet as TZReadOnlyQuery;

    Screen.Cursor := crHourglass;
    QFaks.DisableControls;
    bm := QFaks.GetBookmark;
    QFaks.Last;
    QFaks.First;

    ProgressBar1.Position := 0;
    ProgressBar1.Max := QFaks.RecordCount;




    // First we send the data to a memo
    // works faster than doing it directly to Excel

    //das geht bei lazarus NICHT!!!  mem := TMemo.Create(Self);
    //mem.Parent := Form1;

    Form1.mem.Visible := True;
    Form1.mem.Clear;
    sline := '';

    // add the info for the column names
    for col := 0 to QFaks.FieldCount - 1 do
      sline := sline + QFaks.Fields[col].DisplayLabel + #9;

    (* letzten Tabulator entfernen *)
    sline := copy(sline, 0, length(sline) - 1);

    mem1.Add(sline);

    (* gibts MultiSelect im Grid oder ganzes Grid kopieren *)
    if (screen.ActiveControl as TDBGrid).SelectedRows.Count > 1 then
    begin
      (* also Multiselect! *)
      statusbar1.SimpleText := 'Das DBGrid ''' + (screen.ActiveControl as TDBGrid).Name +
        ''' hat ' + IntToStr((screen.ActiveControl as TDBGrid).SelectedRows.Count) +
        ' Zeilen selektiert!';

      // get the data into the memo
      for row := 0 to (screen.ActiveControl as TDBGrid).SelectedRows.Count - 1 do
      begin
        sline := '';
        QFaks.GotoBookmark(TBookMark((screen.ActiveControl as TDBGrid).SelectedRows[row]));
        ;
        for col := 0 to QFaks.FieldCount - 1 do
          sline := sline + QFaks.Fields[col].AsString + #9;

        (* letzten Tabulator rntfernen *)
        sline := copy(sline, 0, length(sline) - 1);

        mem1.Add(sline);
        //QFaks.Next;
      end;

    end
    else
    begin
      //if Messagedlg('Alle Daten sollen kopiert werden?' + NL + NL +
      //   'Durch ESC-Taste abbrechen!',mtConfirmation,[mbYes,mbNo],0)= mrNo then exit;
      (* kein Multiselect, also ganzes Grid *)
      statusbar1.SimpleText := 'Das ganze Grid wird jetzt kopiert!';
      // get the data into the memo
      for row := 0 to QFaks.RecordCount - 1 do
      begin
        ProgressBar1.Position := QFaks.RecNo;
        Application.ProcessMessages;
        if StoppIt then
        begin

          ProgressBar1.Position := 0;
          Application.ProcessMessages;

          ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
          StoppIt := False;
          exit;
        end;
        sline := '';
        for col := 0 to QFaks.FieldCount - 1 do
          sline := sline + QFaks.Fields[col].AsString + #9;

        (* letzten Tabulator entfernen *)
        sline := copy(sline, 0, length(sline) - 1);


        mem1.Add(sline);
        QFaks.Next;
      end;

    end;


    // we copy the data to the clipboard
    Form1.mem.Text := mem1.Text;
    Form1.mem.SelectAll;
    Form1.mem.CopyToClipboard;
    Form1.mem.Clear;
    Form1.mem.Visible := False;



    ProgressBar1.Position := 0;
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

function TForm1.OpenLog(FName: string): boolean;
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

procedure TForm1.ShortenLog(anzLines: integer; Sender: TObject);
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


      while log.Count >= anzLines do
        log.Delete(0);

      log.SaveToFile(ZSQLMonitor1.FileName);

      ZSQLMonitor1.Active := True;

    end;

  finally
    FreeAndNil(log);
  end;

end;


procedure TForm1.FormCreate(Sender: TObject);
var
  lst: TStringList;
  x: integer;
  OldDelimiter, s: string;
begin

  ExePath := ExtractFilePath(Application.ExeName);
  INIPropStorage1.IniFileName := ChangeFileExt(Application.ExeName, '.ini');
  ZSQLMonitor1.FileName := ChangeFileExt(Application.ExeName, '.log');

  (* speichert die History der ausgeführten SQL-Befehle *)
  SQLHistory := TStringList.Create;



  (* Log für Dbase-Exporte,
     wichtig: LogType muss ltFile sein, sonst wird in Systemlog geschrieben *)
  EventLog1.FileName := ChangeFileExt(Application.ExeName, '_DBase_Exporte.log');
  EventLog1.AppendContent := True;
  EventLog1.LogType := ltFile;
  EventLog1.Active := True;



  (* Database=hlbst02:1521/VASP macht das Gezerre mit tnsnames.ora überflüssig

    gibts die tnsnames.ora, sonst Oracle Fehlermeldung *)
  //if not FileExists(ExePath + 'tnsnames.ora') then
  //begin
  //  ListBox_tnsnames_ora.Items.SaveToFile(ExePath + 'tnsnames.ora');
  //  ShowMessage('Die Oracle Konfigurationsdatei ''' +
  //    ExePath + 'tnsnames.ora'' konnte nicht gefunden werden und wurde neu erzeugt. Einträge ggf. überprüfen!!');
  //end;

  (* soll die ID's aufnehmen, die bei Multiselect Rows in QFaks ausgewählt wurden *)
  RIDS_in_Selection := TStringList.Create;
  RIDS_in_Selection.Sorted := True;

  (* LinienNummern, die beim DBF-Export für AmisData ersetzt werden sollen *)
  if FileExists(ChangeFileExt(Application.ExeName,'_ErsetzungLinienNummern.csv')) then
  GridReplaceLineNumber.LoadFromCSVFile(ChangeFileExt(Application.ExeName,'_ErsetzungLinienNummern.csv'),';');


  (* Gattungsarten, die beim DBF-Export für AmisData ersetzt werden sollen *)
  if FileExists(ChangeFileExt(Application.ExeName,'_ErsetzungGattungsarten.csv')) then
  GridReplaceGattungsart.LoadFromCSVFile(ChangeFileExt(Application.ExeName,'_ErsetzungGattungsarten.csv'),';');

  PageControl1.ActivePage := TabDaten;

end;

procedure TForm1.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  if ((Key = VK_F3) and (gesucht <> '')) then
  begin
    if Screen.ActiveControl.Name = DBGridFaks.Name then
      Pm_SearchClick(Sender)
    else if Screen.ActiveControl.Name = GridReplaceLineNumber.Name then
      MnSearchClick(Sender);
  end
  else if (Key = VK_ESCAPE) then
  begin
    StoppIt := True;
    //raise exception.create('Hurrah Sie haben die ESCAPE-Taste getroffen :-)) ');
  end;

end;

procedure TForm1.IniPropStorage1RestoreProperties(Sender: TObject);
begin
  TarifVersion := SpinEditTarifversion.Value;

end;

procedure TForm1.IniPropStorage1StoredValues1Restore(Sender: TStoredValue;
  var Value: TStoredType);
begin
  SavePath := Value;
  (* Fehler verhindern *)
  if trim(SavePath) = '\'  then
     SavePath := '';
end;

procedure TForm1.IniPropStorage1StoredValues1Save(Sender: TStoredValue;
  var Value: TStoredType);
begin
  (* in ini Speichern Amisdata Pfad Save *)
  Value := IncludeTrailingBackslash(SavePath);
end;

procedure TForm1.IniPropStorage1StoredValues2Restore(Sender: TStoredValue;
  var Value: TStoredType);
begin
  if Value = '' then Value := FormatDateTime('dd.mm.yyyy', incMonth(Date,-1));
end;

procedure TForm1.kopierenClick(Sender: TObject);
begin
  angezeigteDatenkopieren1Click(Sender);
end;

procedure TForm1.ListBoxKnownLinesKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if Key = VK_DELETE then
    ListBoxKnownLines.DeleteSelected;

end;

procedure TForm1.ListFieldsClick(Sender: TObject);
var
  x: integer;
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


    // end;

    //if ExistInDb('f2fsv') then  ShowMessage('f2fsv gefunden');

    if ExistField('journal') then
      ShowMessage('Feld Journal wurde gefunden');

  finally
    Nei;

  end;
end;

procedure TForm1.LookupStornosClick(Sender: TObject);
var
  Filter: string;
  x, y: integer;
  Betrag: currency;
begin
  try
    if Messagedlg(
      'Diese Aktion wird länger dauern, kann aber im ersten Teil durch ESC-Taste abgebrochen werden.',
      mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      exit;

    Jei;

    QFaks.Filtered:=false;
    FilterCombo.Text:= '';
    ShowFilterInfo(False);


    StoppIt := False;


    BM := QFaks.GetBookmark;
    QFaks.DisableControls;


    QFaks.SortedFields := 'Betrag';
    QFaks.Last;
    QFaks.First;
    y := QFaks.RecordCount;


    Filter := '';

    while QFaks.FieldByName('Betrag').AsCurrency < 0 do
    begin

      Application.ProcessMessages;
      if StoppIt then
      begin
        ShowMessage('Bearbeitung durch ESC-Taste abgebrochen!');
        StoppIt := False;
        exit;
      end;

      (* Feld BelegNr UND Bemerkung sind bei Stornos identisch *)
      //(MDEIDINTERN='5012' AND (BELEGNR='7141' OR Bemerkung='7141')) OR (MDEIDINTERN='5005' AND (BELEGNR='6637' OR Bemerkung='6637'))
      if Filter = '' then
      begin
        Filter := '(MDEIDINTERN=' +
          QuotedStr(QFaks.FieldByName('MDEIDINTERN').AsString) + ' AND (BELEGNR=' +
          QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + ' OR Bemerkung=' +
          QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + '))';
      end
      else
      begin
        Filter := Filter + ' OR ' + '(MDEIDINTERN=' +
          QuotedStr(QFaks.FieldByName('MDEIDINTERN').AsString) + ' AND (BELEGNR=' +
          QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + ' OR Bemerkung=' +
          QuotedStr(QFaks.FieldByName('Bemerkung').AsString) + '))';
      end;



      QFaks.Next;
    end;

    mem.Clear;
    mem.Text := Filter;
    mem.SelectAll;
    mem.CopyToClipboard;
    //mem.Lines.SaveToFile(ExePath + 'MyFilter.txt');

    //ShowMessage(Filter);

    FilterCombo.Text := Filter;
    QFaks.Filter := Filter;
    QFaks.Filtered := True;
    ShowFilterInfo(True);

    x := QFaks.RecordCount;

    QFaks.SortedFields := 'Belegnr, Bemerkung';

    (* Welchen Wert hatten die stornierten Fahrscheine? *)
    QFaks.First;
    Betrag := 0;
    while not QFaks.EOF do
    begin
      if QFaks.FieldByName('Betrag').AsCurrency > 0 then
        Betrag := Betrag + QFaks.FieldByName('Betrag').AsCurrency;
      QFaks.Next;
    end;


    ShowMessage('Im Zeitraum: ' + DateEditVon.Text + ' bis ' +
      DateEditBis.Text + ' wurden:' + NL + FormatFloat('#,##0', x / 2) +
      ' von ' + IntToStr(y) + ' Fahrscheinen storniert.' + NL + 'Das entspricht ' +
      FormatFloat('#,##0.00 %', ((x / 2) / y) * 100) + NL + 'Summe: ' +
      FormatFloat('#,##0.00', Betrag) + NL + NL + 'Der Filter wurde kopiert!');


  finally
    Nei;
    if QFaks.BookmarkValid(BM) then
      QFaks.GotoBookmark(BM);
    QFaks.EnableControls;

    ShowRecordCount(Sender);
  end;
end;

procedure TForm1.MenuAddLinieClick(Sender: TObject);
var
  UserString: string;
begin
  UserString := InputBox('Neue Liniennummer erfassen:',
    'die neue Liniennummer ist:', '0');

  if ((UserString <> '0') and (ListBoxKnownLines.Items.IndexOf(UserString) = -1)) then
    ListBoxKnownLines.Items.add(UserString);

end;

procedure TForm1.MenuItem1Click(Sender: TObject);
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
procedure TForm1.MenuItem2Click(Sender: TObject);
var
  tab: TDBF;
  Betrag: currency;
  v: variant;
  ErrLIst: TStringList;
begin
  try
    jei;

    ErrLIst := TStringList.Create;


    tab := TDBF.Create(Application.Mainform);
    tab.TableName := 'F:\AMISdata\Monat\Monatsmld\FAKS_Auswertungen\RID_as_imported.dbf';
    tab.Open;
    tab.First;

    QFaks.DisableControls;
    ProgressBar1.Position := 0;
    ProgressBar1.Max := Tab.RecordCount;
    while not tab.EOF do
    begin
      v := Trim(tab.FieldByName('RID').AsString);
      ZUpDateRid.ParamByName('RID').AsString := v;
      ZUpDateRid.ParamByName('Vertragsnr').AsString := '1';
      ZUpDateRid.ExecSQL;
      //if not QFaks.Locate('RID',v,[]) then ErrList.Add(v);
      tab.Next;
      ProgressBar1.Position := Tab.RecNo;
      ProgressBar1.Update;
    end;

    if ErrList.Count > 1 then
    begin
      ErrList.SaveToFile(EXEPath + 'ErrList.txt');
      ShowMessage('Siehe ' + EXEPath + 'ErrList.txt');
    end;

  finally
    tab.Close;
    FreeAndNil(tab);
    FreeAndNil(ErrLIst);
    QFaks.EnableControls;
    QFaks.Refresh;
    nei;
  end;

end;

procedure TForm1.MnAnrufsammeltaxisClick(Sender: TObject);
var  MissingGAIDENT : TStringlist;
    x : Integer;
begin
  try
   (* Welche unbekannten Gattungsarten gibt es, d.h sind in
      GridReplaceGattungsart noch nicht erfasst.

   *)
   MissingGAIDENT := TStringlist.Create;
   MissingGAIDENT.Sorted:=true;
   MissingGAIDENT.Duplicates:=dupIgnore;

  (* geht nur mit extra SQL laut cbSQL.Checked *)
  if cbSQL.Checked then
  begin

    if QFaks.Filtered then  QFaks.Filtered := False;

    (* Linien Anrufsammeltaxi als Filter setzen *)
    FilterCombo.Text:='Bemerkung2 LIKE ''AST*'' OR MDEID LIKE ''*Notfahr*''';

    QFaks.Filter := FilterCombo.Text;

    QFaks.Filtered := True;

    AddFilterHistory(FilterCombo.Text);


    ShowFilterInfo(True);

    ShowRecordCount(Sender);

    Jei;

    QFaks.First;

    While not QFaks.EOF do
    begin
      if GridReplaceGattungsart.Cols[1].IndexOf(QFaks.FieldByName('GAIDENT').AsString) = -1 then
      MissingGAIDENT.Add(QFaks.FieldByName('GAIDENT').AsString);
      QFaks.next;
    end;

    if MissingGAIDENT.Count > 0 then
    begin
     Nei;
     PageControl1.ActivePage := TabConfig;
     GridReplaceGattungsart.SetFocus;
     Application.ProcessMessages;

     ShowMessage('Diese Gattungsarten fehlen in der Liste auf der Seite Einstellungen:' + NL +
     MissingGAIDENT.Text + NL + 'Strg + c kopiert diese Werte' + NL +
     'Setzen Sie später in den Daten einen Filter auf diese GAIDENT um die Sortennummern zu sehen.');

    end;

  end
  else
     ShowMessage('Bitte erst die Option auf der Seite Einstellungen ''mit Anrufsammeltaxi Namen'' aktivieren und neu starten!');

 finally
   Nei;
   FreeAndNil(MissingGAIDENT);
 end;

end;

procedure TForm1.MNCheckLinieJeMandantClick(Sender: TObject);
begin
  LinieJeMandant(Sender);
end;

procedure TForm1.MnCheckTarifPreiseClick(Sender: TObject);
var RID : string;
    list : TStringLIst;
begin
 (* mit Sortennumemer, Einzelpreis, PeisStDruck und Tarifversion in Amisdata Tabelle Preisliste
    den aktuellen Preis nachschlagen und bei Differenz den Datensatz rausfiltern *)
  try
  jei;
  if not ZConnection2.Connected then
    begin
      ZConnection2.Connect;
      QPreisliste.Active:=true;
    end;

    BM := QFaks.GetBookmark;

    QFaks.DisableControls;

    ProgressBar1.Position:=0;
    ProgressBar1.Max:= QFaks.RecordCount;

    QFaks.First;

    while not QFaks.EOF do
    begin
      StatusBar1.SimpleText:= 'Datensatz: ' + IntToStr(QFaks.RecNo);
      ProgressBar1.Position:=QFaks.RecNo;
      Application.ProcessMessages;
      (* Wenn Feld.IsNull überspringen *)
      if (QFaks.FieldByName('Sortennummer').IsNull
          or
          QFaks.FieldByName('Einzelpreis').IsNull
          or
          QFaks.FieldByName('TarifVersion').IsNull
          or
          QFaks.FieldByName('PreisStDruck').IsNull
          or
          QFaks.FieldByName('Sortennummer').IsNull
          or
          (length(QFaks.FieldByName('PreisStIdent').Value) > 2)
          or
          (QFaks.FieldByName('Betrag').Value < 0)
          or
          (QFaks.FieldByName('Datum').AsDateTime > StrToDate('10.12.' + IntToStr(Yearof(DateEditBis.Date))))
          ) then
          begin
           QFaks.Next;
           continue;
          end;

      if QFaks.FieldByName('Einzelpreis').AsCurrency <>
         QPreisliste.Lookup('RMVTarifNr,Gattungsart,Preisstufe',VarArrayOf([
         SpinEditTarifversion.Value,QFaks.FieldByName('Sortennummer').AsInteger,
         QFaks.FieldByName('PreisStDruck').AsInteger]),'FAHRPREIS')
         then
      begin
         (* zwecks Kontrolle in FAKS_Meldung_DBase_Exporte.log schreiben *)
         //if QFaks.FieldByName('Einzelpreis').AsCurrency = 2.6 then
         //EventLog1.Log('RID=' + QuotedStr(QFaks.FieldByName('RID').AsString)+ ' | ' + 'Einzelpreis=' +
         //    QFaks.FieldByName('Einzelpreis').AsString + ' | ' + 'Fahrpreis=' + FormatFloat('#,##0.00',
         //    QPreisliste.Lookup('RMVTarifNr,Gattungsart,Preisstufe',VarArrayOf([
         //QFaks.FieldByName('TarifVersion').AsInteger,QFaks.FieldByName('Sortennummer').AsInteger,
         //QFaks.FieldByName('PreisStDruck').AsInteger]),'FAHRPREIS'))
         //+ ' | ' + 'TarifVersion=' + QFaks.FieldByName('TarifVersion').AsString + ' | ' + 'Sortennummer=' +
         //QFaks.FieldByName('Sortennummer').AsString + ' | ' + 'PreisStDruck=' + QFaks.FieldByName('PreisStDruck').AsString
         //);

         if RID='' then
           RID := 'RID=' + QuotedStr(QFaks.FieldByName('RID').AsString)
         else
           RID := RID +  ' OR RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);

         QFaks.Next;
         ProgressBar1.Position:=QFaks.RecNo;
         Application.ProcessMessages;

         continue;

      end;

      QFaks.Next;

   end;

    QFaks.EnableControls;

    if RID<>'' then
    begin

    if (Length(RID) > 3000) then
    begin
      try
        List := TStringList.Create;
        List.Add('Bei diesen Datensaetzen passt der Preis nicht zur Tarifversion: ' + IntToStr(SpinEditTarifversion.Value));
        List.Add('');
        List.Add(RID);
        List.SaveToFile(ExePath + 'Faks_Meldung_falscher_Preis.txt');
        Nei;
        OpenUrl(ExePath + 'Faks_Meldung_falscher_Preis.txt');

      finally
        FreeAndNil(List);
        RID := '';

      end;
       exit;
    end;

    QFaks.Filter:=RID;

    QFaks.Filtered := True;

    (* nur zur Liste hinzufügen, falls noch nicht darin enthalten *)
    AddFilterHistory(RID);
    ShowFilterInfo(true);

    ShowRecordCount(Sender);

    Application.ProcessMessages;


     ShowMessage('Bei den angezeigten ' + IntToStr(QFaks.RecordCount) + ' Datensätzen entspricht der Preis nicht der '
     + 'Tarifversion ' + SpinEditTarifversion.Caption);
    end
    else
    begin
     Nei;
     ShowMessage('Alles ok' + NL + NL + 'Alle Preise stehen so auch in AmisData für Tariversion:  '+ IntToStr(SpinEditTarifversion.Value));
    end;

    if QFaks.BookmarkValid(BM) then
    begin
      QFaks.GotoBookmark(BM);
      QFaks.FreeBookmark(BM);
    end;


  finally
    Nei;
    ProgressBar1.Position:=0;
    Application.ProcessMessages;
  end;
end;

procedure TForm1.mnDelFilterClick(Sender: TObject);
begin
  RemoveFilterClick(Sender);
end;

procedure TForm1.MnFieldlistClick(Sender: TObject);
begin
  (* Feldliste anzeigen: SQL Code aus Faks_Meldung.log *)
  if not Form2.ZReadOnlyQuery1.Active then Form2.ZReadOnlyQuery1.Active:=true;
  Form2.ShowModal;
end;

procedure TForm1.MnFixColumnClick(Sender: TObject);
begin
  if DBGridFaks.FixedCols = 1 then
    (* keine Ahnung warum hier plus 2 *)
    DBGridFaks.FixedCols:= DBGridFaks.SelectedColumn.Index +2
  else
    DBGridFaks.FixedCols:=1;
end;

procedure TForm1.MnGroupValuesClick(Sender: TObject);
var list : TStringList;
    FName : String;

begin
  try
    Jei;
    FName := IncludeTrailingBackSlash(DirectoryEdit1.Directory) + 'Werte_in_Spalte_' + DBGridFaks.SelectedField.FieldName + '.txt';
    list := TStringList.Create;
    list.Duplicates:=dupIgnore;
    List.Sorted:=true;

    ProgressBar1.Position:=0;
    ProgressBar1.Max:=QFaks.RecordCount;
    BM := QFaks.GetBookmark;

    QFaks.DisableControls;

    QFaks.First;

    while not QFaks.EOF do
    begin
      list.Add(DBGridFaks.SelectedField.AsString);
      QFaks.Next;
      ProgressBar1.Position := QFaks.RecNo;
    end;


   if ((list.Count < 50) And (list.Count > 0)) then
   begin
      ShowMessage('Diese Werte gabs in ''' + DBGridFaks.SelectedField.FieldName + '''' + NL + NL +
      list.Text + NL + 'Strg + c kopiert das!');

   end
   else
   begin
     list.SaveToFile(FName);

     OpenURL(FName);

   end;

  finally
    Nei;
    FreeAndNil(List);
    if QFaks.BookmarkValid(BM) then
        QFaks.GotoBookmark(BM);

    QFaks.FreeBookmark(BM);

    QFaks.EnableControls;

    ProgressBar1.Position :=0;
    Application.ProcessMessages;


  end;

end;

procedure TForm1.MnJumpToNewValueClick(Sender: TObject);
var val : string;
begin
  try
  (* zum Ende der Werteliste gehen *)
   val := DBGridFaks.SelectedField.AsString;

   QFaks.DisableControls;

   (* damit nur eine Zeile markiert aussieht *)
   DBGridFaks.Options:=DBGridFaks.Options -[dgMultiselect];

   while not QFaks.EOF do
   begin

     Application.ProcessMessages;

     if DBGridFaks.SelectedField.AsString <> val then break;

     QFaks.Next;

   end;

   Application.ProcessMessages;
   DBGridFaks.SetFocus;

   finally
     DBGridFaks.Options:=DBGridFaks.Options +[dgMultiselect];
     QFaks.EnableControls;
   end;
end;

procedure TForm1.mnLoadFilterClick(Sender: TObject);
var OldFilters : TStringList;
    x : integer;
begin

  try
   Jei;
   (* vorhandene Filter sichern *)
   OldFilters := TStringList.Create;

   OpenDialog1.InitialDir:=ExePath;
   OpenDialog1.FilterIndex:=3;

   if OpenDialog1.Execute then
   begin
     OldFilters.Assign(FilterCombo.Items);
     FilterCombo.Items.LoadFromFile(OpenDialog1.FileName);

     (* Filter alt und geladen abgleichen, fehlende im geladenen nachtragen *)
     for x := 0 to OldFilters.Count -1 do
     begin
       if  FilterCombo.Items.IndexOf(OldFilters[x]) = -1 then
           FilterCombo.Items.Add(OldFilters[x]);

     end;

   end;

   QFaks.Filter:=FilterCombo.Items[0];
   QFaks.Filtered:=true;
   FilterCombo.Text:=FilterCombo.Items[0];
   ShowFilterInfo(True);

   finally
     Nei;
     FreeAndNil(OldFilters);

   end;

end;

procedure TForm1.MnManuelleBuchungenClick(Sender: TObject);
begin
  QFaks.Filter:='LENGTH(BELEGNR)>''6''';
  FilterCombo.Text:=QFaks.Filter;
  QFaks.Filtered:=true;
  ShowFilterInfo(true);
  DBGridFaks.SelectedField := QFaks.FieldByName('BELEGNR');

  ShowRecordCount(Sender);
end;

procedure TForm1.MnMarkExportedClick(Sender: TObject);
var
  FName, mark: string;
  MyDirSelect: TSelectDirectoryDialog;
begin

  if Messagedlg(
    'Das ist sehr gefährlich!! Sie werden bereits verarbeitete Daten ändern! Wollen Sie wirklich fortfahren?',
    mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    exit;

  FName := SavePath;

  if InputQuery('Bereits in AmisData importierte Dbf-Datei',
    'Pfad und(!) Dateiname der dbf aus dem Ordner: ' + FName, FName) = False then
    exit
  else
  begin
    StatusBar1.SimpleText := 'bearbeite die Exportmarkierung für: ''' + FName + '''';
    Application.ProcessMessages;
    if InputQuery('Was soll in Spalte VertragsNr eingetragen werden?',
      '1 = wurde exportiert NULL = noch nicht exportiert', mark) = False then
      exit;

    //ShowMessage(mark);

    (* Dialog SelectDirectory erzeugen *)
    MyDirSelect := TSelectDirectoryDialog.Create(Application);

    if not DirectoryExists(SavePath) then
    begin

      (* einen halbwegs passenden Vorgabewert für Directory erzeugen,
         das tatsächliche Directory wird später in ini gespeichert *)
      MyDirSelect.Initialdir := DirectoryEdit1.Directory;

      MyDirSelect.Title :=
        'Wo stehen die bereits in Amisdata Schnittstelle BLE importierten Faks-Daten?';
      (* den Pfad der Dateien ermitteln *)
      if MyDirSelect.Execute then
        SavePath := IncludeTrailingBackslash(MyDirSelect.FileName)
      else
        exit;

      Application.ProcessMessages;
    end;

    FreeAndNil(MyDirSelect);

    MarkExported(FName, mark);
  end;
end;

procedure TForm1.MnOraFixErrorClick(Sender: TObject);
begin
  try
   jei;
    AnalyzeTable('f2fsv');
   finally
    Nei;
   end;
end;

procedure TForm1.mnSaveFilterClick(Sender: TObject);
begin
   SaveDialog1.InitialDir:=ExePath;
   SaveDialog1.FilterIndex:=3;

   SaveDialog1.FileName:='Faks_Meldung_Filter_.flt';

   if SaveDialog1.Execute then
     FilterCombo.Items.SaveToFile(SaveDialog1.FileName);

end;

procedure TForm1.MnSearchClick(Sender: TObject);
var Col, x, y : Integer;
    found : boolean;
begin
  (* richtiges TStringGrid aktivieren *)
  if Sender = BtnSearch1 then
  begin
    GridReplaceGattungsart.SetFocus;
    Application.ProcessMessages;
  end
  else if Sender = BtnSearch then
  begin
    GridReplaceLineNumber.SetFocus;
    Application.ProcessMessages;
  end;

  (* Im StringGrid einen Wert suchen *)
  (Screen.ActiveControl as TStringGrid).SetFocus;
  col := (Screen.ActiveControl as TStringGrid).col;
  y :=   (Screen.ActiveControl as TStringGrid).Row;

  gesucht := (Screen.ActiveControl as TStringGrid).Cells[col,(Screen.ActiveControl as TStringGrid).Row];

  gesucht := InputBox(
    'Welche Zeichenfolge soll in Spalte ''' +
    (Screen.ActiveControl as TStringGrid).Cells[col,0] +
    ''' gesucht werden?', 'Suchbegriff exakte Schreibweise:', gesucht);


  found := false;

  for x := 1 to (Screen.ActiveControl as TStringGrid).RowCount -1 do
  begin
     if (Screen.ActiveControl as TStringGrid).Cells[col,x] = gesucht then
     begin
       found := true;
       (Screen.ActiveControl as TStringGrid).Row:=x;
       (Screen.ActiveControl as TStringGrid).Col:=Col;
       (Screen.ActiveControl as TStringGrid).SetFocus;
       break;
     end;
  end;

  if not found then
  begin
    (Screen.ActiveControl as TStringGrid).Row:=y;
    (Screen.ActiveControl as TStringGrid).Col:=Col;
    (Screen.ActiveControl as TStringGrid).SetFocus;

     ShowMessage(gesucht + ' konnte in Spalte ''' + (Screen.ActiveControl as TStringGrid).Cells[col,0] + ''' nicht gefunden weden.' + NL +
     'Achtung: GROSS/klein ist kriegsentscheidend!');


  end;


end;

procedure TForm1.MnSetAST_TarifVersionClick(Sender: TObject);
var OldSQL : string;
begin
  try
   (* bei Anrufsammeltaxi die Tarifversion setzen *)
   Jei;
   (* alte SQL sichern *)
   OldSQL := ZUpDateRid.SQL.Text;
   ZUpDateRid.SQL.TEXT := 'UPDATE f2fsv SET TARIFVERSION=' + IntToStr(SpinEditTarifversion.Value) +
   ' WHERE (Bemerkung2=''AST'' OR MDEID=''Notfahrkarten'') AND DATUM>=' + QuotedStr('01.01.' + IntToStr(YearOf(DateEditVon.Date)));
   ZUpDateRid.ExecSQL;
   ShowMessage('Anzahl betroffener Datensätze: ' + IntToStr(ZUpDateRid.RowsAffected));
   EventLog1.Log('AST Tarifversion wurde neu gesetzt mit: ' + ZUpDateRid.SQL.TEXT);
  finally
    Nei;
    (* SQL rücksichern *)
    ZUpDateRid.SQL.Text := OldSQL;
    QFaks.Refresh;

  end
end;

procedure TForm1.MnShowRecordCountClick(Sender: TObject);
begin
   ShowMessage('Angezeigt werden ' + FormatFloat('#,##0',QFaks.RecordCount) + ' Datensätze.');
end;

procedure TForm1.MnSortByClick(Sender: TObject);
begin
  try
   Jei;
   (* Sortierung *)
   QFaks.SortedFields:='GATTUNGSART, PreisStDruck, Betrag';
   QFaks.First;
   DBGridFaks.SelectedField := QFaks.FieldByName('Gattungsart');
  finally
    Nei;
  end;
end;

procedure TForm1.MnSucheLinienClick(Sender: TObject);
begin
  ListBoxKnownLines.SetFocus;
  gesucht := ListBoxKnownLines.Items[ListBoxKnownLines.ItemIndex];
  gesucht := InputBox(
    'Welche LinienNummer soll in der Liste gesucht werden?', 'Suchbegriff exakte Schreibweise:', gesucht);

  if ListBoxKnownLines.Items.IndexOf(gesucht) = -1 then
   ShowMessage(gesucht + ' konnte nicht gefunden werden!')
  else
   ListBoxKnownLines.Selected[ListBoxKnownLines.Items.IndexOf(gesucht)] := true ;

end;

procedure TForm1.MnSumColumnClick(Sender: TObject);
var
  x: integer;
  s: Float;
  JFieldName: string;
begin
  try
    Jei;
    QFaks.DisableControls;
    BM := QFaks.GetBookmark;
    StoppIt := False;

    ProgressBar1.Position := 0;
    ProgressBar1.Max := QFaks.RecordCount;

    JFieldName := DBGridFaks.SelectedField.FieldName;

    QFaks.First;
    s := 0;

    while not QFaks.EOF do
    begin
      s := s + QFaks.FieldByName(JFieldName).AsFloat;
      QFaks.Next;
      ProgressBar1.Position := QFaks.RecNo;

      Application.ProcessMessages;
      if Stoppit then
      begin
        ShowMessage('Aktion durch ESC-Taste abgebrochen!');
        StoppIt:= false;
        exit;
      end;

    end;

    ShowMessage('Summe in Spalte ''' + JFieldName + '''  ist: ' +
      FormatFloat('#,##0.00', s));
  finally
    Nei;
    if QFaks.BookmarkValid(BM) then
      QFaks.GotoBookmark(BM);
    QFaks.FreeBookmark(BM);
    QFaks.EnableControls;
    ProgressBar1.Position := 0;

  end;

end;

procedure TForm1.OpenLogFileClick(Sender: TObject);
begin
  OpenLog(ZSQLMonitor1.FileName);
end;

procedure TForm1.Pm_SearchClick(Sender: TObject);
var
  gefunden: boolean;
begin

  //Raise Exception.Create ('Division by Zero would occur');
  //exit;

  if Sender <> Form1 then
  begin

    if DBGridFaks.SelectedField.DataType = ftFloat then
    begin
      gesucht := DBGridFaks.SelectedField.AsFloat;
    end
    else
    begin
      gesucht := DBGridFaks.SelectedField.AsString;
    end;

    gesucht := InputBox(
      'Welche Zeichenfolge soll ab aktueller(!) Position gesucht werden in Spalte ''' +
      DBGridFaks.SelectedField.FieldName +
      '''?', 'Suchbegriff ist (F3=Suche fortsetzen!):', gesucht);

    StatusBar1.SimpleText := 'gesucht wird nach Textbestandteil ''' +
      gesucht + ''' in Spalte ''' + DBGridFaks.SelectedField.FieldName + '''';
  end;

  try
    Jei;
    gefunden := False;
    BM := QFaks.GetBookmark;

    if not QFaks.EOF then
      QFaks.Next;

    while not QFaks.EOF do
    begin
      (* auch Teilzeichenfolge caseinsensitive finden  *)
      //if not  AnsiContainsText(DBGridFaks.SelectedField.AsString ,gesucht) then  QFaks.Next

      (* in Abhängigkeit vom FeldTyp suchen *)
      if DBGridFaks.SelectedField.DataType = ftFloat then
      begin
        if not (DBGridFaks.SelectedField.AsFloat = gesucht) then
          QFaks.Next
        else
        begin
          gefunden := True;
          break;
        end;
      end
      else
      begin
            (* vergleich caseinsensitive *)
        if not AnsiContainsText(DBGridFaks.SelectedField.AsString, gesucht) then
          QFaks.Next
        else
        begin
          gefunden := True;
          break;
        end;
      end;
    end;

    Nei;

    if not gefunden then
    begin
      QFaks.GotoBookmark(BM);
      ShowMessage('''' + VarToStr(gesucht) + ''' konnte in Spalte ''' +
        DBGridFaks.SelectedField.Fieldname + ''' nicht gefunden werden!');
      StatusBar1.SimpleText := '';
    end;

  finally
    QFaks.FreeBookmark(BM);

  end;
end;

procedure TForm1.PopupGridClose(Sender: TObject);
begin
      Application.ProcessMessages;
end;

procedure TForm1.PopupGridPopup(Sender: TObject);
begin
  (* ob das gegen ZugriffsFehler hilft? *)
  Application.ProcessMessages;
  DBGridFaks.SetFocus;
  DBGridFaks.SelectedField := QFaks.DataSetField;
  Application.ProcessMessages;
end;

procedure TForm1.RemoveFilterClick(Sender: TObject);
begin
  QFaks.Filtered := False;
  QFaks.Filter := '';
  FilterCombo.Text := QFaks.Filter;
  StatusBar1.SimpleText := '';

  ShowFilterInfo(False);

  ShowRecordCount(Sender);

end;

procedure TForm1.RID_as_FilterClick(Sender: TObject);
var
  Filter, StrOr: string;
  x, MaxRecs: integer;
  tab : TDBF;
begin
  (* ab wann sollen die RID's in Datei geschrieben werden? *)
  MaxRecs := 100;

  StrOr := ' OR RID=';
  try
    jei;
    tab := TDBF.Create(Application.MainForm);

    QFaks.DisableControls;
    BM := QFaks.GetBookmark;

    if not QFaks.Filtered then
      if Messagedlg(
        'Die Daten sind ungefiltert, der Filterausdruck für die RID kann riesig werden!' + NL +
        'Bei mehr als ' + IntToStr(MaxRecs) + ' Datensätzen werden die RID''s in eine DBF-Datei geschrieben!' +
        NL + NL + 'Wollen Sie wirklich fortfahren?', mtConfirmation, [mbYes, mbNo], 0) =
        mrNo then
        exit;



    QFaks.First;

    if QFaks.RecordCount>=MaxRecs then
    begin
      ProgressBar1.Position:=0;
      ProgressBar1.Max:=QFaks.RecordCount;
      tab.TableName:=IncludeTrailingBackslash(DirectoryEdit1.Directory) + 'RID_Kontrolle.dbf';
      tab.FieldDefs.Add('RID',ftString,30);
      tab.CreateTable;
      tab.Open;
      while not QFaks.EOF do
      begin
        tab.Insert;
        tab.Edit;
        tab.FieldByName('RID').AsString:=QFaks.FieldByName('RID').AsString;
        tab.Post;
        QFaks.Next;
        ProgressBar1.Position:=QFaks.RecNo;

      end;


      if Messagedlg('Explorer öffnen? ' + IntToStr(Tab.Recordcount) + ' Datensätze wurden geschrieben nach ''' + Tab.TableName + '''',mtConfirmation,[mbYes,mbNo],0)= mrYes then
      OpenExplorer(IncludeTrailingBackslash(DirectoryEdit1.Directory) + Tab.TableName);

      tab.close;

    end
    else
    begin
      Filter := 'RID=' + QuotedStr(QFaks.FieldByName('RID').AsString);
      while not QFaks.EOF do
      begin
        QFaks.Next;
        Filter := Filter + StrOr + QuotedStr(QFaks.FieldByName('RID').AsString);
      end;

      mem.Clear;
      mem.Text := Filter;
      mem.SelectAll;
      mem.CopyToClipboard;

      QFaks.GotoBookmark(BM);

      ShowMessage(
        'Der erstellte RID-Filter wurde in die Zwischenablage kopiert und könnte anderswo per Strg + V eingefügt werden.');
    end;
  finally
    Nei;
    QFaks.EnableControls;
    QFaks.FreeBookmark(BM);
    FreeAndNil(tab);
    ProgressBar1.Position:=0;

  end;

end;

procedure TForm1.SaveToFileClick(Sender: TObject);
begin
  SaveDialog1.InitialDir := ExePath;
  if SaveDialog1.Execute then
  begin
    Memo1.Lines.SaveToFile(SaveDialog1.FileName);
  end;
end;

procedure TForm1.SpinEdit1Change(Sender: TObject);
var
  D, M, Y: word;

begin
  try

    VonDatum := incMonth(DateEditVon.Date, SpinEdit1.Value);
    DecodeDate(VonDatum, Y, M, D);
    BisDatum := EndOfAMonth(Y, M);

    StatusBar1.SimpleText := 'VonDatum: ' + FormatDateTime('dd.mm.yyyy', VonDatum) +
      ' ' + 'Bisdatum: ' + FormatDateTime('dd.mm.yyyy', BisDatum);


  finally

  end;
end;

procedure TForm1.SpinEditTarifversionEditingDone(Sender: TObject);
begin
  Memo1.SetFocus;
  ShowMessage('Bitte neu starten!');
end;

procedure TForm1.SQLLoadClick(Sender: TObject);
begin
  OpenDialog1.InitialDir := ExePath;
  if OpenDialog1.Execute then
  begin
    Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
     if QFaks.Filtered then
     ShowMessage('Vorsicht, ein Datenfilter ist noch in Benutzung!!' + NL + NL +
     QFaks.Filter);
  end;
end;

procedure TForm1.UniqueInstance1OtherInstance(Sender: TObject;
  ParamCount: integer; Parameters: array of string);
begin
  (* Anwendung wurde doppelt gestartet: *)
  ShowMessage('Das Programm läuft bereits!' + NL + NL +
    QuotedStr(ExtractFileName(Application.ExeName)) + NL + NL +
  'Die Anwendung kann nicht mehrfach gestartet werden.! ');

  if WindowState = wsMinimized then
    Application.Restore
  else
  begin
    BringToFront;
    SetFocus;
  end;
end;

procedure TForm1.QFaksAfterOpen(DataSet: TDataSet);
var
  x: integer;
begin
  Memo1.Lines.Assign(QFaks.SQL);
  //DBGridFaks.AutoSizeColumns;
  DBGridFaks.AutoAdjustColumns;

  (* dafür sorgen, dass nicht 1,0999999999 statt 1,10 EURO angezeigt werden *)
  for x := 0 to QFaks.FieldCount - 1 do
  begin
    if QFaks.Fields[x].DataType = ftFloat then
    begin
      (QFaks.Fields[x] as TFloatField).Precision := 15;
      (QFaks.Fields[x] as TFloatField).DisplayFormat := '#,##0.00';
    end;

    (*  Feld Zeit und Buchungszeit mit speziellemm Format *)
    if ((QFaks.Fields[x].FieldName = 'ZEIT') OR (QFaks.Fields[x].FieldName = 'BUCHUNGSZEIT')) then
    begin
       (QFaks.Fields[x] AS TDateTimeField).DisplayFormat := 'hh:mm:ss';
    end;

  end;

  (* Hinweis CheckLinien auszuführen *)
  x := trunc(Date - StrToDateDef(IniPropStorage1.StoredValue['CheckLinie'], 0));
  if x >= 30 then
   ShowMessage('Sie haben Check-Linien seit ' + IntToStr(x) + ' Tagen nicht mehr ausgeführt' + NL +
   'Auch die TarifVersion für Anrufsammeltaxi sollte aktualisiert werden (Rechtsklick: Mehr/Setze bei AST die Tarifversion)'
   + NL +
   'Ferner Rechtsklick: Mehr/Anrufsammeltaxi HTK und MTK ausführen!' + NL +
   'Bitte nachholen');

   ShowRecordCount(self);

   PageControl1.ActivePage := TabDaten;
   DBGridFaks.SetFocus;
end;

procedure TForm1.ApplicationProperties1Hint(Sender: TObject);
begin
  StatusBar1.SimpleText := Application.Hint;
end;

procedure TForm1.AuswahlFilterClick(Sender: TObject);
var
  Filter: string;
  row: integer;
begin
  //if QFaks.Filtered then QFaks.Filtered := False;


  (* gibts MultiSelect im Grid oder ganzes Grid kopieren *)
  if DBGridFaks.SelectedRows.Count > 1 then
  begin
    (* also Multiselect! *)
    statusbar1.SimpleText := 'Das DBGrid ''' + DBGridFaks.Name +
      ''' hat ' + IntToStr(DBGridFaks.SelectedRows.Count) + ' Zeilen selektiert!';


    // die RID speichern
    for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
    begin

      QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));

      if row = 0 then
        Filter := 'RID=''' + QFaks.FieldByName('RID').AsString + ''''
      else
        Filter := Filter + ' OR RID= ''' + QFaks.FieldByName('RID').AsString + '''';

    end;

  end
  else
  begin


    (* Filter auf Felder nur mit Zeit kann ich nicht *)
    //if not CheckFilterPossible(DBGridFaks.SelectedField.FieldName) then exit;




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

  (* nur zur Liste hinzufügen, falls noch nicht darin enthalten *)
  AddFilterHistory(Filter);



  (* rotes Panel als Warnhinweis *)
  //FilterCombo.Visible := True;
  ShowFilterInfo(True);

  ShowRecordCount(Sender);

  Application.ProcessMessages;


end;

procedure TForm1.BtnAddLine1Click(Sender: TObject);
begin
  (* Zeile hinzufügen *)
  GridReplaceGattungsart.RowCount:= GridReplaceGattungsart.RowCount +1 ;
  (* neue Zeile auswählen *)
  GridReplaceGattungsart.Row := GridReplaceGattungsart.RowCount -1;
  GridReplaceGattungsart.SetFocus ;

end;

procedure TForm1.BtnAddLineClick(Sender: TObject);
begin
  (* Zeile hinzufügen *)
  GridReplaceLineNumber.RowCount:= GridReplaceLineNumber.RowCount +1 ;
  (* neue Zeile auswählen *)
  GridReplaceLineNumber.Row := GridReplaceLineNumber.RowCount -1;
  GridReplaceLineNumber.SetFocus ;
end;

procedure TForm1.BtnDelLine1Click(Sender: TObject);
begin
  GridReplaceGattungsart.SetFocus;
  if Messagedlg('Soll die aktuelle Zeile: ' + GridReplaceGattungsart.Cells[GridReplaceGattungsart.Col, GridReplaceGattungsart.Row]  + ' wirklich gelöscht werden?',mtConfirmation,[mbYes,mbNo],0)= mrYes then
  GridReplaceGattungsart.DeleteColRow(false,GridReplaceGattungsart.Row);

end;

procedure TForm1.BtnDelLineClick(Sender: TObject);
begin
  GridReplaceLineNumber.SetFocus;
  if Messagedlg('Soll die aktuelle Zeile: ' + GridReplaceLineNumber.Cells[GridReplaceLineNumber.Col, GridReplaceLineNumber.Row]  + ' wirklich gelöscht werden?',mtConfirmation,[mbYes,mbNo],0)= mrYes then
  GridReplaceLineNumber.DeleteColRow(false,GridReplaceLineNumber.Row);
end;

procedure TForm1.BtnSearchClick(Sender: TObject);
begin
  MnSearchClick(Sender);
end;

procedure TForm1.ApplicationProperties1Exception(Sender: TObject; E: Exception);
begin

  if (pos('01722',E.Message) > 0) then
  begin
     ShowMessage('Oracle Fehler 01722, wir versuchens mal mit analyze table, vielleicht hilfts ja?!');
     AnalyzeTable('f2fsv');
  end
  else
  begin
  ShowMessage('Mist ein Fehler:' + NL + NL + E.Message + NL + (* (Sender as TComponent).Name + *) NL +
    'Faks Spalte RID hat den Wert ' + RowID + NL + NL +
    'Oft hilft auch ein Neustart der Anwendung!!');
  (* Oracle LogFile anzeigen *)
  //OpenLogFileClick(Sender);
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  line: string;
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
    line := StringsToStr(Memo1.Lines, '°', True);
    if (SQLHistory.IndexOf(Line) = -1) then
    begin
      SQLHistory.Add(line);
      SQLHistoryIndex := SQLHistory.Count -1;
    end;

    //Memo1.Hint:='Die SQL-History enthält ' + IntToStr(SQLHistory.Count) + ' Einträge. Drücken Sie Strg + F8, um durch die Liste zu schalten.' + NL +
    //'Mit ''' + Button2.Caption + ''' führen Sie den angezeigten SQL-Code aus.';


  finally
    Nei;

  end;

end;

procedure TForm1.CbincMonthClick(Sender: TObject);
begin
  if CbincMonth.Checked then
  begin
    SpinEdit1Change(Sender);

    (* damit das ändern per Code geht: *)
    DateEditVon.DirectInput:=true;
    DateEditBis.DirectInput:=true;

    DateEditVon.Date := VonDatum;
    DateEditBis.Date := BisDatum;

    (* zurücknehmen: *)
    DateEditVon.DirectInput:=false;
    DateEditBis.DirectInput:=false;


    Verbinden();
    CbincMonth.Checked := False;
  end;
end;

procedure TForm1.cbSQLChange(Sender: TObject);
begin
   (* Hinweis, wenn tatsächlich CheckBox Wert geändert wurde und nicht nur ini eingelesen *)
   if ((PageControl1.ActivePage = TabConfig) and (pos('Datensätze',Form1.lbRecordCount.Caption) > 0)) then
   ShowMessage('Bitte die Anwendung neu starten, der SQL-Code muß neu erzeugt werden!');
end;

procedure TForm1.CheckLinieClick(Sender: TObject);
var
  x, y: integer;
  lst: TStringList;
begin

  if not QFaks.Active then
    exit;
  (* Alle Liniennummern sammeln und anzeigen *)
  try
    Jei;
    Application.ProcessMessages;

    lst := TStringList.Create;
    lst.Sorted := True;
    lst.Duplicates := dupIgnore;

    BM := QFaks.GetBookmark;

    ProgressBar1.Position:=0;
    ProgressBar1.Max:=QFaks.RecordCount;


    QFaks.First;

    while not QFaks.EOF do
    begin
      lst.Add(QFaks.FieldByName('LINIE').AsString);
      QFaks.Next;
      ProgressBar1.Position:=QFaks.RecNo;
    end;

    (* nur unbekannte Linien zu ListBoxKnownLines hinzufügen *)
    for x := lst.Count - 1 downto 0 do
    begin
      if ListBoxKnownLines.Items.IndexOf(lst[x]) > -1 then
        lst.Delete(x);
    end;



    Mem.Clear;
    Mem.Text := lst.Text;
    Mem.SelectAll;
    Mem.CopyToClipboard;
    QFaks.GotoBookmark(BM);

    ProgressBar1.Position:=0;
    Application.ProcessMessages;


    if lst.Count > 0 then
    begin
      if Messagedlg('Sollen diese neuen Linien' + NL + lst.Text + NL +
        'Auf der Seite ''Einstellungen'' der Liste bekannter Linien hinzugefügt werden?'
        ,
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        ListBoxKnownLines.Items.AddStrings(lst);
        PageControl1.ActivePage := TabConfig;
        ListBoxKnownLines.SetFocus;
      end;
    end
    else
      ShowMessage('Alle Linien sind bereits bekannt. Vorsicht Linie 901 und so''n Mist vor Import AmisData löschen!'
        + NL + lst.Text);

    (* prüfe auf Mandant-Linie Duplikate *)
    if Messagedlg('In AmisData darfs ja keine doppelten LinienNummern geben, soll Mandant/Linie als Liste angezeigt werden?',mtConfirmation,[mbYes,mbNo],0)= mrYes then
     LinieJeMandant(Sender);

    (* Letzte AusfÜhrung von Check-Linien in Ini speichern.
       Wird bei QFaks.AfterOpen geprüft *)
    IniPropStorage1.StoredValue['CheckLinie'] := DateTimeToStr(Date);
  finally
    Nei;
    QFaks.FreeBookmark(BM);
    FreeAndNil(lst);
  end;
end;

procedure TForm1.DateEditVonAcceptDate(Sender: TObject; var ADate: TDateTime;
  var AcceptDate: boolean);
begin
  try
    //ShowMessage(FormatDateTime('dd.mm.yyyy',ADate));
    if Yearof(Date) <> YearOf(ADate) then
       ShowMessage('Achtung: Sie untersuchen nicht mehr das aktuelle Jahr!' + NL + NL +
       'Bedenken Sie die Auswirkung der Tarifversion auf Seite ''Einstellungen'' im SQL-Ausdruck!!!');
    jei;
    AcceptDate := True;
    (* Filter entfernen -> Daten neu einlesen *)
    RemoveFilterClick(Sender);
    StatusBar1.SimpleText :=
      'Filter wurde entfernt ...';
    Application.ProcessMessages;
  finally
    Nei;
    Application.ProcessMessages;
  end;
end;

procedure TForm1.DBase_exportClick(Sender: TObject);
var
  z, NichtGefunden, x, recs: integer;
  FName, StrLinie, Gattung, NeueLinienNr: string;
  gefunden, Amis: boolean;
  ConvertErrors: TStringList;
  Einnahmen: currency;
  v: variant;
begin
  try
    try


      if Messagedlg(
        'Sollen die Datensätze als nach AmisData exportiert gekennzeichnet werden?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        Amis := True
      else
        Amis := False;

      Application.ProcessMessages;


      (* Tabelle nach Spalte RID aufsteigend sortieren *)
      QFaks.SortedFields := 'RID';
      (* stAscending ist defibniert in ZAbstractRODataset *)
      QFaks.SortType := stAscending;

      Einnahmen := 0;
      Screen.ActiveControl.Invalidate;
      Application.MainForm.Invalidate;
      Application.ProcessMessages;


      (* Fortschrittsbalken auf 0 stellen *)
      ProgressBar1.Position := 0;

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
      if ((DateEditBis.Date - DateEditVon.Date <= 31) and
        (MonthOf(DateEditVon.Date) = MonthOf(DateEditBis.Date))) then
      begin
        FName := 'FaksDaten_' + FormatDateTime('mmyy', DateEditVon.Date);
      end
      else
      begin
        FName := 'FaksDaten_' + FormatDateTime('dd.mm.yy', DateEditVon.Date) + '-' +
          FormatDateTime('dd.mm.yy', DateEditBis.Date);
      end;

      Dbf1.TableName := IncludeTrailingBackslash(DirectoryEdit1.Directory) +
        FName + '.dbf';

      (* Dbf1 hat die FieldDefs der Elgeba *.dbf aus Tabelle gespeichert, plus ein neues Feld RID *)
      Dbf1.CreateTableEx(Dbf1.DBFFieldDefs);

      Jei;

      DBGridFaks.SetFocus;

      (* um das höchste Buchungsdatum und Buchungszeit im ersten Datensatz zu haben *)
      QFaks.SortedFields := 'Buchungsdatum, BuchungsZeit';
      QFaks.SortType := stDescending;
      QFaks.Refresh;

      QFaks.First;

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
          Dbf1.FieldByName('VDTNR').AsString :=
            '999' + QFaks.FieldByName('MDEIDINTERN').AsString;
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
          //if (QFaks.FieldByName('ID_F2MANDANT').AsInteger = 1) then
          //begin

          (* Keine Ahnung ob das besser ist *)
          //if ((QFaks.FieldByName('HSTSTARTIDENT').isNull) and (QFaks.FieldByName('VertriebsHSTIdent').AsInteger > 0)) then
          //if (QFaks.FieldByName('HSTSTARTIDENT').isNull) then
          //  Dbf1.FieldByName('HALTNR').AsString := QFaks.FieldByName('VertriebsHSTIdent').AsString
          //else
            Dbf1.FieldByName('HALTNR').AsString := QFaks.FieldByName('HSTSTARTIDENT').AsString;

          (* sicherstellen, dass die HALTNR maximal 5 Zeichen hat *)
          if length(Dbf1.FieldByName('HALTNR').AsString) > StrToInt(Edit_Max_HSTIDENT.Text) then
             Dbf1.FieldByName('HALTNR').AsString := copy(Dbf1.FieldByName('HALTNR').AsString,1,StrToInt(Edit_Max_HSTIDENT.Text));

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
          if QFaks.FieldByName('PV').AsString = 'RMV' then
          begin
            if QFaks.FieldByName('PREISSTIDENT').AsString = '205021' then
              Dbf1.FieldByName('PSTUFE').AsString := '25'
            else if QFaks.FieldByName('PREISSTIDENT').AsString = '205018' then
              Dbf1.FieldByName('PSTUFE').AsString := '28'
            else
              (* wird von AmisData nicht akzeptiert: PREISSTIDENT *)
              Dbf1.FieldByName('PSTUFE').AsString :=
                QFaks.FieldByName('PREISSTdruck').AsString;
          end;


          Dbf1.FieldByName('PREIS').AsCurrency :=
            QFaks.FieldByName('Betrag').AsCurrency;

          (* für das Eventlog1 *)
          Einnahmen := Einnahmen + QFaks.FieldByName('Betrag').AsCurrency;

          (* BuchungsDatum und -Zeit *)
          Dbf1.FieldByName('ZDEDATUM').AsDateTime :=
            QFaks.FieldByName('Buchungsdatum').AsDateTime;

          Dbf1.FieldByName('ZDEZEIT').AsString :=
            QFaks.FieldByName('BuchungsZeit').AsString;


          Dbf1.FieldByName('PREIS2').AsCurrency := 0;

          (* um den Ursprung des Datensatzes ggf. zu identifizieren *)
          Dbf1.FieldByName('RID').AsString :=
            QFaks.FieldByName('RID').AsString;

          (* idiotische Linie 31/32 korrigieren *)
          StrLinie := ExtractNumbers(QFaks.FieldByName('Linie').AsString);

          (* zu ersetzende LinienNummer nachschlagen *)
          Dbf1.FieldByName('LINIE').AsString := LookUpStringGrid(GridReplaceLineNumber,QFaks.FieldByName('ID_F2MANDANT').AsString,StrLinie);

          (*
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
          *)


          (* Sortennummer *)
          //if ((QFaks.FieldByName('PV').AsString = 'RMV') and
          //  (QFaks.FieldByName('TarifVersion').AsInteger < TarifVersion)) then
          //begin
          //  //if QFaks.FieldByName('SORTENNUMMER').IsNull then
          //  Dbf1.FieldByName('SORTE').AsString :=
          //    FormatFloat('00', QFaks.FieldByName('GAIDENT').AsFloat) +
          //    FormatFloat('00', QFaks.FieldByName('PreisStDruck').AsFloat);
          //end
          //else
          //begin

            (* Sonderfall Anrufsammeltaxi OHNE Sortennummer *)
            if ((QFaks.FieldByName('MDEID').AsString='Notfahrkarten') and (QFaks.FieldByName('SORTENNUMMER').isNull)) then
            begin
               //if QFaks.FieldByName('GAIDENT').AsString = '29' then Dbf1.FieldByName('SORTE').AsString :='3300';
               //if QFaks.FieldByName('GAIDENT').AsString = '27' then Dbf1.FieldByName('SORTE').AsString :='3100';
               //if QFaks.FieldByName('GAIDENT').AsString = '3' then Dbf1.FieldByName('SORTE').AsString :='300';
               //if QFaks.FieldByName('GAIDENT').AsString = '18' then Dbf1.FieldByName('SORTE').AsString :='2000';
               //if QFaks.FieldByName('GAIDENT').AsString = '16' then Dbf1.FieldByName('SORTE').AsString :='1800';
               //if QFaks.FieldByName('GAIDENT').AsString = '54' then Dbf1.FieldByName('SORTE').AsString :='1201';

               (* zu ersetzende Gattungsart nachschlagen *)
               Dbf1.FieldByName('SORTE').AsString := LookUpStringGrid(GridReplaceGattungsart,QFaks.FieldByName('ID_F2MANDANT').AsString,QFaks.FieldByName('GAIDENT').AsString);
               (* zur Kontrolle in Log schreiben *)
               //EventLog1.Log('Bei RID ' + QFaks.FieldByName('RID').AsString  + ' wurde für Gattungsart ' + QFaks.FieldByName('GAIDENT').AsString + ' Sortennummer ' + Dbf1.FieldByName('SORTE').AsString + ' eingetragen');
            end
            else
            Dbf1.FieldByName('SORTE').AsString := QFaks.FieldByName('SORTENNUMMER').AsString;

            (* Sonderfall Hessenticket *)
            if ((QFaks.FieldByName('SORTENNUMMER').IsNull) and
              (QFaks.FieldByName('PreisStDruck').AsString = '50')) then
              Dbf1.FieldByName('SORTE').AsString := '1750';
        (*
        Copy(QFaks.FieldByName('SORTENNUMMER').AsString,1,2) +
        FormatFloat('00', QFaks.FieldByName('PreisStDruck').AsFloat);
        *)

         // end; (* Sortennummer *)


          //end;


        except
          on E: EConvertError do
          begin
            ShowMessage('Es ist ein Fehler aufgetreten bei ROWID. ' +
              RowId + NL + NL + E.Message + NL + NL +
              'Fortsetzung erfolgt mit nächstem Datensatz!!');
            QFaks.Next;
            ConvertErrors.Add('ROWID=' + RowId);
            continue;
          end;
        end;




        Dbf1.FieldByName('STORNIERT').AsBoolean := False;

        if QFaks.FieldByName('PV').AsString = 'RMV' then
          Dbf1.FieldByName('WABE').AsString := '1';


        Dbf1.FieldByName('ZAHLART').AsString :=
          QFaks.FieldByName('ZAHLART').AsString;

        Dbf1.Post;

        (* Datensatzzähler erhöhen und in Prgressbar anzeigen *)
        Inc(recs);
        ProgressBar1.Position := recs;
        ProgressBar1.Invalidate;

        (* Datensatz als exportiert kennzeichen *)
        if Amis then
        begin
          v := Trim(QFaks.FieldByName('RID').AsString);
          ZUpDateRid.ParamByName('RID').AsString := v;
          ZUpDateRid.ParamByName('Vertragsnr').AsInteger := 1;
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
          NL + NL + FName + NL + 'Die Haltestellennummer wurde ggf. auf ' + Edit_Max_HSTIDENT.Text + ' Zeichen gekürzt!' + NL + NL +
          'Soll die Datei im Dateimanager Explorer angezeigt werden?',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          OpenExplorer(FName);
        end;

        (* Eventlog schreiben *)
        EventLog1.Log(FormatFloat('#,##0', z) + ' Datensätze und ' +
          FormatCurr('#,##0.00 (Summe der Spalte PREIS)', Einnahmen) +
          ' wurden geschrieben in ' + FName);
        EventLog1.Log('SQL-Code der Abfrage war: ' + StringsToStr(Memo1.Lines, ' ', True));
      end;
    finally
      nei;
      QFaks.EnableControls;

      if ConvertErrors.Count > 0 then
      begin
        ConvertErrors.Insert(0,
          'Bei diesen Datensaetzen gabs Konvertierungsfehler, sie sind wahrscheinlich unvollstaendig in der DBase Datei: '
          + FName);
        ConvertErrors.SaveToFile(ChangeFileExt(FName, '_Error.txt'));
        Shellexecute(Application.MainForm.Handle, 'open', PChar(
          ChangeFileExt(FName, '_Error.txt')), '', '', SW_NORMAL);

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

      ShowMessage('Bei dem FAKS-Datensatz mit der  RID: ' + NL +
        RowID + NL + ' ist ein Fehler aufgetreten. Die RID wurde kopiert!' +
        NL + 'Die Fehlermeldung lautet:' + NL + E.Message);

    end;
  end;

  QFaks.Refresh;

  (* Progressbar wieder auf 0 setzen *)
  ProgressBar1.Position := 0;
  ProgressBar1.Invalidate;
end;

procedure TForm1.DBGridFaksDblClick(Sender: TObject);
begin
  (* nur zu Fehlerdokumentation  *)
  ShowMessage('AsString: ' + QFaks.FieldByName('Betrag').AsString + NL +
    //'Value: ' +QFaks.FieldByName('Betrag').Value + NL +
    'AsFloat: ' + FormatFloat('#.000000', QFaks.FieldByName(
    'Betrag').AsFloat));
end;

procedure TForm1.DBGridFaksPrepareCanvas(sender: TObject; DataCol: Integer;
  Column: TColumn; AState: TGridDrawState);
begin
    (* siehe:  http://forum.lazarus.freepascal.org/index.php/topic,38357.msg260169.html#msg260169
      [gdSelected, gdFocused] * AState -> gesucht ist die Schnittmenge ausgedrückt durch *
    *)
    if (([gdSelected, gdFocused] * AState <> []) and (DBGridFaks.SelectedColumn = Column)) then
  begin
    DBGridFaks.Canvas.Brush.Color := clRed;
    DBGridFaks.Canvas.Font.Color := clWhite;
  end;
end;

(* Anzeige von up und down Arrows laut:
http://wiki.freepascal.org/Grids_Reference_Page#Sorting_columns_or_rows_in_DBGrid_with_sort_arrows_in_column_header *)
procedure TForm1.DBGridFaksTitleClick(Column: TColumn);
const
  ImageArrowUp = 0; //should match image in imagelist
  ImageArrowDown = 1; //should match image in imagelist

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
      QFaks.SortType := stAscending;
      Column.Title.ImageIndex := ImageArrowUp;

    end
    else
    begin
      QFaks.SortedFields := Column.Title.Caption;
      QFaks.SortType := stDescending;
      Column.Title.ImageIndex := ImageArrowDown;
    end;

    QFaks.First;
   // QFaks.GotoBookmark(BM);

  finally
    QFaks.FreeBookmark(BM);

    QFaks.EnableControls;

    // Remove the sort arrow from the previous column we sorted
    if (FLastColumn <> nil) and (FlastColumn <> Column) then
      FLastColumn.Title.ImageIndex := -1;

    FLastColumn := column;

    Nei;
  end;
end;

procedure TForm1.DeleteSelectedClick(Sender: TObject);
var
  x: integer;
begin
  x := ListBoxKnownLines.ItemIndex;
  if x > -1 then
    ListBoxKnownLines.DeleteSelected;
end;

procedure TForm1.DirectoryEdit1ButtonClick(Sender: TObject);
begin
  DirectoryEdit1.RootDir := ExePath;
end;

procedure TForm1.EventLog_anzeigenClick(Sender: TObject);
begin
  (* Eventlog anzeigen *)
  if FileExists(EventLog1.FileName) then
    ShellExecute(Application.MainForm.Handle, 'open', PChar(EventLog1.FileName),
      '', '', SW_Normal)
  else
    ShowMessage(EventLog1.FileName + NL + NL + 'wurde nicht gefunden!');
end;

procedure TForm1.ExportExcelClick(Sender: TObject);
begin
  try
    jei;
    Application.ProcessMessages;
    ExportDatasetToExcel(QFaks);
  finally
    nei;

  end;
end;

procedure TForm1.FilterAbClick(Sender: TObject);
var
  Filter: string;
begin
  (* Filter auf Felder nur mit Zeit kann ich nicht *)
  if not CheckFilterPossible(DBGridFaks.SelectedField.FieldName) then exit;

  (* ggf. vorhandenen Filter durch ' AND ' ergänzen *)
  if Trim(QFaks.Filter) <> '' then
    Filter := Trim(QFaks.Filter) + ' AND ';

  (* den Filter zusammensetzen *)
  Filter := Filter + DBGridFaks.SelectedField.FieldName + '>=' +
    QuotedStr(QFaks.Fields[DBGridFaks.SelectedField.Index].AsString);


  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  AddFilterHistory(Filter);


  ShowFilterInfo(True);

  (* zum ersten Datensatz springen *)
  QFaks.First;

  ShowRecordCount(Sender);

  Application.ProcessMessages;


end;

procedure TForm1.FilterBisClick(Sender: TObject);
var
  Filter: string;
begin
  (* Filter auf Felder nur mit Zeit kann ich nicht *)
  if not CheckFilterPossible(DBGridFaks.SelectedField.FieldName) then exit;

  (* ggf. vorhandenen Filter durch ' AND ' ergänzen *)
  if Trim(QFaks.Filter) <> '' then
    Filter := Trim(QFaks.Filter) + ' AND ';

  (* den Filter zusammensetzen *)
  Filter := Filter + DBGridFaks.SelectedField.FieldName + '<=' +
    QuotedStr(QFaks.Fields[DBGridFaks.SelectedField.Index].AsString);


  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  AddFilterHistory(Filter);

  ShowFilterInfo(True);

  (* zum letzten Datensatz springen *)
  QFaks.Last;

  ShowRecordCount(Sender);

  Application.ProcessMessages;


end;

procedure TForm1.FilterComboKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    DBGridFaks.SetFocus;

    if QFaks.Filtered then QFaks.Filtered := False;

    FilterCombo.Text := trim(FilterCombo.Text);

    QFaks.Filter := FilterCombo.Text;

    if QFaks.Filter > '' then
    begin
      QFaks.Filtered := True;

      (* zur Liste der Filter hinzufügen *)
      AddFilterHistory(QFaks.Filter);

      ShowFilterInfo(True);
    end
    else
    ShowFilterInfo(false);

    ShowRecordCount(Sender);

    Application.ProcessMessages;


  end;

end;

procedure TForm1.FilterSetzenClick(Sender: TObject);
var
  Filter: string;
begin
  if QFaks.Filtered then
    QFaks.Filtered := False;



  QFaks.Filter := Filter;
  QFaks.Filtered := True;

  AddFilterHistory(QFaks.Filter);

  ShowFilterInfo(true);


  if QFaks.RecordCount = 1 then
    PageControl1.ActivePage := TabDaten;
end;

procedure TForm1.FormActivate(Sender: TObject);
var
  warning: boolean;
  msg: string;
  j : integer;
begin
  (* das in FormCreate führt zu Access Violation!! *)
  ShowVersionInfo;

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
  //if trim(lbHostname.Text) = '' then
  //begin
  //  warning := True;
  //  msg := msg + ' Der Hostname existiert nicht.';
  //end;
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
    Memo1.SetFocus;
    ShowMessage('Bitte erst die Einstellungen durchführen und dann neu starten.Siehe die angezeigten Vorgabewerte (Stand 11/2016).' +
      NL + NL + msg);
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

  (* Checken, ob die Tarifversion richtig eingestellt ist: 2017 war es die 36 *)
  j := TarifVersion;
  (* korrekte Tarifversion berechnen: zu 36 die Differenz von Aktuellem Jahr - 2017 hinzuzählen  *)
  j := 36 + Yearof(Date) -2017;

  (* nur so als Test für git push *)

  if  ((j <> TarifVersion) and ShowAgain and (MonthOf(Date) > 1 )) then
  begin
     ShowAgain := false;
     PageControl1.ActivePage := TabConfig;
     SpinEditTarifversion.SetFocus;
     ShowMessage('Die Tarifverion: ' + IntToStr(TarifVersion) + ' scheint nicht zu stimmen.' + NL +
     'Im Jahr 2017 wars die TarifVersion 36, also Sollte es im Jahr ' + IntToStr(Yearof(Date)) + ' die ' +
     IntToStr(j) + ' sein.');
  end;

end;

procedure TForm1.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  QFaks.Close;
  //SQLTransaction2.Active := False;
  ZConnection1.Connected := False;

end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: boolean);
var
  x: integer;
begin

  (* leere Einträge aus FilterCombo löschen *)
  for x := FilterCombo.Items.Count - 1 downto 0 do
  begin
    if trim(FilterCombo.Items[x]) = '' then
      FilterCombo.Items.Delete(x);
  end;

  (* FilterCombo-Einträge ggf. löschen *)
  while FilterCombo.Items.Count >= 20 do
    FilterCombo.Items.Delete(FilterCombo.Items.Count -1);


  (* LinienErsetzungen für AmisData DBF-Export speichern *)
  GridReplaceLineNumber.SaveToCSVFile(ChangeFileExt(Application.ExeName,'_ErsetzungLinienNummern.csv'),';');

  (* GattungsartenErsetzungen für AmisData DBF-Export speichern *)
  GridReplaceGattungsart.SaveToCSVFile(ChangeFileExt(Application.ExeName,'_ErsetzungGattungsarten.csv'),';');




  FreeAndNil(SQLHistory);
  FreeAndNil(RIDS_in_Selection);

end;

procedure TForm1.JEi;
begin
  screen.cursor := crHourglass;
  Application.ProcessMessages;
end;

procedure TForm1.NEi;
begin
  screen.cursor := crDefault;
  Application.ProcessMessages;
end;





function TForm1.Verbinden(): boolean;
var
  d, m, y, d1, m1, y1: word;
  (* Vormonat und VorVormonat!!!! *)
  DatumVon, DatumBis, DatumVonVormonat, DatumBisVormonat: TDateTime;
  sql,  line, JBuchungsDatum, JBuchungsZeit, ExtraFilter : string;
  x: integer;
  F: TFloatField;
  dbf: TDbf;
  list: TStringList;
begin
  (* CAST(Betrag AS  NUMBER(8,2)) AS BETRAG *)

  (* Datum aus RID extrahieren: TO_DATE(SUBSTR(RID,1,8),'YYYYMMDD') AS RID_DATUM *)

  //ShowMessage('Verbinden!!');
  {$IFDEF WINDOWS}
  (* Datenbank öffnen *)
  try
    ZConnection1.Connected := False;


    //SQLTransaction2.Active := False;
    QFaks.Close;

    if ((trim(lbDatabase.Text) <> '') or
      (trim(lbUserName.Text) <> '') or (trim(lbPassword.Text) <> '')) then
    begin
      ZConnection1.Database := lbDatabase.Text;
      ZConnection1.User := lbUserName.Text;
      ZConnection1.Password := lbPassword.Text;
    end
    else
    begin
      PageControl1.ActivePage := TabConfig;
      lbDataBase.SetFocus;
      lbDataBase.SelectAll;
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


    ExtraFilter := '/* TO_CHAR(Zeit, ''HH24:MI:SS'') BETWEEN ''23:00:00'' AND ''23:30:00'' AND */';

    if not cbSQL.Checked then
    begin
    (* sehr tückisch ist die Klammersetzung für die Logik der SQL Verarbeitung!!!!!!! *)

    (* noch Ohne Tabelle F2Personal wegen Sammeltaxi, siehe unten *)
    sql :=
      'SELECT RID, VID, ID_F2MANDANT, DATUM,  ZEIT, DATUMFAHRT, Buchungsdatum, '
      + NL +
      '  BUCHUNGSZEIT, JOURNAL,MDEID, MDEIDINTERN, BELEGNR, Bemerkung, Bemerkung2, PNR, '
      + NL + 'LINIE, FKART, anzahl, Einzelpreis, ' +
      NL + '  BETRAG, GAIDENT, GATTUNGSART, PreisStDruck, PREISSTIDENT, ' +
      NL + 'Zahlart, Storno, DatumZeit, TarifVersion, NETZ, ORTStart, OrtZiel, PV, LfdNrPV, '
      + NL +
      'Storniert, Sortennummer, TZSTARTIDENT, TZZIELIDENT, TZVIAIDENT, HSTSTARTIDENT,HSTSTART, HSTZIELIDENT, HSTZIEL, VERTRIEBSHSTIDENT, Vertragsnr '
      + NL + 'FROM F2FSV  WHERE ' + ExtraFilter + ' ID_F2MANDANT <> ''3'' AND (PV=''RMV'') AND TarifVersion >='''
      +
      IntToStr(TarifVersion) + ''' AND ((DATUM BETWEEN ' +
      NL + '''' + DateTimeToStr(DateEditVon.Date) + ''' AND ' + NL +
      '''' + DateTimeToStr(DateEditBis.Date) + ''')' + NL +
      ' OR ( Vertragsnr is NULL AND DATUM <= ' + NL + '''' +
      DateTimeToStr(DateEditBis.Date) + ''' AND DATUM >=''01.01.' +
      IntToStr(Yearof(DateEditVon.Date)) + '''))';
    end
    else
    begin

    (* GEHT  jetzt auch bei dbf-Kontrolle über viel Monate
    mit left outer join auf F2Personal für Anrufsammeltaxi *)
    sql :=
      'SELECT a.RID, a.VID, a.ID_F2MANDANT, a.DATUM,  a.ZEIT, a.DATUMFAHRT, a.Buchungsdatum, '
      + NL +
      '  a.BUCHUNGSZEIT, a.JOURNAL, a.MDEID, a.MDEIDINTERN, a.BELEGNR, a.Bemerkung, a.Bemerkung2, a.PNR,b.Name, '
      + NL + ' a.LINIE, a.FKART, a.anzahl, a.Einzelpreis, ' +
      NL + '  a.BETRAG, a.GAIDENT, a.GATTUNGSART, a.PreisStDruck, a.PREISSTIDENT, ' +
      NL + ' a.Zahlart, a.Storno, a.DatumZeit, a.TarifVersion, NETZ, a.ORTStart, a.OrtZiel, a.PV, a.LfdNrPV, '
      + NL +
      ' a.Storniert, a.Sortennummer, a.TZSTARTIDENT, a.TZZIELIDENT, a.TZVIAIDENT, a.HSTSTARTIDENT, a.HSTSTART, a.HSTZIELIDENT, a.HSTZIEL, a.VERTRIEBSHSTIDENT, a.Vertragsnr '
      + NL +
      ' FROM F2FSV a ' +  NL + ' left outer join F2PERSONAL b ' + NL +
      ' on ( a.ID_F2PERSONAL=b.ID_F2PERSONAL  ' + NL +
      ' AND ' + NL +
      ' a.PNR=b.PNR  ' + NL +
      ' AND ' + NL +
      ' a.ID_F2MANDANT=b.ID_F2MANDANT ) ' + NL +
      ' WHERE ' + ExtraFilter + ' a.ID_F2MANDANT <> ''3'' AND (a.PV=''RMV'') AND a.TarifVersion >='''
      +
      IntToStr(TarifVersion) + ''' AND ((a.DATUM BETWEEN ' +
      NL + '''' + DateTimeToStr(DateEditVon.Date) + ''' AND ' + NL +
      '''' + DateTimeToStr(DateEditBis.Date) + ''')' + NL +
      ' OR ( a.Vertragsnr is NULL AND a.DATUM <= ' + NL + '''' +
      DateTimeToStr(DateEditBis.Date) + ''' AND a.DATUM >=''01.01.' +
      IntToStr(Yearof(DateEditVon.Date)) + '''))';
    end;(* cbSQL.Checked *)


    QFaks.SQL.Text := sql;

    //ShowMessage(QFaks.SQL.Text);

    //exit;

    //ShowMessage(sql);

    //exit;
    try
      ZConnection1.Connected := True;
      QFaks.Active := True;
      PageControl1.ActivePage := TabDaten;
    except
      on E: Exception do
      begin
       ShowMessage('Fehler:' + NL + E.Message );
      end;
    end;



    (* SQL in History speichern *)
    line := StringsToStr(Memo1.Lines, '°', True);
    if (SQLHistory.IndexOf(Line) = -1) then
    begin
      SQLHistory.Add(line);
      SQLHistoryIndex := SQLHistory.Count -1;
    end;



  finally
    nei;
    Result := True;
    //FreeAndNil(dbf);
    //FreeAndNil(List);
  end;
  {$ENDIF}
  ;

end;

procedure TForm1.BtnExportExcelClick(Sender: TObject);
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

procedure TForm1.SumClick(Sender: TObject);
var
  Preis, PreisDelta: currency;
  row: integer;
  VorVorMonat: TDateTime;
begin
  try


    if not QFaks.Active then
      exit;

    Screen.ActiveControl.Invalidate;

    Application.ProcessMessages;


     (* keine Ahnung wozu das mal nützlich sein sollte *)
    //(* Tabelle nach Spalte RID aufsteigend sortieren *)
    //QFaks.SortedFields := 'RID';
    //(* stAscending ist defibniert in ZAbstractRODataset *)
    //QFaks.SortType := stAscending;

    BM := QFaks.GetBookmark;
    QFaks.DisableControls;
    QFaks.First;
    Preis := 0;
    PreisDelta := 0;
    row := 0;
    Jei;


    if DBGridFaks.SelectedRows.Count > 1 then
    begin
      for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
      begin
        QFaks.GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));
        Preis := Preis + QFaks.FieldByName('Betrag').AsCurrency;
      end;
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

    if DBGridFaks.SelectedRows.Count > 1 then
      (* nur für gewählte Zellen *)
    begin
      ShowMessage('Anzahl der Datensätze: ' + IntToStr(row + 1) +
        NL + NL + 'Summe Spalte ''Betrag'': ' + FormatFloat('#,##0.00', Preis) +
        NL + NL + 'Davon Summe Feld Betrag, wenn noch nicht gemeldet (Feld VertragsNr leer ist): '
        + FormatFloat('#,##0.00', JNochZuMelden(Sender)));
    end
    else
    begin
      (*  für die ganze Tabelle *)
      ShowMessage('Anzahl der Datensätze: ' + IntToStr(QFaks.RecordCount) +
        NL + NL + 'Summe Spalte ''Betrag'': ' + FormatFloat('#,##0.00', Preis) +
        NL + NL + 'Davon Summe Feld Betrag, wenn noch nicht gemeldet (Feld VertragsNr leer ist): '
        + FormatFloat('#,##0.00', JNochZuMelden(Sender)));

      (* welcher Monat liegt zwei Monate zurück? *)
      VorVorMonat := incMonth(Date, -2);

      (* wird der richtige Monat betrachtet: ist er im Zeitraum oder kleiner?
         und liegen Beginn und Ende im selben Monat? *)
      if (((VorVorMonat >= DateEditVon.Date) and
        (VorVorMonat <= DateEditBis.Date)) or
        (VorVorMonat > DateEditVon.Date))
      (* Beginn und Ende im selben Monat *)
      (* and (FormatDateTime('mmyyyy',DateEditVon.Date) = FormatDateTime('mmyyyy',DateEditBis.Date)) *)
      then
      begin
        (* Summe der Einnahmen laut Preis in globaler Variablen speichern *)
        GesamtEinnahme := Preis;

        (* Kontroll-Vergleich mit bereits in Amisdata importierten Daten *)
        if Messagedlg('Soll der gerade ermittelte Betrag: ' + NL +
          FormatFloat('#,##0.00', Preis) + NL +
          'mit bereits in Amsidata für den Zeitraum ' +
          FormatDateTime('dd.mm.yyyy', DateEditVon.Date) + ' bis ' +
          FormatDateTime('dd.mm.yyyy', DateEditBis.Date) +
          ' gemeldeten Monatsmeldungen verglichen werden?' + NL + NL +
          'Die SEHR langwierige Aktion kann duch ESC-Taste abgebrochen werden!', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          BereitsGemeldetWurden(Sender);

      end;

    end;


  finally
    if QFaks.BookmarkValid(BM) then
      QFaks.GotoBookmark(BM);
    QFaks.EnableControls;

  end;
end;

procedure TForm1.QFaksBeforeOpen(DataSet: TDataSet);
begin
  if not ZSQLMonitor1.Active then
    ZSQLMonitor1.Active := True;
end;

(* wichtig: Damit das Ü in Grünberg erhalten bleibt: PWideChar(UTF8Decode(Zeile)) *)
function TForm1.ExportDatasetToExcel(JDataset: TDataset): boolean;
var
  x, row, recs, MaxRecs: integer;
  F: TextFile;
  Zeile: string;
  FName_csv, FName_xls, command: variant;
  V, VApp: olevariant;
  ExcelHandle: HWND;



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

  ProgressBar1.Position := 0;
  ProgressBar1.Max := JDataset.RecordCount;
  recs := 0;

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
      (* wichtig für deutsche Umlauts PWideChar(UTF8Decode(Zeile)) *)
      Zeile := PWideChar(UTF8Decode(Zeile));
      Writeln(F, Zeile);

      (* Mit und ohne Multiselect im Grid: *)
      if DBGridFaks.SelectedRows.Count > 1 then
      begin
        (* also mit Multiselect *)
        ProgressBar1.Position := 0;
        ProgressBar1.Max := DBGridFaks.SelectedRows.Count;
        recs := 0;

        Zeile := '';
        for row := 0 to DBGridFaks.SelectedRows.Count - 1 do
        begin
          GotoBookmark(TBookMark(DBGridFaks.SelectedRows[row]));

          for x := 0 to Fieldcount - 1 do
          begin
            if x = 0 then
            begin
              if Fields[x].isnull then
                Zeile := ''
              else
                Zeile := Fields[x].AsString;
            end
            else
            begin
              if Fields[x].isnull then
                Zeile := Zeile + ';' + ''
              else
                Zeile := Zeile + ';' + Fields[x].AsString;
            end;
          end;
          Zeile := PWideChar(UTF8Decode(Zeile));
          Writeln(F, Zeile);
          Inc(recs);
          ProgressBar1.Position := recs;
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
                Zeile := ''
              else
                Zeile := Fields[x].AsString;
            end
            else
            begin
              if Fields[x].isnull then
                Zeile := Zeile + ';' + ''
              else
                Zeile := Zeile + ';' + Fields[x].AsString;
            end;
          end;
          Zeile := PWideChar(UTF8Decode(Zeile));
          Writeln(F, Zeile);
          Inc(recs);
          ProgressBar1.Position := recs;
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

      if recs > 65500 then
        ShowMessage('Achtung, die Daten haben ' + IntToStr(recs) +
          ' Zeilen, evtl. wird in Excel nicht alles angezeigt.' + NL + NL +
          'Ggf. Rechtsklick, kopieren!!');

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

      (* maximale Anzahl Zeilen des sheets in Excel ermitteln *)
      MaxRecs := V.ActiveSheet.rows.Count;

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
        command := 'Achtung: Bei Wetterau Linie 56 in 1056 umbenennen!!!' +
          NL + 'Datenfilter ist: ' + QFaks.Filter;
        V.Selection.Characters.Text := command;
      end;

      V.Range['A2'].Select;


      (* jetzt als richtige Excel Datei speichern *)
      V.ActiveSheet.SaveAs(FName_xls, 1);
      (* die csv-Quelldatei kann weg *)
      if recs < MaxRecs then
        SysUtils.DeleteFile(ExePath + Name + '.csv')
      else
        ShowMessage('Siehe auch die Daten in' + NL + NL + ExePath +
          Name + '.csv' + NL + 'mit ' + IntToStr(recs) + ' Zeilen' + NL + NL +
          'Das aktuelle Excel akzeptiert maximal ' + IntToStr(MaxRecs) + ' Zeilen.');

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
      ProgressBar1.Position := 0;

    end;
  end;
end;

function TForm1.OpenExplorer(FName: string): boolean;
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

procedure TForm1.QFaks2DbaseClick(Sender: TObject);
var
  x: integer;
  Dbase: TDBF;
begin
  try
    (* wird nicht genutzt, da DBase IV Feldlänge 11 Zeichen und höher geht nicht *)
    Jei;
    Dbase := TDBF.Create(Application.Mainform);
    Dbase.TableLevel := 4;
    Dbase.FieldDefs.Assign(QFaks.FieldDefs);
    Dbase.TableName := Exepath + 'FaksExport.dbf';
    Dbase.CreateTable;
    Dbase.Open;
    BM := QFaks.GetBookmark;

    QFaks.DisableControls;

    QFaks.First;


    while not QFaks.EOF do
    begin

      Dbase.Append;

      for x := 0 to QFaks.FieldCount - 1 do
      begin
        //if ((QFaks.Fields[x].FieldName <> 'BETRAG') and (QFaks.Fields[x].FieldName <> 'EINZELPREIS')) then
        //ShowMessage(QFaks.Fields[x].FieldName + ' ' + QFaks.Fields[x].AsString);
        Dbase.Fields[x].AsVariant := QFaks.Fields[x].AsVariant;

      end;

      Dbase.Post;

      QFaks.Next;

    end;

    ShowMessage('Die Daten stehen jetzt in: ' + NL + NL + EXEPATH +
      Dbase.TableName + NL + NL + '... der Explorer wird geöffnet!' + NL +
      NL + 'Achtung: die Feldnamen haben nur 10 Buchstaben und es gibt keine Umlauts!!');

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

(* Die Datei aus Elgeba-FKE wird aktualisiert:

   bei MTV-Sued kommen auch Verkaufsdatensätze mit LinienNummern
   401 ... die in 991401 usw übersetzt weden müssen
   Übersetzung laut Tabelle GridReplaceLineNumber


*)
procedure TForm1.UpdateElgebaLinienNrClick(Sender: TObject);
var Elgeba : TDBF;
    x : integer;
begin
  try
    Elgeba := TDBF.Create(Application.MainForm);
    OpenDialog1.FilterIndex:=2;
    if DirectoryExists('F:\AMISdata\Import\Elgeba-FKE\') then
    OpenDialog1.InitialDir:='F:\AMISdata\Import\Elgeba-FKE\'
    else
    OpenDialog1.InitialDir:='F:\AMISdata\AMISdata\Import\Elgeba-FKE\';


  if Messagedlg('Ab 11.12.2016. Sollen die LinienNummern für MTK-Süd laut Tabelle auf Register ''Einstellungen'' wirklich berichtigt werden?',mtConfirmation,[mbYes,mbNo],0)= mrNo then exit;

  if OpenDialog1.Execute then
  begin
    Jei;
    Elgeba.TableName:=OpenDialog1.FileName;
    Elgeba.Open;
    ProgressBar1.Position:=0;
    ProgressBar1.Max:=Elgeba.RecordCount;
    StatusBar1.SimpleText :='In Datei ' + Elgeba.TableName + '=' + IntToStr(Elgeba.RecordCount) + ' Datensätze werden bei LinienNummern korrigiert!';
    Application.ProcessMessages;
    Elgeba.First;

    while not Elgeba.EOF do
    begin
      (* zu ersetzende LinienNummer nachschlagen *)
      if Elgeba.FieldByName('Datumv').AsDateTime >= StrToDate('11.12.2016') then
      begin
        Elgeba.Edit;
        Elgeba.FieldByName('Linie').AsString := LookUpStringGrid(GridReplaceLineNumber,'5',Elgeba.FieldByName('LINIE').AsString);
        Elgeba.Post;
      end;


      ProgressBar1.Position := Elgeba.RecNo;
      Elgeba.Next;
    end;
    Elgeba.Close;
    ShowMessage('Die LinienNummern für MTK-Süd wurden aktualisiert!');
  end;

  finally
    NEI;
    FreeAndNil(Elgeba);
    ProgressBar1.Position:=0;
    Application.ProcessMessages;

  end;
end;

procedure TForm1.UpDown1Click(Sender: TObject; Button: TUDBtnType);
begin

  if Button = btNext then
  begin
   if SQLHistoryIndex < SQLHistory.Count -1  then
   begin
     inc(SQLHistoryIndex,1);

     StrToStrings(SQLHistory[SQLHistoryIndex], '°', Memo1.Lines, True);

     case Memo1.Color of
       clWhite:
         Memo1.Color := clMoneyGreen;
       clSkyBlue:
         Memo1.Color := clMoneyGreen;
       clMoneyGreen:
         Memo1.Color := clSkyBlue;
     end;


   end;

  end
  else
  begin
    if SQLHistoryIndex > 0 then
    begin
      dec(SQLHistoryIndex,1);

      StrToStrings(SQLHistory[SQLHistoryIndex], '°', Memo1.Lines, True);

      case Memo1.Color of
        clWhite:
          Memo1.Color := clSkyBlue;
        clSkyBlue:
          Memo1.Color := clMoneyGreen;
        clMoneyGreen:
          Memo1.Color := clSkyBlue;
      end;


    end;


  end;

  StatusBar1.SimpleText:='SQL-History Eintrag ''' + IntToStr(SQLHistoryIndex +1) + '/' + IntToStr(SQLHistory.Count);
  Application.ProcessMessages;

end;

initialization

end.
