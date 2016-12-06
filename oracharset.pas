program oracharset;
 
{ Shows client and server character set/NLS info}
{ PLEASE EDIT PASSWORDS ETC BELOW. }
 
{$mode objfpc}{$H+}
 
uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Classes,
  SysUtils,
  sqldb,
  oracleconnection;
 
var
  Col: integer;
  Conn: TOracleConnection;
  Tran: TSQLTransaction;
  Q: TSQLQuery;
begin
  Conn := TOracleConnection.Create(nil);
  Tran := TSQLTransaction.Create(nil);
  Q := TSQLQuery.Create(nil);
  try
    // * EDIT IDENTIFYING INFO AS NEEDED*
    Conn.HostName := 'hlbst02';
    Conn.UserName := 'sysadm';
    Conn.Password := 'sysadm';
    Conn.DatabaseName := 'VASP';
    // *END IDENTIFIYING INFO*
    Conn.Transaction := Tran;
    Q.DataBase := Conn;
    Conn.Open;
    Tran.Active := true;
 
    writeln('Server character set info:');
    Q.SQL.Text := 'SELECT value$ FROM sys.props$ WHERE name like ''NLS_%'' ';
    Q.Open;
    Q.First;
    while not (Q.EOF) do
    begin
      writeln('*****************');
      for Col := 0 to Q.Fields.Count - 1 do
      begin
        try
          writeln(Q.Fields[Col].DisplayLabel + ':');
          writeln(Q.Fields[Col].AsString);
        except
          writeln('Error retrieving field ', Col);
        end;
      end;
      Q.Next;
    end;
    Q.Close;
 
    writeln('');
    writeln('Client character set info:');
    Q.SQL.Text := 'SELECT * FROM NLS_SESSION_PARAMETERS ';
    Q.Open;
    Q.First;
    while not (Q.EOF) do
    begin
      writeln('*****************');
      for Col := 0 to Q.Fields.Count - 1 do
      begin
        try
          writeln(Q.Fields[Col].DisplayLabel + ':');
          writeln(Q.Fields[Col].AsString);
        except
          writeln('Error retrieving field ', Col);
        end;
      end;
      Q.Next;
    end;
    Q.Close;
    // *END EXAMPLE BUG TESTING CODE*
    Conn.Close;
  finally
    Q.Free;
    Tran.Free;
    Conn.Free;
  end;
  writeln('Program complete. Press a key to continue.');
  readln;
end.