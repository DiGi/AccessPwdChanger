program AccessPwdChanger;
{$APPTYPE CONSOLE}

uses
  SysUtils, Windows, ActiveX, AdoDb;

const
  // ADODB p�ipojovac� �et�zec s p�ipojen�m na Syst�movou datab�zi s hesly
  C_Connection = 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:System database=%s;Mode=Share Deny None;Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;' +
    'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:SFP=False;User ID=%s;Password=%s';
  // samotn� m�n�c� SQL dotaz
  C_ChangePass = 'ALTER USER [%s] PASSWORD %s %s';

// chybov� hl�en� ve spr�vn�m k�dov�n�
function CharToOem(const s: string): string;
begin
  SetLength(Result, Length(s));
  CharToOemA(PAnsiChar(s), PAnsiChar(Result));
end;

// chybov� zpr�vy s ExitCode
procedure ErrorMessage(const s: string; const Code: Integer = 1);
begin
  Writeln(Format('Chyba: %s', [s]));// + CharToOem(s));
  ExitCode := Code;
end;

// Hlavn� procedura se zm�nu hesla
procedure ChangePassword;
var
  DbConnection : TADOConnection;
  Query : TADOQuery;
begin
  DbConnection := TADOConnection.Create(nil);
  Query := TADOQuery.Create(nil);
  try
    if ParamCount = 6 then // p�il�en� Adminem
      DbConnection.ConnectionString := Format(C_Connection,
        [ParamStr(1), ParamStr(2), ParamStr(5), ParamStr(6)])
    else                   // p�ihl�en� u�ivatelem
      DbConnection.ConnectionString := Format(C_Connection,
        [ParamStr(1), ParamStr(2), ParamStr(3), ParamStr(4)]);
    DbConnection.LoginPrompt := False;
    try
      DbConnection.Connected := True;
//      Writeln('Connected');// + CharToOem(s));
    except
      on E: Exception do
        begin
          ErrorMessage(E.Message, 10);
          Exit;
        end;
    end;
    try
      if DbConnection.Connected then
      begin
        Query.Connection := DbConnection;
        if ParamCount = 6 then  // zm�na hesla adminem
          Query.SQL.Text := Format(C_ChangePass,
            [ParamStr(3), ParamStr(4), ''''''])
        else                    // zm�na hesla u�ivatelem
          Query.SQL.Text := Format(C_ChangePass,
            [ParamStr(3), ParamStr(5), ParamStr(4)]);
        // proveden� zm�ny hesla
        Query.ExecSQL;
      end;
    except
      on E: Exception do
        begin
          ErrorMessage(E.Message, 20);
          Exit;
        end;
    end;
  finally
    Query.Close;
    Query.Free;
    DbConnection.Close;
    DbConnection.Free
  end;
end;

begin
  if ParamCount in [5, 6] then
  begin
    CoInitialize(nil);
    ChangePassword
  end
  else
  begin
    writeln('MS Access MDW Password Changer');
    if ParamCount < 5 then
      ErrorMessage('Nedostatek parametr�')
    else
      ErrorMessage('P��li� mnoho parametr�. Uzav�ete cesty nebo loginy s mezerami do uvozovek (nap�: "c:\Program Files\Databaze\Soubor.mdb")');
    ErrorMessage('Pou�it�:');
    writeln(' AccessPwdChanger DB.mdb WorkGroup.mdw Uzivatel StareHeslo NoveHeslo');
    writeln(' AccessPwdChanger DB.mdb WorkGroup.mdw Uzivatel NoveHeslo dbAdmin dbAdminPwd');
  end;
end.
