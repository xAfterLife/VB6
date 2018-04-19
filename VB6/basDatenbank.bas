Attribute VB_Name = "basDatenbank"
Option Explicit
Private db As Database

'!!!WICHTIG!!!
'DAO OBJECT muss als Verweis hinzugefügt sein.


'Öffnet eine Verbindung zur Datenbank
Public Sub SetDB(PfadDB As String)
10        On Error GoTo ErrHandler
20        Set db = OpenDatabase(PfadDB)
30        GoTo EndErrHandler
ErrHandler:
40        Select Case FehlerMeldung(Err, "Modul Common(SetDB)", Erl, Error$)
          Case 0
50            Resume Next
60        Case 1
70            Resume
80        End Select
EndErrHandler:
End Sub

'Liest einen Datensatz aus der Derzeit geöffneten Datenbank aus.
'BSP: DBSelect False, "Benutzer", "ID", "Name='PETER'", "ID"
'Beispiel für INNER JOIN:
'SELECT Orders.*, Customers.ID //Reihe
'FROM Orders INNER JOIN Customers ON Orders.CustomerID=CustomerID //Tabelle
'Sprich: DBSelect False, "Orders INNER JOIN Customers ON Orders.CustomerID=CustomerID", "Orders.*, Customers.ID"
Public Function DBSelect(Distinct As Boolean, Tabelle As String, Reihe As String, Optional Bedingung As String, Optional OrdnenNach As String) As DAO.Recordset
          Dim SQL As String
          Dim rs As Recordset

10        On Error GoTo ErrHandler

20        SQL = "SELECT "
30        If Distinct = True Then
40            SQL = SQL & "DISTINCT "
50        End If
60        SQL = SQL & Reihe & " FROM " & Tabelle
70        If Bedingung <> "" Then
80            SQL = SQL & " WHERE " & Bedingung
90        End If
100       If OrdnenNach <> "" Then
110           SQL = SQL & " ORDER BY " & OrdnenNach
120       End If

130       Set DBSelect = db.OpenRecordset(SQL)

140       GoTo EndErrHandler
ErrHandler:
150       Select Case FehlerMeldung(Err, "Modul Common(DBSelect)", Erl, Error$)
          Case 0
160           Resume Next
170       Case 1
180           Resume
190       End Select
EndErrHandler:
End Function

'Updated Werte in der derzeit geöffneten Datenbank
'BSP: DBUpdate "Benutzer", "Name='Peter'", "ID < 10"
Public Sub DBUpdate(Tabelle As String, Änderung As String, Optional Bedingung As String)
          Dim SQL As String

10        On Error GoTo ErrHandler

20        SQL = "UPDATE "
30        SQL = SQL & Tabelle & " SET "
40        SQL = SQL & Änderung
50        If Bedingung <> "" Then
60            SQL = SQL & " WHERE " & Änderung
70        End If

80        db.Execute (SQL)

90        GoTo EndErrHandler
ErrHandler:
100       Select Case FehlerMeldung(Err, "Modul Common(DBUpdate)", Erl, Error$)
          Case 0
110           Resume Next
120       Case 1
130           Resume
140       End Select
EndErrHandler:
End Sub

'Fügt einen Datensatz der Derzeit geöffneten Datenbank hinzu
'BSP: DBInsert "Benutzer", "Name", "'Baum'"
Public Sub DBInsert(Tabelle As String, Reihen As String, Wert As String)
          Dim SQL As String

10        On Error GoTo ErrHandler

20        SQL = "INSERT INTO "
30        SQL = SQL & Tabelle & "( "
40        SQL = SQL & Reihen & ") VALUES("
50        SQL = SQL & Wert & ")"

60        db.Execute (SQL)

70        GoTo EndErrHandler
ErrHandler:
80        Select Case FehlerMeldung(Err, "Modul Common(DBInsert)", Erl, Error$)
          Case 0
90            Resume Next
100       Case 1
110           Resume
120       End Select
EndErrHandler:
End Sub

'Löscht Werte aus der Derzeit geöffneten Datenbank
'BSP: DBDelete "Benutzer", "Name = 'Peter'"
'Bedingung kann mit "AND nächste Bedingung" erweiter werden
Public Sub DBDelete(Tabelle As String, Bedingung As String)
          Dim SQL As String

10        On Error GoTo ErrHandler

20        SQL = "DELETE FROM "
30        SQL = SQL & Tabelle & "WHERE "
40        SQL = SQL & Bedingung

50        db.Execute (SQL)

60        GoTo EndErrHandler
ErrHandler:
70        Select Case FehlerMeldung(Err, "Modul Common(DBInsert)", Erl, Error$)
          Case 0
80            Resume Next
90        Case 1
100           Resume
110       End Select
EndErrHandler:
End Sub

'Möglichkeit aus der aktuell geöffneten Datenbank einen Wert auszulesen oder einen SQL Befehl auszuführen in einer Funktion
Public Function DBSQL(SQL As String) As DAO.Recordset
10        On Error GoTo ErrHandler

20        If UCase(Left(SQL, 6)) = "SELECT" Then
30            Set DBSQL = db.OpenRecordset(SQL)
40        Else
50            db.Execute (SQL)
60        End If

70        GoTo EndErrHandler
ErrHandler:
80        Select Case FehlerMeldung(Err, "Modul Common(DBInsert)", Erl, Error$)
          Case 0
90            Resume Next
100       Case 1
110           Resume
120       End Select
EndErrHandler:
End Function

'Fehlerbehandlung eben
Function FehlerMeldung(nErrNum&, sErrMod$, nErrLine&, sErrText$) As Integer
10        If MsgBox("Fehlernummer: " & nErrNum & vbNewLine & "Modul: " & sErrMod & vbNewLine & "Zeile: " & nErrLine & vbNewLine & "Fehlertext: " & sErrText & vbNewLine, vbInformation & vbRetryCancel, "Houston, wir haben ein Problem!") = vbRetry Then
20            FehlerMeldung = 1
30        Else
40            FehlerMeldung = 0
50        End If
End Function
