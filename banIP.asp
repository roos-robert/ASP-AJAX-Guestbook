<!-- Copyright Robert Roos, Roosweb.se 2007 All rights reserved -->

<%

' Deklarerar Variabler
Dim IPnummer
IPnummer = Request.Querystring("IPnummer")

' Kollar s� att alla formul�r �r ifyllda, k�r detta med serversidan kod d� man kan avaktivera
' Javascript.

If Trim(IPnummer) = "" Then
Response.Write("G&aring;r inte att banna detta IP-nummer.")
Else

' �ppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' L�gger in informationen fr�n formul�ret i databasen
strSQL = "INSERT INTO tblBanningar(IPnummer) VALUES('"& IPnummer &"')"
Connect.Execute(strSQL)
Response.Write("IP-nummret &auml;r nu blockerat!")
 
' H�r st�nger jag databaskopplingen igen
Connect.Close : Set Connect = Nothing
End If
%>