<!-- Copyright Robert Roos, Roosweb.se 2007 All rights reserved -->

<%

' Deklarerar Variabler
Dim IPnummer
IPnummer = Request.Querystring("IPnummer")

' Kollar så att alla formulär är ifyllda, kör detta med serversidan kod då man kan avaktivera
' Javascript.

If Trim(IPnummer) = "" Then
Response.Write("G&aring;r inte att banna detta IP-nummer.")
Else

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Lägger in informationen från formuläret i databasen
strSQL = "INSERT INTO tblBanningar(IPnummer) VALUES('"& IPnummer &"')"
Connect.Execute(strSQL)
Response.Write("IP-nummret &auml;r nu blockerat!")
 
' Här stänger jag databaskopplingen igen
Connect.Close : Set Connect = Nothing
End If
%>