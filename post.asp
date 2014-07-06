<!-- Copyright Robert Roos, Roosweb.se 2007 All rights reserved -->

<%
' Deklarerar Variabler
Dim Namn,Epost,Hemsida,Post,Connect,strSQL
Namn = Request.Querystring("namn")
Epost = Request.Querystring("epost")
Hemsida = Request.Querystring("hemsida")
Post = Request.Querystring("post")
IPnummer = Request.ServerVariables("REMOTE_ADDR")


' Kollar så att alla formulär är ifyllda, kör detta med serversidan kod då man kan avaktivera
' Javascript.

If Trim(Namn) = "" Or Trim(Epost) = "" Or Trim(Post) = "" Then
Response.Write("Du måste fylla i alla fälten markerade med *")
Else

' Skydd mot SQL-injections
Function SqlSkydd(ByVal strText)
 strText = Replace(strText,"'","")
 strText = Replace(strText,"\","\\")
 strText = Replace(strText,";","")
 SqlSkydd = strText
End Function

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Lägger in informationen från formuläret i databasen
strSQL = "INSERT INTO tblGastbok(Namn, Datum, Epost, Hemsida, Post, IPnummer) VALUES('"& SqlSkydd(Namn) &"','"& Date() &"','"& SqlSkydd(Epost) &"','"& SqlSkydd(Hemsida) &"','"& SqlSkydd(Post) &"','"& SqlSkydd(IPnummer) &"')"
Connect.Execute(strSQL)
 
' Här stänger jag databaskopplingen igen
Connect.Close : Set Connect = Nothing
End If
%>