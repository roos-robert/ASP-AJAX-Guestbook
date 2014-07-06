<!--#include file="includes/RegExp.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<title>Roosweb:s Ajax Gästbok - Version 1.0 BETA</title>
<link rel="stylesheet" href="includes/ajaxGB.css" />
<script type="text/javascript" src="RWAjax.js"></script>
</head>

<body>
<% 
If Request.QueryString("page")="" Then
If Session("AdminOnline")="True" Then
' Om man är inloggad som admin visas allt mellan If och Else nedan
Response.Write("V&auml;lkommen, du &auml;r nu inloggad som administrat&ouml;r.")
%>
<p></p>
--><a href="#" onclick="makeRequest('logout.asp?rnd='+Math.random(),'showPosts');"> Logga ut</a><br />
--><a href="#" onclick="makeRequest('getIP.asp?rnd='+Math.random(),'showPosts');"> IP nummer / Banningar</a><br />
--><a href="#" onclick="makeRequest('adminconfig.asp?rnd='+Math.random(),'showPosts');"> Inst&auml;llningar</a><br />
<p></p>
<% 
' Om man inte är inloggad admin visas allt mellan Else och End If
Else 
%>
<table width="100%" cellspacing="5">
<tr>
<td width="120"><b>Anv&auml;ndarnamn:</b></td>
<td><input type="text" size="25" id="anv" name="anv" /></td>
</tr>
<tr>
<td width="120"><b>L&ouml;senord:</b></td>
<td><input type="password" size="25" id="losen" name="losen" />
    <input type="hidden" id="page" name="login" value="login" /></td>
</tr>
<tr>
<td width="120"><input type="submit" value="Logga In" onclick="makeRequest( 'login.asp?page='+document.getElementById('page').value+'&anv='+document.getElementById('anv').value+'&losen='+document.getElementById('losen').value+'&rnd='+Math.random(), 'showPosts');" /><br />
</td>
<td>&nbsp;</td>
</tr>
</table>
<% End If %>

<%
' När man klickar på logga in så körs denna koden för att logga in.
ElseIf Request.QueryString("page")="login" Then

' Deklarerar Variabler
Dim anv, losen, strSQL, Connect
anv = Request.Querystring("anv")
losen = Request.Querystring("losen")

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

' Tar fram info ur databasen för att se om anv och lösen finns
strSQL="SELECT * From tblAdmin where anv='"& SqlSkydd(anv) &"' AND losen='"& SqlSkydd(losen) &"'"
Set objRS = Connect.Execute(strSQL)

' Om adminen inte finns får man ett felmess
If objRS.EOF Then
Response.Write("<font color='Red'>Fel anv&auml;ndarnamn eller l&ouml;senord, f&ouml;rs&ouml;k igen.</font><br />")

' Annars skapas en adminsession
Else
    Session("AdminOnline")="True"
    Session.TimeOut = 999	

' Stänger databasen och recsetet vi öppnade
objRS.Close : Set objRS = Nothing
Connect.Close : Set Connect = Nothing

Response.Redirect("login.asp")

End If
End If
%>
</body>
</html>