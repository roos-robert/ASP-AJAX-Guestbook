<%
' Deklarerar Variabler
Dim Connect,objRS,strSQL

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Tar fram admin ur databasen
Set objRS = Connect.Execute("SELECT IPnummer, Datum, Namn FROM tblGastbok")
If objRS.EOF Then
Response.Write("Inga IP-nummer i databasen för tillfället")
Else
%>


<table align="center" width="100%" cellpadding="2">
  <tr>
    <td><b>Namn:</b></td>
    <td><%=objRS("Namn")%></td>
  </tr>
    <tr>
    <td width="120"><b>Datum</b></td>
    <td><%=objRS("Datum")%></td>
  </tr>
    <tr>
    <td width="120"><b>IP-nummer</b></td>
    <td><%=objRS("IPnummer")%>
    <input type="hidden" id="IPnummer" name="IP" value="<%=objRS("IPnummer")%>" /></td>
  </tr>
    <tr>
    <td width="120"><b><a href="#" onclick="makeRequest('banIP.asp?page='+document.getElementById('IPnummer').value+'&rnd='+Math.random(), 'showPosts');"><span style="color: red;">Banna IP</span></a></b></td>
  </tr>
</table>

<%
' Stänger IF Sats
End If

' Om man valt att uppdatera uppgifterna körs denna koden
If Request.Querystring("page")="updateAdmin" Then

' Skydd mot SQL-injections
Function SqlSkydd(ByVal strText)
 strText = Replace(strText,"'","")
 strText = Replace(strText,"\","\\")
 strText = Replace(strText,";","")
 SqlSkydd = strText
End Function

' Uppdaterar databasen med den nya informationen
strSQL="UPDATE tblAdmin Set anv='"& SqlSkydd(Request.Querystring("anv")) &"' , losen='"& SqlSkydd(Request.Querystring("losen")) &"' Where ID="&clng(Request.Querystring("ID"))&""
Connect.Execute(strSQL)

Response.Write("Uppgifterna &auml;r uppdaterade, klicka p&aring; administrera f&ouml;r att g&aring; till adminstartsidan igen!")
End If

' Stänger databasen och recsetet vi öppnade
objRS.Close : Set objRS = Nothing
Connect.Close : Set Connect = Nothing
%>