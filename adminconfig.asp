<%
' Deklarerar Variabler
Dim Connect,objRS,strSQL

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Tar fram admin ur databasen
Set objRS = Connect.Execute("SELECT * FROM tblAdmin ORDER BY ID DESC")
If objRS.EOF Then
Response.Write("Inga inlägg i gästboken just nu, bli den förste!")
Else
%>

<table align="center" width="100%" cellpadding="5">
  <tr>
    <td width="120"><b>Nytt Adminnamn:</b></td>
    <td><input type="text" value="<%=objRS("anv")%>" id="anv" size="25" /></td>
  </tr>
  <tr>
    <td><b>Nytt L&ouml;senord:</b></td>
    <td>
    <input type="password" value="<%=objRS("losen")%>" id="losen" size="25" />
    <input type="hidden" value="updateAdmin" id="updateAdmin" />
    <input type="hidden" value="<%=objRS("ID")%>" id="ID" />
    </td>
  </tr>
  <tr>
    <td><input type="button" value="Uppdatera" id="btnUpdate" onclick="makeRequest('adminconfig.asp?page='+document.getElementById('updateAdmin').value+'&anv='+document.getElementById('anv').value+'&losen='+document.getElementById('losen').value+'&ID='+document.getElementById('ID').value+'&rnd='+Math.random(), 'showPosts');" /></td>
    <td></td>
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