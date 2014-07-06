<!--#include file="includes/RegExp.asp" -->
<%
' Deklarerar Variabler
Dim Connect,objRS

' Ser till att vi får datumformateringen på Svenska
Session.LCID = 1053
Response.Buffer = True

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Tar fram allting ur databasen och lägger det i ett pagingsystem
Set objRS = Connect.Execute("SELECT * FROM tblGastbok ORDER BY ID DESC")
If objRS.EOF Then
Response.Write("Inga inl&auml;gg i g&auml;stboken just nu, bli den f&ouml;rste!")
Else
Do until objRS.EOF
%>

<table align="center" width="100%" cellpadding="5" style="border: 1px solid; border-color: #cacaca;">
  <tr>
    <td style="background-color: #fefc9f;">
    <div style="float: left;">
    <% If Session("AdminOnline")="True" Then %>
    <a href="#" onclick="makeRequest(' answer.asp?ID='+document.getElementById('ID').value+'&rnd='+Math.random(), 'showAnswer');"><span style="color: green;">Svara</span></a>&nbsp;-
	<a href="#" onclick="makeRequest(' deletepost.asp?ID='+document.getElementById('ID').value+'&rnd='+Math.random(), 'showAlerts'); makeRequest('getpost.asp?rnd='+Math.random(),'showPosts');"><span style="color: red;">Ta bort</span></a>&nbsp;-&nbsp;
    <%
	Else
	End If %>
    <a href="#"><%=ReplaceMent(objRS("Namn"))%></a>&nbsp;-&nbsp;<%=FormatDateTime(objRS("Datum"), 1)%></div>    <% If objRS("Hemsida")="" Then 
	   Else %><div style="float: right"><a href="<%=objRS("Hemsida")%>" target="_blank">http://</a><input type="hidden" value="<%=objRS("ID")%>" name="ID" id="ID" /></div> <% End If %></td>
  </tr>
  <tr>
    <td><%=ReplaceMent(objRS("Post"))%></td>
  </tr>
</table>
<p></p>
<%
objRS.MoveNext
Loop
End If

' Stänger databasen och recsetet vi öppnade
objRS.Close : Set objRS = Nothing
Connect.Close : Set Connect = Nothing
%>