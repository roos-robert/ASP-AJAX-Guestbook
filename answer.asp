<%
If Request.Querystring("page")="" Then
%>
<b>Svara</b><br />
<textarea id="Svaret" cols="60" rows="5"></textarea><br />
<b>Svar av:</b><br />
<input type="text" id="Admin" size="25" />
<input type="hidden" value="answer" id="page" />
<input type="hidden" value="<%=Request.Querystring("ID")%>" id="PostID" />
<p></p>
<input type="button" value="Svara" id="btnSend" onclick="makeRequest('answer.asp?page='+document.getElementById('page').value+'&PostID='+document.getElementById('PostID').value+'&Svaret='+document.getElementById('Svaret').value+'&Admin='+document.getElementById('Admin').value+'&rnd='+Math.random(), 'showPosts');" />

<%
ElseIf Request.Querystring("page")="answer" Then

' Deklarerar Variabler
Dim PostID,Svaret,Admin,Connect,strSQL
PostID = Request.Querystring("PostID")
Svaret = Request.Querystring("Svaret")
Admin = Request.Querystring("Admin")


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
strSQL = "INSERT INTO tblSvar(PostID, Svaret, Admin) VALUES('"& SqlSkydd(PostID) &"','"& SqlSkydd(Svaret) &"','"& SqlSkydd(Admin) &"')"
Connect.Execute(strSQL)

Response.Write("Ditt svar &auml;r nu sparat!")
 
' Här stänger jag databaskopplingen igen
Connect.Close : Set Connect = Nothing
End If
%>