<%
' Deklarerar Variabler
Dim strSQL, Connect

' Kollar om man �r inloggad admin
If Session("AdminOnline")="True" Then

' �ppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Tar bort det valda g�stboksinl�gget
strSQL="DELETE FROM tblGastbok Where ID="& clng(Request.Querystring("ID")) &""
Connect.Execute(strSQL)

'Response.Write("<span style=""color: red;"">Inl�gget �r raderat!</span>")

' St�nger DB koppling
Connect.Close
Set Connect = Nothing

Else
Response.Write("Du �r inte inloggad som admin!")
End If
%>