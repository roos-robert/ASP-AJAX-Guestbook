<%
' Deklarerar Variabler
Dim strSQL, Connect

' Kollar om man är inloggad admin
If Session("AdminOnline")="True" Then

' Öppnar databasanslutningen
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("databas/RWAjaxDB.mdb") &";"

' Tar bort det valda gästboksinlägget
strSQL="DELETE FROM tblGastbok Where ID="& clng(Request.Querystring("ID")) &""
Connect.Execute(strSQL)

'Response.Write("<span style=""color: red;"">Inlägget är raderat!</span>")

' Stänger DB koppling
Connect.Close
Set Connect = Nothing

Else
Response.Write("Du är inte inloggad som admin!")
End If
%>