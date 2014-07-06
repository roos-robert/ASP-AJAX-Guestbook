<!-- #include file = "Includes/connect.asp" -->
<%
    tid = now



      
Function Fixbug(strText)
strText = Replace(strText,"'","''")
Fixbug = strText
End Function

If Request.Form("kon")="Kille" Or Request.Form("kon")="Tjej" Then

      
     strSQL="UPDATE tblAdmins Set presentation='"& Fixbug(Request.Form("presentation")) &"' , hemsida='"& Fixbug(Request.Form("hemsida")) &"' , msn='"& Fixbug(Request.Form("msn")) &"' , namn='"& Fixbug(Request.Form("namn")) &"' , efternamn='"& Fixbug(Request.Form("efternamn")) &"' , adminEmail='"& Fixbug(Request.Form("email")) &"' , bild='"& Fixbug(Request.Form("bild")) &"' , civilstatus='"& Fixbug(Request.Form("civilstatus")) &"' , politik='"& Fixbug(Request.Form("politik")) &"' , lan='"& Fixbug(Request.Form("lan")) &"' , stad='"& Fixbug(Request.Form("stad")) &"' , kon='"& Fixbug(Request.Form("kon")) &"' , alder='"& Fixbug(Request.Form("alder")) &"' Where adminID="&clng(Session("adminID"))&""
     Connect.Execute(strSQL)


           ' Stänger DB koppling
             Connect.Close
               Set Connect = Nothing
    Else
	Response.Write("<script type='text/javascript'>alert('Försök inte!'); top.location.href = 'community.asp?page=installningar';</script>")
	End If


    Response.Redirect("community.asp?page=profilklar")
%>
