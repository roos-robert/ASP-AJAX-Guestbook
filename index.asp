<!--#include file="includes/RegExp.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<title>Roosweb:s Ajax G�stbok - Version 1.0 BETA</title>
<link rel="stylesheet" href="includes/ajaxGB.css" />
<script type="text/javascript" src="RWAjax.js"></script>
</head>

<body onload="makeRequest('getpost.asp?rnd='+Math.random(),'divid')">
<!-- Copyright Robert Roos, Roosweb.se 2007 All rights reserved -->
<div id="Huvudruta">
<h4>V�lkommen till Roosweb:s Ajax G�stbok - Version 1.0 BETA.</h4>
<h5 style="margin-top: -15px;"><a href="#">Operation Helix</a></h5>
V�lkommen till BETA versionen av min Ajax g�stbok, du f�r anv�nda detta script som du vill p� din
egen webbsida s� l�nge copyright texterna st�r kvar! Support vid buggar kan du f� via forumet p� 
Roosweb.se
<table style="margin-top: 10px;" border="0" align="center" width="90%" cellspacing="5" id="table1">
	<tr>
		<td width="80"><b>Namn:</b>&nbsp;*</td>
		<td><input id="namn" type="text" name="namn" size="25" /></td>
	</tr>
	<tr>
		<td><b>Epost:</b>&nbsp;*</td>
		<td><input id="epost" type="text" name="epost" size="25" /></td>
	</tr>
	<tr>
		<td><b>Hemsida:</b></td>
		<td><input id="hemsida" type="text" name="hemsida" size="25" /></td>
	</tr>
	<tr>
		<td><b>Inl�gget:</b>&nbsp;*</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
	    <td>&nbsp;</td>
		<td><textarea id="post" name="post" cols="45" rows="9"></textarea></td>
	</tr>
	<tr>
	    <td>&nbsp;</td>
		<td><input type="button" value="Skriv inl�gg" id="btnSend" onclick="makeRequest( 'post.asp?namn='+document.getElementById('namn').value+'&epost='+document.getElementById('epost').value+'&hemsida='+document.getElementById('hemsida').value+'&post='+document.getElementById('post').value+'&rnd='+Math.random(), 'showAlerts'); makeRequest('getpost.asp?rnd='+Math.random(),'showPosts');" /></td>
	</tr>
</table>
<span style="font-size: 9px;">* = M�ste fyllas i</span>


<table width="100%" border="0">
<tr>
 <td align="left"><a href="#" onclick="makeRequest('getpost.asp?rnd='+Math.random(),'showPosts');">Visa g�stboksinl�gg</a></td>
 <td align="right"><a href="#" onclick="makeRequest('login.asp?rnd='+Math.random(), 'showPosts'); ">Administrera</a></td>
</tr>
</table>

<hr />
<div id="showAnswer"></div><br />
<div id="showAlerts"></div><br />
<div id="showPosts"></div>

</div>

</body>
</html>