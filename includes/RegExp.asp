<%
Sub FixReg(re,pattern,replace,str)
    re.pattern = pattern
    str = re.replace(str,replace)
End Sub

Function ReplaceMent(ByVal strText)
  Dim Reg
  set Reg = new RegExp
    Reg.Global=True
    Reg.IgnoreCase=True
    strText = Replace(strText,":(","<img src=""community/Includes/smileys/sur.gif"">")
    strText = Replace(strText,":-(","<img src=""community/Includes/smileys/grater.gif"">")
    strText = Replace(strText,";)","<img src=""community/Includes/smileys/blink.gif"">")
    strText = Replace(strText,":D","<img src=""community/Includes/smileys/superglad.gif"">")
    strText = Replace(strText,":0","<img src=""community/Includes/smileys/smiley3.gif"">")
    strText = Replace(strText,":S","<img src=""community/Includes/smileys/smiley5.gif"">")
    strText = Replace(strText,":@","<img src=""community/Includes/smileys/smiley7.gif"">")
	strText = Replace(strText,"�","&Aring;")
    strText = Replace(strText,"�","&Auml;")
    strText = Replace(strText,"�","&Ouml;")
    strText = Replace(strText,"�","&aring;")
    strText = Replace(strText,"�","&auml;")
    strText = Replace(strText,"�","&ouml;")
    strText = Replace(strText, "[p]", "<p></p>")
    FixReg Reg, "(^|\W):\)(?=$|\W)", "$1<img src=""community/Includes/smileys/glad.gif"">", strText
    FixReg Reg, "(^|\W)><(?=$|\W)", "$1<img src=""community/Includes/smileys/ahh.gif"">", strText
    FixReg Reg, "(^|\W):P(?=$|\W)", "$1<img src=""community/Includes/smileys/P.gif"">", strText
    FixReg Reg, "\<embed\>","****", strText
    FixReg Reg, "\<script\>","****Jag �r en d�lig hacker****", strText
    FixReg Reg, "[^\=](http://)([\w~%&=#:\-/\?\.]*[\w~%&=#:\-/])", "<a href=""$1$2"" target=""_blank"">$2</a>", strText
    FixReg Reg, "\[(b|\/b|i|\/u|i|\/i)\]", "<$1>", strText
ReplaceMent=strText
End Function		  
%>