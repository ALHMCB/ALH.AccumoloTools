<%
Sub PermanentlyMoved(byval newUrl)
	Response.Status = "301 Permanently Moved"
	Response.AddHeader "Location", newUrl
	Response.Write "<HTML>"
	Response.Write "<BODY>"
	Response.Write "This file was moved permanently to "
	Response.Write "<A HREF=""" & newUrl & """>here.<A>"
	Response.Write "</BODY>"
	Response.Write "</HTML>"
	Response.End
End Sub
 %>