<%
if connWriteCreated&""="" then
	Set connWrite = Server.CreateObject("ADODB.Connection")
	connWrite.Open Application("ConnWrite")
	connWriteCreated = true
end if

sub connWriteClose()
  if (connWriteCreated) then
  	connWrite.close
	  set connWrite = nothing
	end if
	connWriteCreated = false
end sub
%>
