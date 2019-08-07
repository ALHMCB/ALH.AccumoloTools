<%
if connCreated&""="" then
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Conn")
	connCreated = true
end if

sub connClose()
  if (connCreated) then
  	conn.close
	  set conn = nothing
	end if
	connCreated = false
end sub
%>
