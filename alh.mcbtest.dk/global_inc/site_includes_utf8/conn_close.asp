<%
if connCreated&""<>"" then
	conn.close
	set conn = nothing
end if
%>