<%
if connWriteCreated&""<>"" then
	connWrite.close
	set connWrite = nothing
end if
%>