<%
function aspConsoleLog(value)
    response.Write("<script language=javascript>console.log(" & chr(34) & value & chr(34) &"); </script>")
end function

Response.CodePage = 65001
Response.CharSet = "utf-8"

sub closeConnections()
  call connClose()
  call connWriteClose()
end sub

if inStr(Request.ServerVariables("HTTP_USER_AGENT"),"iPad")>0 then
	isTablet = true
end if

if request.QueryString("standardCompliance")&"" = "1" then
	session("standardCompliance") = 1
elseif request.QueryString("standardCompliance")&"" = "0" then
	session("standardCompliance") = 0
end if



%>
<!--#include virtual="/global_inc/site_includes/conn_create.asp"-->
<%if doDbWrite then%>
<!--#include virtual="/global_inc/site_includes/connWrite_create.asp"-->
<%end if%>
<!--#include virtual="/global_inc/site_includes/variableText.asp"-->

<%

select case msLanguage
	case 1 ' Dansk
		session.LCID = 1030
	case 2 ' Engelsk
		session.LCID = 1033
	case 3 ' Finsk
		session.LCID = 1035
	case 4 ' Norsk
		session.LCID = 1044
	case 5 ' Svensk
		session.LCID = 1053
	case 6 ' Hollandsk
		session.LCID = 1043
	case 9 ' Tysk
		session.LCID = 1031
	case 1014 ' Tysk - Om CCHobby
		session.LCID = 1031
	case else
		session.LCID = 1030
end select
if msThisIsAvailable then
	response.Write("#"&vbCrLf)
end if
%>
<!--#include virtual="/global_inc/site_includes/findLinks.asp"-->
<!--#include virtual="/global_inc/site_includes_utf8/textReplace.asp"-->
<!--#include virtual="/global_inc/site_includes/301redirect.asp"-->
<!--#include virtual="/global_inc/site_includes_utf8/utils.asp"-->
<!--#include virtual="/global_inc/site_includes/stripHTML.asp"-->
<!--#include virtual="/includes/functions.asp"-->