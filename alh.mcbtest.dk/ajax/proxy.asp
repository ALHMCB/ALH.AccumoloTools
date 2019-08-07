<%
doDbWrite = true
%>
<!--#include virtual="/includes/basic_inc.asp"-->
<!--#include virtual="/ajax/site_utils.asp"-->
<!--#include virtual="/ajax/goglobal.asp"-->
<!--#include virtual="/ajax/domf.asp"-->

<%
postAction = replace(Request.Form("f"),"'","''")
siteguid = replace(Request.Form("siteguid"),"'","''")

select case postAction
	case "loadSites"
		call getSiteList()
	case "loadData"
		call getLangAndCurrency()
	case "loadLC"
		call getAvailableLangAndCurrency()
	case "loadCountry"
		call getCountryList()
	case "addEntity"
		call addEntity()
	case "loadDOMFStatus"
		call getDOMFStatus()
	case "makeDOMF"
		call makeDOMF()
end select


function truncateJSONString(json)
	if len(json) > 1 then
		json = Left(json, len(json) - 1)
	end if
	truncateJSONString = json
end function

%>
<!--#include virtual="/global_inc/site_includes/conn_close.asp"-->
<%if doDbWrite then%>
<!--#include virtual="/global_inc/site_includes/connWrite_close.asp"-->
<%end if%>