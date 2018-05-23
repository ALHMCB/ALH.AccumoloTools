<%
doDbWrite = true
%>
<!--#include virtual="/includes/basic_inc.asp"-->

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
end select

sub getSiteList()
	siteSQL = "SELECT SiteGuid, CostumerFullname FROM TBLadmin_sites WHERE AccountingMarkAsInactive = 0 AND SiteTypeGuid in (3, 6) ORDER BY CostumerFullname"
	set rsSites = conn.Execute(siteSQL)

	if not rsSites.eof then
		xSites = ""
		do while not rsSites.eof
			xSites = xSites & "{""Guid"":""" & rsSites("SiteGuid") & """, ""Name"":""" & rsSites("CostumerFullname") & """},"
			rsSites.MoveNext
		loop

		xSites = truncateJSONString(xSites)
	end if
	rsSites.close
	set rsSites = nothing

	xSites = "[" & xSites & "]"

	response.AddHeader "X-JSON", "{""EntityType"":""Site""}"
	response.Write(xSites)
end sub

sub getCountryList()
	sqlGetCountryList = "SELECT Name, Guid, IsoCode, IsEuMember from TBLCommerce_Country " &_
					"WHERE [Exists] = 1 " &_
					"ORDER BY IsEuMember DESC, Name"

	set rsCountryList = conn.execute(sqlGetCountryList)
	if not rsCountryList.eof then
		countryList = ""

		do while not rsCountryList.eof
			countryList = countryList & "{""Name"":""" & rsCountryList("Name") & """, ""Guid"":""" & rsCountryList("Guid") & """, ""IsoCode"":""" & rsCountryList("IsoCode") & """, ""IsEuMember"":""" & rsCountryList("IsEuMember") & """},"
			rsCountryList.MoveNext
		loop

		countryList = truncateJSONString(countryList)
		rsCountryList.close
		set rsCountryList = nothing
	end if

	gData = "[" & countryList & "]"

	response.AddHeader "X-JSON", "{""EntityType"":""CountryList""}"
	response.Write(gData)
end sub

sub getLangAndCurrency()
	sqlLang = "SELECT  A.LanguageGuid AS [Guid], UPPER(B.Code) AS [Language], A.[Standard] " &_ 
				"FROM TBLcommerce_siteLanguage A INNER JOIN TBLCommerce_language B ON A.LanguageGuid = B.Guid " &_
				"WHERE (A.SiteGuid = " & siteguid & ") " &_
				"ORDER BY A.[Standard] DESC"
	set rsLang = conn.execute(sqlLang)
	if not rsLang.eof then
		gLang = ""

		do while not rsLang.eof
			gLang = gLang & "{""Language"":""" & rsLang("Language") & """, ""Guid"":""" & rsLang("Guid") & """, ""Standard"":""" & rsLang("Standard") & """},"
			rsLang.MoveNext
		loop

		gLang = truncateJSONString(glang)
		rsLang.close
		set rsLang = nothing
	end if

	sqlCurrency = "SELECT A.CurrencyGuid AS [Guid], B.AlphabeticCode AS CurrencyCode, A.[Standard] " &_ 
					"FROM TBLcommerce_siteCurrency A INNER JOIN TBLCommerce_currency B ON A.CurrencyGuid = B.Guid " &_
					"WHERE (A.SiteGuid = " & siteguid & ") " &_ 
					"ORDER BY A.[Standard] DESC"
	set rsCurrency = conn.execute(sqlCurrency)
	if not rsCurrency.eof then
		gCurr = ""

		do while not rsCurrency.eof
			gCurr = gCurr & "{""Currency"":""" & rsCurrency("CurrencyCode") & """, ""Guid"":""" & rsCurrency("Guid") & """, ""Standard"":""" & rsCurrency("Standard") & """},"
			rsCurrency.MoveNext
		loop

		gCurr = truncateJSONString(gCurr)
		rsCurrency.close
		set rsCurrency = nothing		
	end if

	gData = "{""Languages"":[" & gLang & "], ""Currencies"":[" & gCurr & "]}"

	response.AddHeader "X-JSON", "{""EntityType"":""AddedLangAndCurr""}"
	response.Write(gData)
end sub

sub getAvailableLangAndCurrency()
	sqlAvailLang = "SELECT [Guid], UPPER(Code) AS [Language] FROM TBLCommerce_language " &_ 
					"WHERE [Guid] NOT IN ( SELECT LanguageGuid FROM TBLcommerce_siteLanguage WHERE (SiteGuid = " & siteguid & ") ) " &_
					"ORDER BY [Language]"

	sqlAvailCurr = "SELECT [Guid], AlphabeticCode AS CurrencyCode FROM TBLCommerce_currency " &_ 
					"WHERE [Guid] NOT IN ( SELECT CurrencyGuid FROM TBLcommerce_siteCurrency WHERE (SiteGuid = " & siteguid & ") ) " &_
					"ORDER BY CurrencyCode"

	set rsAvailLang = conn.execute(sqlAvailLang)
	if not rsAvailLang.eof then
		availLang = ""

		do while not rsAvailLang.eof
			availLang = availLang & "{""Language"":""" & rsAvailLang("Language") & """, ""Guid"":""" & rsAvailLang("Guid") & """},"
			rsAvailLang.MoveNext
		loop

		availLang = truncateJSONString(availLang)
		rsAvailLang.close
		set rsAvailLang = nothing
	end if

	set rsAvailCurr = conn.execute(sqlAvailCurr)
	if not rsAvailCurr.eof then
		availCurr = ""

		do while not rsAvailCurr.eof
			availCurr = availCurr & "{""Currency"":""" & rsAvailCurr("CurrencyCode") & """, ""Guid"":""" & rsAvailCurr("Guid") & """},"
			rsAvailCurr.MoveNext
		loop

		availCurr = truncateJSONString(availCurr)
		rsAvailCurr.close
		set rsAvailCurr = nothing
	end if

	gData = "{""Languages"":[" & availLang & "], ""Currencies"":[" & availCurr & "]}"

	response.AddHeader "X-JSON", "{""EntityType"":""AvailableLangAndCurr""}"
	response.Write(gData)

end sub

sub addEntity()
	entityType = replace(Request.Form("entityType"),"'","''")
	entityVal  = replace(Request.Form("guid"),"'","''")

	select case entityType
		case "language"
			tableName = "TBLcommerce_siteLanguage"
			keyField  = "LanguageGuid"
		case "currency"
			tableName = "TBLcommerce_siteCurrency"
			keyField  = "CurrencyGuid"
	end select

	'Check if the entity exist
	isExistSQL = "SELECT * FROM " & tableName & " WHERE SiteGuid = " & siteguid & " AND " & keyField & " = " & entityVal
	set rsIsExist = conn.execute(isExistSQL)

	if not rsIsExist.eof then
		rsIsExist.close
		set rsIsExist = nothing

		response.AddHeader "X-JSON", "{""EntityType"":""ShowMessage"",""MessageText"":""The entity you are trying to add already exists.""}"
		gData = "{""Error"":""1""}"
	else
		addEntitySQL = "INSERT INTO " & tableName & " (" & keyField & ", SiteGuid, Standard) VALUES(" & entityVal & ", " & siteguid & ", 0)"

		connWrite.execute(addEntitySQL)

		response.AddHeader "X-JSON", "{""EntityType"":""UpdateEntityTable""}"
		gData = "{""Success"":""1""}"		
	end if

	response.Write(gData)

end sub

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