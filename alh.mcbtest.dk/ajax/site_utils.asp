<%
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

%>