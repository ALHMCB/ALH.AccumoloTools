<%
function ensureAddedToNewsLetter(strFirstName,strLastName,strFullName,strEmail,bolMultiRelation,intSubgroupGuid,intSiteGuid)
	strFirstName = replace(strFirstName,"'","''")
	strLastName = replace(strLastName,"'","''")
	strFullName = replace(strFullName,"'","''")
	strEmail = replace(strEmail,"'","''")	
	intSubgroupGuid = int(intSubgroupGuid)
	intSiteGuid = int(intSiteGuid)
	strUserString = bolUserExists(intSiteGuid,intSubgroupGuid,strEmail)
	existingUserGuid = split(strUserString,",")(0)
	bolUserInRightGroup = split(strUserString,",")(1)
	
	if not bolMultiRelation then	
		if existingUserGuid=0 then
			strSQL = "SET NOCOUNT ON Insert into TBLsite_user (SiteUserName, SiteUserFirstname, SiteUserLastname, SiteUserEmail, SubGroupGuid, SiteGuid) Values ('"& strFullName &"','"& strFirstName &"','"& strLastName &"','"&strEmail&"',"& intSubgroupGuid&", "&intSiteGuid&") SELECT SCOPE_IDENTITY() as LastID"
			set rsLastID = connWrite.execute(strSQL)
			if not rsLastID.eof then
				intLastID = int(rsLastID("LastID")&"")
			end if
			set rsLastID = nothing			
			ensureAddedToNewsLetter = intLastID		
		elseif existingUserGuid>0 and not bolUserInRightGroup then
			strSQL = "UPDATE TBLsite_user SET SubGroupGuid="& intSubgroupGuid&" WHERE (SiteUserGuid = "&existingUserGuid&") AND (SiteGuid = "&intSiteGuid&")"
			connWrite.execute (strSQL)	
			ensureAddedToNewsLetter = existingUserGuid			
		else
			ensureAddedToNewsLetter = existingUserGuid
		end if
	else
		if existingUserGuid=0 then
			strSQL = "SET NOCOUNT ON Insert into TBLsite_user (SiteUserName, SiteUserFirstname, SiteUserLastname, SiteUserEmail, SiteGuid) Values ('"& strFullName &"','"& strFirstName &"','"& strLastName &"','"&strEmail&"', "&intSiteGuid&") SELECT SCOPE_IDENTITY() as LastID"
			set rsLastID = connWrite.execute(strSQL)
			if not rsLastID.eof then
				intLastID = int(rsLastID("LastID")&"")
			end if
			set rsLastID = nothing
			strSQL = "INSERT INTO TBLsite_User_Subgroup (siteUserGuid, subGroupGuid) VALUES (" & intLastID & ", " &intSubgroupGuid& ")"
			connWrite.execute(strSQL)
			ensureAddedToNewsLetter = intLastID
		elseif existingUserGuid>0 and not bolUserInRightGroup then
			strSQL = "SELECT siteUserGuid FROM TBLsite_User_Subgroup WHERE (siteUserGuid = "&existingUserGuid&") AND (subGroupGuid = "&intSubgroupGuid&")"
			set rsUsers = conn.execute(strSQL)		
			if rsUsers.eof then
				strSQL = "INSERT INTO TBLsite_User_Subgroup (siteUserGuid, subGroupGuid) VALUES (" & existingUserGuid & ", " &intSubgroupGuid& ")"
				connWrite.execute(strSQL)
			end if
			set rsUsers = nothing
			ensureAddedToNewsLetter = existingUserGuid
		end if
	end if
end function

function bolUserExists(SiteGuid,SubgroupGuid,Email)
	userAddedToMailGroup = ""
	userGuid = ""
	sql="SELECT siteUserGuid FROM TBLSite_User WHERE (SiteGuid="& SiteGuid &") AND SiteUserEmail='" & Email &"'"
	set RS = conn.execute (sql)
	
	if RS.eof then 
		userGuid = 0
	else
		userGuid = int(RS("siteUserGuid")&"")
		
		strSQL = "SELECT SiteUserGuid FROM TBLsite_User_Subgroup WHERE siteUserGuid="& userGuid &" AND subGroupGuid="& SubgroupGuid
		set RS2 = conn.execute (strSQL)
		
		if RS2.eof then
			userAddedToMailGroup = 0
		else
			userAddedToMailGroup = 1
		end if
	end if
	set RS = nothing
	set RS2 = nothing
	if userAddedToMailGroup&""="" then
		userAddedToMailGroup = 0
	end if
	bolUserExists = userGuid&","&userAddedToMailGroup
end function

function getAndSetSEOForArticle(byval ArticleGuid)
	ArticleGuid = int(ArticleGuid)
	strGetSEO = "SELECT ArticleName, PageTitle, MetaDescription FROM TBLsite_articles WHERE (ArticleGuid = "&ArticleGuid&")"
	set rsSEO = conn.execute(strGetSEO)
	if not rsSEO.eof then
		PageTitle = Server.HTMLEncode(rsSEO("PageTitle")&"")
		if PageTitle&""="" then
			PageTitle = Server.HTMLEncode(rsSEO("ArticleName")&"")
		end if
		MetaDescription = Server.HTMLEncode(rsSEO("MetaDescription")&"")
	end if
	set rsSEO = nothing
end function

function getAndSetSEOForFrontpage(byval FrontpageGuid)
	call getAndSetSEOForFrontpageMulti( FrontpageGuid, 0) 
end function
function getAndSetSEOForFrontpageSubgroup(byval FrontpageSubgroupGuid)
	call getAndSetSEOForFrontpageMulti( 0, FrontpageSubgroupGuid) 
end function

function getAndSetSEOForFrontpageMulti(byval FrontpageGuid, byval FrontpageSubgroupGuid)
	if FrontpageGuid&"" <> "0" and isnumeric(FrontpageGuid&"") then
		FrontpageGuid = int(FrontpageGuid)
		strGetSEO = "SELECT fpHeadline, PageTitle, MetaDescription FROM TBLsite_frontpage WHERE (fpGuid = "&FrontpageGuid&")"
	elseif FrontpageSubgroupGuid&"" <> "0" and isnumeric(FrontpageSubgroupGuid&"") then
		FrontpageSubgroupGuid = int(FrontpageSubgroupGuid)
		strGetSEO = "SELECT fpHeadline, PageTitle, MetaDescription FROM TBLsite_frontpage WHERE (fpStart <= GETDATE()) AND (fpEnd >= GETDATE()) AND (SubgroupGuid = "&FrontpageSubgroupGuid&") OR (fpStandard = 1) AND (SubgroupGuid = "&FrontpageSubgroupGuid&") ORDER BY fpStandard DESC, fpStart DESC"
	else
		strGetSEO = "SELECT fpHeadline, PageTitle, MetaDescription FROM TBLsite_frontpage WHERE (fpStart <= GETDATE()) AND (fpEnd >= GETDATE()) AND (SiteGuid = "&msSiteGuid&") OR (fpStandard = 1) AND (SiteGuid = "&msSiteGuid&") ORDER BY fpStandard DESC, fpStart DESC"
	end if
	set rsSEO = conn.execute(strGetSEO)
	if not rsSEO.eof then
		PageTitle = Server.HTMLEncode(rsSEO("PageTitle")&"")
		if PageTitle&""="" then
			PageTitle = Server.HTMLEncode(rsSEO("fpHeadline")&"")
		end if
		MetaDescription = Server.HTMLEncode(rsSEO("MetaDescription")&"")
	end if
	set rsSEO = nothing
end function

function getSiteURL(siteGuid)
	strGetSiteUrl = "SELECT TOP (1) TBLWebsiteConfig_Value.Value " &_
	"FROM TBLWebsiteConfig_Configuration INNER JOIN " &_
	"TBLWebsiteConfig_Value ON TBLWebsiteConfig_Configuration.id = TBLWebsiteConfig_Value.ConfigurationId INNER JOIN " &_
	"TBLWebsiteConfig_Parameter ON TBLWebsiteConfig_Value.ParameterId = TBLWebsiteConfig_Parameter.Id " &_
	"WHERE (TBLWebsiteConfig_Parameter.Name = 'msSiteurl') AND (TBLWebsiteConfig_Configuration.SiteGuid = "& siteGuid &")"
	set rsSiteUrl = conn.execute(strGetSiteUrl)

	if not rsSiteUrl.eof then
		if instr(rsSiteUrl("Value")&"","http://") = 0 and instr(rsSiteUrl("Value")&"","https://") = 0 then
			getSiteURL = "http://" & rsSiteUrl("Value")
		else
			getSiteURL = rsSiteUrl("Value")
		end if
	else
		getSiteURL = ""
	end if

	set rsSiteUrl = nothing
end function

function limitLength(byval strString, byval intLength)
	if intLength>3 then
		posNextSpace = instr(intLength-3,strString," ",1)
		if posNextSpace>0 then
			strString = left(strString,posNextSpace-1) & "..."
		end if
	end if
	limitLength = strString
end function

function lefz(byval strNumber, byval digitCount)
	strDigits = ""
	for i = 1 to digitCount
		strDigits = strDigits & "0"
	next
	lefz = right(strDigits&strNumber,digitCount)
end function

function regexTest(byval text, byval pattern)
	set regEx = New RegExp
	regEx.Pattern = pattern
	regEx.IgnoreCase = false
	regexTest = regEx.Test(text) and pattern&"" <> ""
	set regEx = nothing
end function

function getRegexMatches(byval text, byval pattern, onlyFirst)
	set regEx = new RegExp
	with regEx
		.Pattern = pattern
		.IgnoreCase = false
		.Global = true
	end with
	set matches = regEx.Execute(text)	
	if onlyFirst then
		getRegexMatches = matches(0)
	else
		getRegexMatches = matches
	end if
end function

function getRegexMatch(byval text, byval pattern)
	getRegexMatch = getRegexMatches(text, pattern, true)
end function

function getSharedEntityList(siteGuid,moduleGuid)
	if isnumeric(siteGuid&"") and isnumeric(moduleGuid&"") then
		strCheckSharedMenu = "SELECT TOP (1) TBLsite_menu_siteSubModule.SharedSiteGuid" &_
			" FROM TBLsite_menu_siteSubModule INNER JOIN" &_
				" TBLsite_menu_subModule ON TBLsite_menu_siteSubModule.SubModuleGuid = TBLsite_menu_subModule.guid" &_
			" WHERE (TBLsite_menu_siteSubModule.SiteGuid = " & siteGuid & ") AND (TBLsite_menu_subModule.moduleGuid = " & moduleGuid & ") AND (TBLsite_menu_siteSubModule.SharedSiteGuid IS NOT NULL)"
		set rsCheckSharedMenu = conn.execute(strCheckSharedMenu)
		if not rsCheckSharedMenu.eof then
			strGetEntities = "SELECT EntityGuid" &_
				" FROM TBLsite_menu" &_
				" WHERE (SiteGuid = " & siteGuid & ") AND (MenuActive = 1) AND (ModuleGuid = " & moduleGuid & ")"
			set rsGetEntities = conn.execute(strGetEntities)
			first = true
			do while not rsGetEntities.eof
				if not first then
					getSharedEntityList = getSharedEntityList & ","
				end if
				getSharedEntityList = getSharedEntityList & rsGetEntities("EntityGuid")
				first = false
				rsGetEntities.movenext
			loop
		end if
	end if
end function

function safeFileStorage(siteGuid)
	if conf("msLoginType")&"" = "6" then
		SQLtemp = "SELECT safeFileStorage FROM TBLadmin_sites WHERE SiteGuid=" & siteGuid
		set adminSite = conn.execute(SQLtemp)
		safeFileStorage = (adminSite("safeFileStorage")&"" = "True")
	else
		safeFileStorage = false
	end if
end function

function showSafeFilePermission(moduleGuid,subModuleGuid,entityGuid)
	showSafeFilePermission = false
	
	if not isnumeric(session("SiteUserGuid")&"") or not isnumeric(moduleGuid&"") or not isnumeric(subModuleGuid&"") or not isnumeric(entityGuid&"") or not isnumeric(msSiteGuid&"") then
		exit function
	end if
	
	if msLoginType = 6 then
		if conf("msIntranetDisableEntitySQLCheck") then
			showSafeFilePermission = true
			exit function
		end if
		
		if msAllowContentFromOtherSites&"" <> "" then
			useSiteGuids = msAllowContentFromOtherSites
		else
			useSiteGuids = msSiteGuid
		end if
		
		if session("directSubjectGuids")&"" <> "" then
			strGetSubjectOrFilter = " OR SubjectGuid IN ("&session("directSubjectGuids")&") OR UnderSubjectGuid IN ("&session("directSubjectGuids")&") "
		else
			strGetSubjectOrFilter = ""
		end if		
		
		'getSqlWhereForLoginType6 = " AND ((SELECT COUNT(toEntityGuid) AS count1 FROM TBLSite_relation WHERE ((fromModuleGuid = 17) AND (fromSubModuleGuid = 1) AND (fromEntityGuid IN (" & fromEntityGuidValue & ")) OR (fromModuleGuid = 31) AND (fromSubModuleGuid = 0) AND (fromEntityGuid IN (SELECT SubjectGuid FROM TBLsite_Subject WHERE (SubGroupGuid IN (" & fromEntityGuidValue & ")))))" & strExtraSql & " AND (toModuleGuid = " & toModuleGuidValue & ") AND (toSubModuleGuid = " & toSubModuleGuidValue & ") AND (toEntityGuid = " & toEntityGuidName & ")) > 0)"
		
		
		'strPermission = "SELECT TOP (1) toEntityGuid FROM TBLSite_relation WHERE ((fromModuleGuid = 17) AND (fromSubModuleGuid = 1) AND (fromEntityGuid IN (" & session("subGroupGuids") & ")) OR (fromModuleGuid = 31) AND (fromSubModuleGuid = 0) AND (fromEntityGuid IN (SELECT SubjectGuid FROM TBLsite_Subject WHERE (SubGroupGuid IN (" & session("subGroupGuids") & "))))) AND (toModuleGuid = " & moduleGuid & ") AND (toSubModuleGuid = " & subModuleGuid & ") AND (toEntityGuid = " & entityGuid & ") AND (fromSiteGuid IN ("&useSiteGuids&")) AND (toSiteGuid IN ("&useSiteGuids&"))"
        if conf("msIntranetSafeFileStorageIPCheckColumn")&"" <> "" then
			safeFileStorageIPCheck = true
			sqlExtraColumn = ", TBLsite_group_sub." & conf("msIntranetSafeFileStorageIPCheckColumn") & " AS IPWhiteList"
		end if
		strPermission = "SELECT TBLSite_relation.toEntityGuid" & sqlExtraColumn &_
			" FROM TBLSite_relation INNER JOIN" &_
				" TBLsite_group_sub ON CASE TBLSite_relation.fromModuleGuid WHEN 17 THEN TBLSite_relation.fromEntityGuid ELSE" &_
					" (SELECT TOP (1) SubGroupGuid" &_
					" FROM TBLsite_Subject" &_
					" WHERE (SubjectGuid = TBLSite_relation.fromEntityGuid)) END = TBLsite_group_sub.SubGroupGuid" &_
			" WHERE ((fromModuleGuid = 17) AND (fromSubModuleGuid = 1) AND (fromEntityGuid IN (" & session("subGroupGuids") & ")) OR (fromModuleGuid = 31) AND (fromSubModuleGuid = 0) AND (fromEntityGuid IN (SELECT SubjectGuid FROM TBLsite_Subject WHERE (SubGroupGuid IN (" & session("subGroupGuids") & ") "&strGetSubjectOrFilter&")))) AND (toModuleGuid = " & moduleGuid & ") AND (toSubModuleGuid = " & subModuleGuid & ") AND (toEntityGuid = " & entityGuid & ") AND (fromSiteGuid IN ("&useSiteGuids&")) AND (toSiteGuid IN ("&useSiteGuids&"))"
        set rsPermission = conn.execute(strPermission)
		'response.Write("strPermission:<br>"&strPermission&"<br><br>")
		if not rsPermission.eof then
			if not safeFileStorageIPCheck then
				showSafeFilePermission = true
				'response.Write("showSafeFilePermission 1: "&showSafeFilePermission&"<br><br>")
			else
				remoteIP = request.ServerVariables("REMOTE_ADDR")
				do while not rsPermission.eof
					IPWhiteList = rsPermission("IPWhiteList")
					if IPWhiteList&"" = "" then
						showSafeFilePermission = true
					elseif instr(","&replace(IPWhiteList&""," ","")&",", ","&remoteIP&",") > 0  then
						showSafeFilePermission = true
					end if
					'response.Write("showSafeFilePermission 2: "&showSafeFilePermission&"<br>")
					'response.Write("IPWhiteList: "&IPWhiteList&"<br><br>")
					rsPermission.movenext
				loop
			end if
		elseif not safeFileStorageIPCheck then
			strEditPermission = "SELECT toEntityGuid" &_
				" FROM TBLSite_relation" &_
				" WHERE (fromModuleGuid = 21) AND (fromSubModuleGuid = 1) AND (fromSiteGuid = " & msSiteGuid & ") AND (toModuleGuid = " & moduleGuid & ") AND (toSubModuleGuid = " & subModuleGuid & ") AND (toEntityGuid = " & entityGuid & ") AND (toSiteGuid IN (" & useSiteGuids & ")) AND (fromEntityGuid = " & session("SiteUserGuid") & ") AND (relationStartDate IS NULL OR relationStartDate <= GETDATE()) AND (relationEndDate IS NULL OR relationEndDate >= GETDATE())"
			set rsEditPermission = conn.execute(strEditPermission)
			showSafeFilePermission = not rsEditPermission.eof
		end if
	end if
	'response.Write("showSafeFilePermission endelig: "&showSafeFilePermission&"<br><br>")
	'response.End()
end function
function getLogicalName(moduleGuid, subModuleGuid, EntityGuid)
	tmpResult = ""
	strGetLN = "SELECT TOP (1) LogicalName FROM TBLSite_Module_Entity_LogicalName WHERE (SiteGuid = "&msSiteGuid&") AND (ModuleGuid = "&moduleGuid&") AND (SubModuleGuid = "&subModuleGuid&") AND (EntityGuid = "&EntityGuid&") AND (LanguageGuid IS NULL OR LanguageGuid = "&msLanguage&") ORDER BY LanguageGuid"
	set rsLogicalName = conn.execute(strGetLN)
	if not rsLogicalName.eof then
		tmpResult = rsLogicalName("LogicalName")
	end if
	rsLogicalName.close
	set rsLogicalName = nothing
	getLogicalName = tmpResult
end function
function getCommerceLanguageGuid(websiteLanguageGuid)
	getCommerceLanguageGuid = websiteLanguageGuid
end function

function getSqlForCalendarSearch(byval calendarSubgroupGuid, byval calendarFromDate, byval calendarToDate, byval siteGuid, byval language)
	' C is getting all the events where date and relations are correct.
	' C is then used do filter the events by calendarGuid so only one event of a repeated event is selected for each day
	getSqlForCalendarSearch = "WITH C AS" &_
	" (" &_
	" SELECT DATEADD(dd, DATEDIFF(dd, 0, CalenderDate), 0) AS CalenderDate2, COALESCE (CalendarParentGuid, CalenderGuid) AS masterCalenderGuid" &_
	" FROM TBLsite_calender" &_
	" WHERE (SiteGuid IN (" & siteGuid & ")) AND (CalenderDate >= CONVERT(DATETIME, '" & year(calendarFromDate) & "-" & lefz(month(calendarFromDate),2) & "-" & lefz(day(calendarFromDate),2) & " 00:00:00', 102)) AND (CalenderDate <= CONVERT(DATETIME, '" & year(calendarToDate) & "-" & lefz(month(calendarToDate),2) & "-" & lefz(day(calendarToDate),2) & " 23:59:59', 102)) AND (COALESCE (CalendarParentGuid, CalenderGuid) IN" &_
		" (SELECT TBLsite_portal_relation.entityGuid" &_
		" FROM TBLsite_portal_relation INNER JOIN" &_
			" TBLsite_Subject ON TBLsite_portal_relation.subjectGuid = TBLsite_Subject.SubjectGuid" &_
		" WHERE (TBLsite_portal_relation.siteGuid IN (" & siteGuid & ")) AND (TBLsite_portal_relation.moduleGuid = 28) AND (TBLsite_Subject.SubGroupGuid IN (" & calendarSubgroupGuid & "))" &_
		"))" &_	
	" )" &_
    " SELECT CalenderGuid, SiteGuid, CalenderDate, (DATEADD(dd, DATEDIFF(dd, 0, CalenderDate), 0)) AS CalenderDate2, CalenderEnddate, CalenderName, CalenderText, CalenderTicketCount, CalenderTime, CalendarParentGuid, SubGroupGuid, picturePath, pictureName, logicalName, COALESCE (CalendarParentGuid, CalenderGuid) AS masterCalenderGuid" &_
	" FROM TBLsite_calender AS TBLsite_calender_2" &_
	" WHERE (SiteGuid IN (" & siteGuid & ")) AND (CalenderGuid IN" &_
		" (SELECT" &_
			" (SELECT TOP (1) CalenderGuid" &_
			" FROM TBLsite_calender AS TBLsite_calender_1" &_
			" WHERE (COALESCE (CalendarParentGuid, CalenderGuid) = C_1.masterCalenderGuid) AND (DATEADD(dd, DATEDIFF(dd, 0, CalenderDate), 0) = C_1.CalenderDate2)" &_
			" ORDER BY CalenderDate) AS CalenderGuid" &_
		" FROM C AS C_1" &_
		" GROUP BY CalenderDate2, masterCalenderGuid)" &_
		" )" &_
	" ORDER BY CalenderDate"
end function

function getSqlForCalendarSearch(byval calendarSubgroupGuid, byval calendarFromDate, byval calendarToDate, byval siteGuid, byval language)
 ' C is getting all the events where date and relations are correct.
 ' C is then used do filter the events by calendarGuid so only one event of a repeated event is selected for each day
 subjectSearchfield = "SubGroupGuid"
 if conf("msSearchCalendarBySubjectGuid") then
	subjectSearchfield = "subjectGuid"
 end if

 getSqlForCalendarSearch = "WITH C AS" &_
 " (" &_
 " SELECT DATEADD(dd, DATEDIFF(dd, 0, CalenderDate), 0) AS CalenderDate2, COALESCE (CalendarParentGuid, CalenderGuid) AS masterCalenderGuid" &_
 " FROM TBLsite_calender" &_
 " WHERE (SiteGuid IN (" & siteGuid & ")) AND (CalenderDate >= CONVERT(DATETIME, '" & year(calendarFromDate) & "-" & lefz(month(calendarFromDate),2) & "-" & lefz(day(calendarFromDate),2) & " 00:00:00', 102)) AND (CalenderDate <= CONVERT(DATETIME, '" & year(calendarToDate) & "-" & lefz(month(calendarToDate),2) & "-" & lefz(day(calendarToDate),2) & " 23:59:59', 102)) AND (COALESCE (CalendarParentGuid, CalenderGuid) IN" &_
  " (SELECT TBLsite_portal_relation.entityGuid" &_
  " FROM TBLsite_portal_relation INNER JOIN" &_
   " TBLsite_Subject ON TBLsite_portal_relation.subjectGuid = TBLsite_Subject.SubjectGuid" &_
  " WHERE (TBLsite_portal_relation.siteGuid IN (" & siteGuid & ")) AND (TBLsite_portal_relation.moduleGuid = 28) AND (TBLsite_Subject."&subjectSearchfield&" IN (" & calendarSubgroupGuid & "))" &_
  "))" &_ 
 " )" &_
    " SELECT TBLsite_calender_2.CalenderGuid, TBLsite_calender_2.SiteGuid, TBLsite_calender_2.CalenderDate,"&_
  " DATEADD(dd, DATEDIFF(dd, 0, TBLsite_calender_2.CalenderDate), 0)"&_
     " AS CalenderDate2, TBLsite_calender_2.CalenderEnddate, TBLsite_calender_2.CalenderName, TBLsite_calender_2.CalenderText,"&_
  " TBLsite_calender_2.CalenderTicketCount, TBLsite_calender_2.CalenderTime, TBLsite_calender_2.CalendarParentGuid, TBLsite_calender_2.SubGroupGuid,"&_
        " TBLsite_calender_2.picturePath, TBLsite_calender_2.pictureName, TBLsite_calender_2.logicalName, COALESCE (TBLsite_calender_2.CalendarParentGuid,"&_
        " TBLsite_calender_2.CalenderGuid) AS masterCalenderGuid"&_

 " FROM     TBLsite_calender AS TBLsite_calender_2 INNER JOIN"&_
    "   TBLsite_calenderText ON COALESCE (TBLsite_calender_2.CalendarParentGuid, TBLsite_calender_2.CalenderGuid) = TBLsite_calenderText.CalenderGuid"&_
    " WHERE  (TBLsite_calender_2.SiteGuid IN ("&siteGuid&")) AND (TBLsite_calender_2.CalenderGuid IN"&_
   		" (SELECT (SELECT TOP (1) CalenderGuid"&_
    		" FROM      TBLsite_calender AS TBLsite_calender_1"&_
    		" WHERE   (COALESCE (CalendarParentGuid, CalenderGuid) = C_1.masterCalenderGuid) AND (DATEADD(dd, DATEDIFF(dd, 0, CalenderDate), 0)"&_ 
    		"  = C_1.CalenderDate2)"&_
    		" ORDER BY CalenderDate) AS CalenderGuid"&_
    	" FROM      C AS C_1"&_
    " GROUP BY CalenderDate2, masterCalenderGuid)) AND (TBLsite_calenderText.LanguageGuid = "&language&") AND (TBLsite_calenderText.CalenderText LIKE '% %')"&_
    " ORDER BY TBLsite_calender_2.CalenderDate"
end function


function getTicketTypeName(ticketTypeGuid)
	if isnumeric(ticketTypeGuid&"") then
		if not calendarMultipleLanguage(msSiteGuid) then
			strGetName = "SELECT TOP (1) TicketType FROM TBLsite_calender_ticketType WHERE (Guid = " & ticketTypeGuid & ")"
		else
			strGetName="SELECT TOP (1) TicketType FROM TBLSite_calender_ticketType_Text WHERE (ticketTypeGuid = " & ticketTypeGuid & ") AND (languageGuid = " & msLanguage & ")"
		end if
		set rsGetName = conn.execute(strGetName)
		if not rsGetName.eof then
			getTicketTypeName = rsGetName("TicketType")
		end if
	end if
end function

function calendarMultipleLanguage(byval siteGuid)
	strGetModuleParams = "SELECT TOP (1) TBLadmin_site_module_parameters.ModuleParameterGuid, TBLadmin_site_module_parameters.SiteModuleParameterValue FROM TBLadmin_site_module INNER JOIN TBLadmin_site_module_parameters ON TBLadmin_site_module.SiteModuleGuid = TBLadmin_site_module_parameters.SiteModuleGuid WHERE (((TBLadmin_site_module.SiteGuid)="&siteGuid&") AND ((TBLadmin_site_module.ModuleGuid)=28)) AND (TBLadmin_site_module_parameters.ModuleParameterGuid = 1952)"
	Set rsGetModuleParams = Conn.Execute(strGetModuleParams)
	calendarMultipleLanguage = false
	if not rsGetModuleParams.eof then
		calendarMultipleLanguage = rsGetModuleParams("SiteModuleParameterValue")&"" = "1"
	end if
end function

function ceil(x)
	dim temp
 
	temp = Round(x)
 
	if temp < x then
		temp = temp + 1
	end if
 
	ceil = temp
end function
function floor(x)
	dim temp

	temp = Round(x)

	if temp > x then
	temp = temp - 1
	end if

	floor = temp
end function

function getMenuItem(xmlNodeSingle)

	redim arrMenu(6)
	arrMenu(0) = Server.HTMLEncode(xmlNodeSingle.selectSingleNode("./Name").Text&"") 'MenuName
	arrMenu(1) = Server.HTMLEncode(xmlNodeSingle.selectSingleNode("./Link").Text&"") 'MenuLink
	set arrMenu(2) = xmlNodeSingle.selectNodes("./SubMenus/SiteMenuLight") 'SubMenu
	arrMenu(3) = xmlNodeSingle.selectSingleNode("./StackMenu").Text&""="true" 'StackMenu
	arrMenu(4) = false 'First in stack
	arrMenu(5) = false 'Last in stack
	arrMenu(6) = int(xmlNodeSingle.selectSingleNode("./Guid").Text) 'MenuGuid
	getMenuItem = arrMenu

end function

function getMenuList(xmlNode)

	redim arrMenuList(xmlNode.length-1)
	
	for i = 0 to xmlNode.length-1
		arrMenuList(i) = getMenuItem(xmlNode(i))
	next

	for i = 0 to xmlNode.length-1
		if i = 0 then
			'First item, if stackMenu then it's also the first stacked
			if arrMenuList(i)(3) then
				arrMenuList(i)(4) = true
			end if
		elseif i = xmlNode.length-1 then
			'Last item, if stackMenu then it's also the last stacked
			if arrMenuList(i)(3) then
				arrMenuList(i)(5) = true
			end if
			if not arrMenuList(i-1)(3) then
				'The previous item wasn't stacked so this is the first stacked
				arrMenuList(i)(4) = true
			end if
		else
			if arrMenuList(i)(3) then
				'If this is stacked and not first or last
				if not arrMenuList(i+1)(3) then
					'The next menu item is not stackMenu so this is last stacked
					arrMenuList(i)(5) = true
				end if
				if not arrMenuList(i-1)(3) then
					'The previous item wasn't stacked so this is the first stacked
					arrMenuList(i)(4) = true
				end if
			end if
		end if
	next

	getMenuList = arrMenuList
end function

function isExternalLink(val)	
	internalDomains = conf("msInternalSites")

	if internalDomains&""<>"" then
		arrInternalDomains = Split(internalDomains&"", ",")
		for each x in arrInternalDomains
			if inStr(lcase(val),lcase(x)) > 0 and x&"" <> "" then
				isExternalLink = false
				exit function
			end if
		next
	end if
	if instr(1,val,"http")>0 then
		isExternalLink = true
	else
		isExternalLink = false
	end if
end function

function csvContainValue(byval csv, byval value)
	csvContainValue = instr(replace(","&csv&","," ",""), replace(","&value&","," ","")) > 0 and isCsv(csv) and trim(value&"") <> ""
end function

function isCsv(str)
	set regEx = New RegExp
	regEx.Pattern = "^[^,]+([\,]{1}\s?[^,]+)*$"
	isCsv = regEx.Test(str&"")
	set regEx = nothing
end function

function isNumberCsv(str)
	set regEx = New RegExp
	regEx.Pattern = "^(\d+){1}((,( )?\d+)+)?$"
	isNumberCsv = regEx.Test(str&"")
	set regEx = nothing
end function

function addModuleEntityToMail(byval persitsMailSender, byval moduleGuid, byval entityGuid)
	' persitsMailSender is optional. If ont provided, the function will return a new object
	if typename(persitsMailSender)&"" = "IMailSender" then
		set FormMail = persitsMailSender
	else
		set FormMail = Server.CreateObject("Persits.MailSender")
	end if
	
	' Styling
	mHeaderFontSize = 21
	mHeaderMarginBottom = 16
	
	mContentMarginBottom = 15
	mContentHeaderMarginBottom = 5
	
	mContentFontSize = 12
	mContentLineHeight = 16
	mContentFontFamily = "Arial, Helvetica, sans-serif"
	mContentFileMarginTop = 7
	
	
	if msAllowContentFromOtherSites&"" <> "" then
		siteGuids = msAllowContentFromOtherSites
	else
		siteGuids = msSiteGuid
	end if
	
	if moduleGuid = 1 then
		rsGetNews = conn.execute("SELECT NewsGuid, NewsManchet, NewsText, NewsHeadline, NewsStartDate, PicturePath, PictureName, PictureLink, PictureText, fileName, filePath, fileSize, SiteGuid" &_
			" FROM TBLsite_news" &_
			" WHERE (NewsGuid = " & entityGuid & ") AND (SiteGuid IN (" & siteGuids & "))")
		if not rsGetNews.eof then
			mailBody = "<div style=""font-size:" & mHeaderFontSize & "px; margin-bottom:" & mHeaderMarginBottom & "px;"">" & server.HTMLEncode(rsGetNews("NewsHeadline")&"") & "</div>"
			
			
			
			
			' Start of each paragraph
				mailBody = mailBody &_
					"<div style=""margin-bottom:" & mContentMarginBottom & "px;"" font-size:" & mContentFontSize & "px; line-height:" & mContentLineHeight & "px; font-family:" & mContentFontFamily & ";>"
			
			
			
			
			if rsGetNews("PictureName")&"" <> "" then
				
			end if
			
		end if
		
	elseif moduleGuid = 25 then
		set rsGetArticle = conn.execute("SELECT TOP (1) ArticleGuid, ArticleName, SiteGuid" &_
			" FROM TBLsite_articles" &_
			" WHERE (ArticleGuid = " & entityGuid & ") AND (SiteGuid IN (" & siteGuids & "))")
		if not rsGetArticle.eof then
			mailBody = "<div style=""font-size:" & mHeaderFontSize & "px; margin-bottom:" & mHeaderMarginBottom & "px;"">" & server.HTMLEncode(rsGetArticle("ArticleName")&"") & "</div>"
			
			set rsGetParagraphs = conn.execute("SELECT ParagraphGuid, ParagraphHeadline, ParagraphText, pictureName, picturePath, fileName, filePath, fileSize, layoutGuid" &_
				" FROM TBLsite_article_paragraph" &_
				" WHERE (ArticleGuid = " & rsGetArticle("ArticleGuid") & ") AND (SiteGuid IN (" & siteGuids & ")) AND (ParagraphActive = 1)" &_
				" ORDER BY ParagraphOrder")
			do while not rsGetParagraphs.eof
				if rsGetParagraphs("layoutGuid")&"" = "1" then
					picStyle = "float:left; padding:0px 10px 5px 0px;"
					picSize = "small"
				elseif rsGetParagraphs("layoutGuid")&"" = "2" then
					picStyle = "float:left; padding:0px 10px 5px 0px;"
					picSize = "medium"
				elseif rsGetParagraphs("layoutGuid")&"" = "3" then
					picStyle = "float:right; padding:0px 0px 5px 10px;"
					picSize = "small"
				elseif rsGetParagraphs("layoutGuid")&"" = "4" then
					picStyle = "float:right; padding:0px 0px 5px 10px;"
					picSize = "medium"
				elseif rsGetParagraphs("layoutGuid")&"" = "5" then
					picStyle = "padding:0px 0px 10px 0px;"
					picSize = "large"
				elseif rsGetParagraphs("layoutGuid")&"" = "" or rsGetParagraphs("layoutGuid")&"" = "0" then
					picStyle = "float:left; padding:0px 10px 5px 0px;"
					picSize = "medium"
				end if
				
				' Start of each paragraph
				mailBody = mailBody &_
					"<div style=""margin-bottom:" & mContentMarginBottom & "px;"" font-size:" & mContentFontSize & "px; line-height:" & mContentLineHeight & "px; font-family:" & mContentFontFamily & ";>"
				
				' Headline above picture
				if rsGetParagraphs("ParagraphHeadline")&"" <> "" and not conf("msParagraphHeadlineUnderPicture") then
					mailBody = mailBody &_
						"<div style=""margin-bottom:" & mContentHeaderMarginBottom & "px; font-weight:bold;"">" &_
							server.HTMLEncode(rsGetParagraphs("ParagraphHeadline")&"") &_
						"</div>"
				end if
				
				' Picture
				if rsGetParagraphs("picturePath")&"" <> "" and rsGetParagraphs("pictureName")&"" <> "" then
					cid = "p-" & rsGetParagraphs("ParagraphGuid")
					strImageUri = application("UploadPath_single") & msSiteGuid & "\" & moduleGuid & "\" & picSize & "\" & rsGetParagraphs("PictureName")
					mailBody = mailBody &_
						"<div style=""" & picStyle & """>" &_
							"<img src=""cid:" & cid & """ border=""0"" alt="""" />" &_
						"</div>"
					FormMail.AddEmbeddedImage strImageUri, cid
				end if
				
				' Headline below picture
				if rsGetParagraphs("ParagraphHeadline")&"" <> "" and conf("msParagraphHeadlineUnderPicture") then
					mailBody = mailBody &_
						"<div style=""margin-bottom:" & mContentHeaderMarginBottom & "px; font-weight:bold;"">" &_
							server.HTMLEncode(rsGetParagraphs("ParagraphHeadline")&"") &_
						"</div>"
				end if
				
				' Paragraph text
				text = rsGetParagraphs("ParagraphText")
				if not HTML.Item("25-1") then
					text = replaceBBCode(text)
					text = findLinks(text,internLinkSite, openLinkThrough, intTarget, extTarget )
					text = replace(text,vbcrLF & vbcrLF, vbcrLF & "<br style='line-height:7px'>")
					text = replace(text,vbcrLF,"<br>")
					text = replace(text,"  ","&nbsp;&nbsp;")
				end if
				mailBody = mailBody &_
					text
				
				' File
				if rsGetParagraphs("fileName")&"" <> "" and not instr(lcase(rsGetParagraphs("fileName")),".flv") > 0 then
					if safeFileStorage(rsGetArticle("SiteGuid")) then
						filePath = replace(Application("ProtectedFilePath"), "\SecureFiles\", "") & replace(rsGetParagraphs("filePath"), "/", "\") & rsGetParagraphs("fileName")
					else
						filePath = Application("UploadetFiles_Path") & replace(replace(rsGetParagraphs("filePath"),"/UploadetFiles/",""),"/","\") & rsGetParagraphs("fileName")
					end if
					FormMail.AddAttachment(filePath)
					
					mailBody = mailBody &_
						"<div style=""padding-top:" & mContentFileMarginTop & "px;"">"
				
					if not msHideFileLink then
						mailBody = mailBody &_
							"<b>" & text013 & ":</b>&nbsp;&nbsp;"
                    end if
					
					mailBody = mailBody &_
						"<i>" & rsGetParagraphs("fileName") & "</i> (Vedhæftet mailen)</div>"
				end if
                
				' End of each paragraph
				mailBody = mailBody &_
						"<div style=""clear:both;""></div>" &_
					"</div>"
				rsGetParagraphs.movenext
			loop
		end if
	end if
	
	FormMail.Body = FormMail.Body & mailBody
	set addModuleEntityToMail = FormMail
end function

function confExt(byval name, byval configId, byval siteGuid)
	if name&"" <> "" and (isnumeric(configId&"") or isnumeric(siteGuid&"")) then
		strGetValue = "SELECT TOP (1) TBLWebsiteConfig_Value.Value" &_
			" FROM TBLWebsiteConfig_Configuration LEFT OUTER JOIN" &_
				" TBLWebsiteConfig_Value INNER JOIN" &_
				" TBLWebsiteConfig_Parameter ON TBLWebsiteConfig_Value.ParameterId = TBLWebsiteConfig_Parameter.Id ON" &_
				" TBLWebsiteConfig_Configuration.id = TBLWebsiteConfig_Value.ConfigurationId" &_
			" WHERE (TBLWebsiteConfig_Parameter.Name = '" & replace(name,"'","''") & "')"
		if isnumeric(configId&"") then
			strGetValue = strGetValue & " AND (TBLWebsiteConfig_Configuration.id = " & replace(configId,"'","") & ")"
		end if
		if isnumeric(siteGuid&"") then
			strGetValue = strGetValue & " AND (TBLWebsiteConfig_Configuration.SiteGuid = " & replace(siteGuid,"'","") & ")"
		end if
		set rsGetValue = conn.execute(strGetValue)
		if not rsGetValue.eof then
			confExt = rsGetValue("Value")
		end if
	end if
end function

Function createGuid()
  Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
  tg = TypeLib.Guid
  createGuid = left(tg, len(tg)-2)
  Set TypeLib = Nothing
End Function

function loadDesignElements()
	'Variables must be dim'ed outside of this function to work, otherwise they won't be available
	strGetDesignElements = "SELECT designElementGuid, html, color, text FROM TBLsite_designElementItem WHERE (SiteGuid = " & msSiteGuid &") AND (designElementGuid IN (16,17,18,19,20,21,43,53,54,59,67,83))"
	if conf("msUseEntitySubgroupGuid") and conf("msEntitySubgroupGuidActivateStylingSystem") then
		strGetDesignElements = strGetDesignElements & " AND (SubgroupGuid = "&conf("msEntitySubgroupGuidActivateStylingSystemMainSubgroupGuid")&")"
	end if
	set rsDesignElements = conn.execute(strGetDesignElements)
	do while not rsDesignElements.EOF
		select case rsDesignElements("designElementGuid")
			case 16 'Topbillede
				designElement16 = rsDesignElements("html")
			case 17 'Baggrundsfarve
				designElement17 = rsDesignElements("color")
			case 18 'Box top farve
				designElement18 = rsDesignElements("color")
			case 19 'Farve på ramme
				designElement19 = rsDesignElements("color")
			case 20 'Farve på skriften i toppen af boksene
				designElement20 = rsDesignElements("color")
			case 21 'Farve på skriften i adresselinien
				designElement21 = rsDesignElements("color")
			case 43 'Google Analytics (Traditionel, ga.js)
				designElement43 = rsDesignElements("text")
			case 53 'HTML kode til HEAD
				designElement53 = rsDesignElements("text")
			case 54 'HTML kode til BODY
				designElement54 = rsDesignElements("text")
			case 59 'Google Analytics (Ny, asynchronous)
				designElement59 = rsDesignElements("text")
			case 67 'Grafik til baggrund
				designElement67 = rsDesignElements("html")	  
			case 83 'Google Analytics (Ny, asynchronous)
				designElement83 = rsDesignElements("text")	  
		end select
		rsDesignElements.MoveNext
	loop
	rsDesignElements.Close
	set rsDesignElements = nothing 
end function

%>
