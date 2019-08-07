<%

function checkDOMFAvailability(byRef xPageSystem, byRef xBlockGroup, byRef xPageGroup, byRef xBoxTemplate, byRef xWebsiteSkin)
	sqlSiteSettings = "Declare @SiteGuid int = " & siteguid & ";" &_
						"Select A.*, B.Name From TBLpage_SiteSystem A INNER JOIN TBLpage_system B " &_
							"ON A.PageSystemGuid = B.Guid " &_
							"Where siteguid = @SiteGuid;" &_
						"Select A.*, B.Name From TBLpage_siteBlockGroup A INNER JOIN TBLpage_BlockGroup B " &_
							"ON A.PageBlockGroupGuid = B.Guid " &_
							"Where SiteGuid = @SiteGuid;" &_
						"Select A.*, B.Name From TBLpage_sitePageGroup A INNER JOIN TBLpage_PageGroup B " &_
							"ON A.PagePageGroupGuid = B.Guid " &_
							"Where SiteGuid = @SiteGuid;" &_
						"Select A.*, B.Name From TBLSite_BoxTemplateGroupSite A INNER JOIN TBLSite_BoxTemplateGroup B " &_
							"ON A.boxTemplateGroupGuid = B.Guid " &_
							"Where SiteGuid = @SiteGuid;" &_
						"Select A.*, B.Name From TBLSite_WebsiteSkinGroupSite A INNER JOIN TBLSite_WebsiteSkinGroup B " &_
							"ON A.websiteSkinGroupGuid = B.Guid " &_
							"Where SiteGuid = @SiteGuid;"

	set rsSiteSettings = conn.Execute(sqlSiteSettings)

	isDOMF = false

	if not rsSiteSettings.eof then
		do while not rsSiteSettings.eof
			xPageSystem = xPageSystem & "{""Guid"":""" & rsSiteSettings("PageSystemGuid") & """, ""Name"":""" & rsSiteSettings("Name") & """},"
			if Cint(rsSiteSettings("PageSystemGuid")) = 1 then
				isDOMF = true
			end if
			rsSiteSettings.MoveNext
		loop

		xPageSystem = truncateJSONString(xPageSystem)

	end if

	set rsSiteSettings = rsSiteSettings.NextRecordset
	if not rsSiteSettings.eof then
		do while not rsSiteSettings.eof
			xBlockGroup = xBlockGroup & "{""Guid"":""" & rsSiteSettings("PageBlockGroupGuid") & """, ""Name"":""" & rsSiteSettings("Name") & """},"
			rsSiteSettings.MoveNext
		loop

		xBlockGroup = truncateJSONString(xBlockGroup)

	end if

	set rsSiteSettings = rsSiteSettings.NextRecordset
	if not rsSiteSettings.eof then
		do while not rsSiteSettings.eof
			xPageGroup = xPageGroup & "{""Guid"":""" & rsSiteSettings("PagePageGroupGuid") & """, ""Name"":""" & rsSiteSettings("Name") & """},"
			rsSiteSettings.MoveNext
		loop

		xPageGroup = truncateJSONString(xPageGroup)

	end if

	set rsSiteSettings = rsSiteSettings.NextRecordset
	if not rsSiteSettings.eof then
		do while not rsSiteSettings.eof
			xBoxTemplate = xBoxTemplate & "{""Guid"":""" & rsSiteSettings("boxTemplateGroupGuid") & """, ""Name"":""" & rsSiteSettings("Name") & """},"
			rsSiteSettings.MoveNext
		loop

		xBoxTemplate = truncateJSONString(xBoxTemplate)

	end if

	set rsSiteSettings = rsSiteSettings.NextRecordset
	if not rsSiteSettings.eof then
		do while not rsSiteSettings.eof
			xWebsiteSkin = xWebsiteSkin & "{""Guid"":""" & rsSiteSettings("websiteSkinGroupGuid") & """, ""Name"":""" & rsSiteSettings("Name") & """},"
			rsSiteSettings.MoveNext
		loop

		xWebsiteSkin = truncateJSONString(xWebsiteSkin)

	end if

	rsSiteSettings.close
	set rsSiteSettings = nothing
	
	checkDOMFAvailability = isDOMF

end function

sub getDOMFStatus()
	xPageSystem		= ""
	xBlockGroup		= ""
	xPageGroup		= ""
	xBoxTemplate	= ""
	xWebsiteSkin	= ""

	isDOMFSite = checkDOMFAvailability(xPageSystem, xBlockGroup, xPageGroup, xBoxTemplate, xWebsiteSkin)

	' Process query result
	xSettings = "{""PageSystem"":[" & xPageSystem & "], ""BlockGroup"":[" & xBlockGroup & "], ""PageGroup"":[" & xPageGroup & "], ""BoxTemplate"":[" & xBoxTemplate & "], ""WebsiteSkin"":[" & xWebsiteSkin & "]}"

	message = ""
	if isDOMFSite = true then 
		message = "This is already a DOMF site. No upgrade required."
	else 
		message = "This is a legacy webshop. DOMF upgrade available."
	end if
	response.AddHeader "X-JSON", "{""EntityType"":""Domf"",""IsDOMFSite"":""" & isDOMFSite & """,""MessageText"":""" & message & """}"
	response.Write(xSettings)

end sub

sub createSPUPdateModuleParam()
	sqlCleanup = 	"IF OBJECT_ID('tempdb..#UpdateModuleParameters') IS NOT NULL " &_
						"DROP PROC #UpdateModuleParameters;"
	connWrite.execute(sqlCleanup)

	sqlSP = "CREATE PROCEDURE #UpdateModuleParameters " &_
				"@iSiteGuid int," &_
				"@iModuleGuid int," &_
				"@iParameterGuid int " &_
			"AS DECLARE " &_
					"@iSiteModuleGuid int," &_
					"@iSiteModuleParameterGuid int," &_
					"@iNewParameterValue int = 1;" &_

				"SELECT @iSiteModuleGuid = SiteModuleGuid " &_
					"FROM TBLadmin_site_module " &_
					"WHERE SiteGuid = @iSiteGuid AND ModuleGuid = @iModuleGuid;" &_

				"SELECT @iSiteModuleParameterGuid = SiteModuleParameterGuid " &_
				"FROM TBLadmin_site_module_parameters " &_
				"WHERE " &_
					"ModuleParameterGuid = @iParameterGuid " &_
					"AND SiteModuleGuid = @iSiteModuleGuid;" &_

				"IF @iSiteModuleParameterGuid IS NULL " &_
					"INSERT INTO TBLadmin_site_module_parameters " &_
						"(" &_
							"SiteModuleGuid," &_
							"ModuleParameterGuid," &_
							"SiteModuleParameterValue" &_
						") " &_
					"VALUES " &_
						"(" &_
							"@iSiteModuleGuid," &_
							"@iParameterGuid," &_
							"@iNewParameterValue" &_
						") " &_
				"ELSE " &_
					"UPDATE TBLadmin_site_module_parameters SET " &_
						"SiteModuleParameterValue = @iNewParameterValue " &_
					"WHERE " &_
						"SiteModuleParameterGuid = @iSiteModuleParameterGuid " &_
						"AND SiteModuleGuid = @iSiteModuleGuid " &_
						"AND ModuleParameterGuid = @iParameterGuid;"
	connWrite.execute(sqlSP)

end sub

sub makeDOMF()
	isNewSite = replace(Request.Form("isNewSite"),"'","''")
	isB2B = replace(Request.Form("isB2B"),"'","''")
	delOldMenu = replace(Request.Form("delOldMenu"),"'","''")
	addDefaultStockTypes = replace(Request.Form("addDefaultStockTypes"),"'","''")

	isDOMFSite = checkDOMFAvailability("", "", "", "", "")

	if isDOMFSite then 
		response.AddHeader "X-JSON", "{""EntityType"":""ShowMessage"",""MessageText"":""This site has already been upgraded to DOMF. Operation terminated.""}"
		xResponse = "{""Error"":""1""}"
	else 
		sqlMakeDOMF =	"Declare " &_
							"@SiteGuid int = " & siteguid & ", " &_
							"@IsNewSite bit = " & isNewSite & ", " &_
							"@IsB2B bit = " & isB2B & ", " &_
							"@RemoveOldStyleMenu bit = " & delOldMenu & ", " &_
							"@AddDefaultStockTypes bit = " & addDefaultStockTypes & ",	" &_
							
							"@PageSystemGuidResponsive int = 1, " &_
							"@PageSystemGuidMobile int = 3, " &_
							"@PageSystemGuidEmail int = 5, " &_
							"@Block_Group_MobileStandard int = 2, " &_
							"@Block_Group_MobileAccount int = 3, " &_
							"@Block_Group_WebshopDOMF int = 65, " &_
							"@Block_Group_WebshopDOMFDev int = 66, " &_
							"@Block_Group_AccumoloB2C int = 96, " &_
							"@Block_Group_AccumoloB2B int = 80, " &_
							"@Block_Group_Email int = 69, " &_
							"@Page_Group_MobileStandard int = 4, " &_
							"@Page_Group_MobileAccount int = 5, " &_
							"@Page_Group_WebshopDOMF int = 37, " &_
							"@Page_Group_AccumoloB2B int = 50, " &_
							"@Page_Group_Email int = 41, " &_
							"@newBoxTemplateGroupGuid int = 51, " &_
							"@newWebsiteSkinGroupGuid int = 20, " &_
							"@migratedWebsiteSkinGroupGuid int = 19, " &_
							"@ModuleGuid int, " &_
							"@ParameterGuid int; " &_

						"BEGIN " &_
							"INSERT INTO dbo.TBLpage_SiteSystem (SiteGuid, PageSystemGuid) " &_
								"VALUES " &_
									"(@SiteGuid, @PageSystemGuidResponsive), " &_
									"(@SiteGuid, @PageSystemGuidEmail);" &_
							
							"INSERT INTO TBLpage_siteBlockGroup (SiteGuid, PageBlockGroupGuid) " &_
								"VALUES " &_
									"(@SiteGuid, @Block_Group_WebshopDOMF), " &_
									"(@SiteGuid, @Block_Group_WebshopDOMFDev), " &_
									"(@SiteGuid, @Block_Group_AccumoloB2C), " &_
									"(@SiteGuid, @Block_Group_Email);" &_

							"INSERT INTO TBLpage_sitePageGroup (SiteGuid, PagePageGroupGuid) " &_
								"VALUES " &_
									"(@SiteGuid, @Page_Group_WebshopDOMF)," &_ 
									"(@SiteGuid, @Page_Group_Email);" &_

							"BEGIN " &_
								"SET @ModuleGuid = 270;" &_
								"SET @ParameterGuid = 3357;" &_

								"EXEC #UpdateModuleParameters @SiteGuid, @ModuleGuid, @ParameterGuid;" &_
							"END;" &_

							"BEGIN " &_
								"SET @ModuleGuid = 211;" &_
								"SET @ParameterGuid = 2647;" &_

								"EXEC #UpdateModuleParameters @SiteGuid, @ModuleGuid, @ParameterGuid;" &_
							"END;" &_

							"BEGIN " &_
								"SET @ModuleGuid = 211;" &_
								"SET @ParameterGuid = 1737;" &_

								"EXEC #UpdateModuleParameters @SiteGuid, @ModuleGuid, @ParameterGuid;" &_
							"END;" &_
						"END;" &_

						"IF @IsB2B = 1 " &_
							"BEGIN " &_
								"INSERT INTO TBLpage_siteBlockGroup (SiteGuid, PageBlockGroupGuid) " &_
								"VALUES " &_
									"(@SiteGuid, @Block_Group_AccumoloB2B);" &_

								"INSERT INTO TBLpage_sitePageGroup (SiteGuid, PagePageGroupGuid) " &_
								"VALUES " &_
									"(@SiteGuid, @Page_Group_AccumoloB2B);" &_
							"END;" &_

						"IF @IsNewSite = 1 " &_
							"BEGIN " &_
								"INSERT INTO dbo.TBLpage_SiteSystem ( SiteGuid, PageSystemGuid) " &_
									"VALUES ( @SiteGuid, @PageSystemGuidMobile );" &_

								"INSERT INTO TBLpage_siteBlockGroup (SiteGuid, PageBlockGroupGuid) " &_
									"VALUES " &_
										"(@SiteGuid, @Block_Group_MobileStandard), " &_
										"(@SiteGuid, @Block_Group_MobileAccount);" &_

								"INSERT INTO TBLpage_sitePageGroup (SiteGuid, PagePageGroupGuid) " &_
									"VALUES " &_
										"(@SiteGuid, @Page_Group_MobileStandard), " &_
										"(@SiteGuid, @Page_Group_MobileAccount);" &_

								"BEGIN " &_
									"DECLARE " &_
										"@tmpBox int, " &_
										"@tmpSkin int;" &_

									"SELECT @tmpBox = boxTemplateGroupGuid " &_
										"FROM TBLSite_BoxTemplateGroupSite " &_
										"WHERE SiteGuid = @SiteGuid;" &_

									"SELECT @tmpSkin = websiteSkinGroupGuid " &_
										"FROM TBLsite_websiteSkinGroupSite " &_
										"WHERE SiteGuid = @SiteGuid;" &_

									"IF @tmpBox IS NULL " &_
										"INSERT INTO TBLSite_BoxTemplateGroupSite (SiteGuid, boxTemplateGroupGuid) " &_
										"VALUES (@SiteGuid, @newBoxTemplateGroupGuid);" &_
									"ELSE " &_
										"UPDATE TBLSite_BoxTemplateGroupSite " &_ 
											"SET boxTemplateGroupGuid = @newBoxTemplateGroupGuid " &_
										"WHERE SiteGuid = @SiteGuid;" &_

									"IF @tmpSkin IS NULL " &_
										"INSERT INTO TBLSite_WebsiteSkinGroupSite (SiteGuid, websiteSkinGroupGuid) " &_
										"VALUES (@SiteGuid, @newWebsiteSkinGroupGuid);" &_
									"ELSE " &_
										"UPDATE TBLSite_WebsiteSkinGroupSite " &_
											"SET websiteSkinGroupGuid = @newWebsiteSkinGroupGuid " &_
										"WHERE SiteGuid = @SiteGuid;" &_
								"END;" &_

								"IF @RemoveOldStyleMenu = 1 " &_
									"BEGIN " &_
										"DELETE FROM TBLSite_Menu WHERE SiteGuid = @SiteGuid AND (SubGroupGuid = 0 OR SubGroupGuid IS NULL); " &_
									"END;" &_

								"IF @AddDefaultStockTypes = 1 " &_
									"BEGIN " &_
										"INSERT INTO TBLcommerce_stock_Site_StockType (SiteGuid, StockType) values(@SiteGuid,1); " &_
										"INSERT INTO TBLcommerce_stock_Site_StockType (SiteGuid, StockType) values(@SiteGuid,3); " &_
										"INSERT INTO TBLcommerce_stock_Site_StockType (SiteGuid, StockType) values(@SiteGuid,7); " &_
									"END;" &_
							"END;" &_
						"ELSE " &_
							"BEGIN " &_
								"DECLARE @tmpMigratedSkin int;" &_

								"SELECT @tmpMigratedSkin = websiteSkinGroupGuid "&_
									"FROM TBLsite_websiteSkinGroupSite "&_
									"WHERE SiteGuid = @SiteGuid;" &_

								"IF @tmpMigratedSkin IS NULL " &_
									"INSERT INTO TBLSite_WebsiteSkinGroupSite (SiteGuid, websiteSkinGroupGuid) " &_
									"VALUES (@SiteGuid, @migratedWebsiteSkinGroupGuid);" &_
								"ELSE " &_
									"Update TBLsite_websiteSkinGroupSite " &_
										"SET websiteSkinGroupGuid = @migratedWebsiteSkinGroupGuid " &_
									"Where SiteGuid = @SiteGuid;" &_
							"END;"

		call createSPUPdateModuleParam()

		connWrite.execute(sqlMakeDOMF)

		response.AddHeader "X-JSON", "{""EntityType"":""UpdateDOMFStatus""}"
		xResponse = "{""Success"":""1""}"

	end if
	
	response.write(xResponse)

end sub

%>