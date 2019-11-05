<%
function checkDOMFAvailability()
	sqlDOMFCheck = 	"Declare @SiteGuid int = " & siteguid & ";" &_
					"Select CASE WHEN EXISTS ( " &_
						"SELECT PageSystemGuid " &_
						"FROM TBLpage_SiteSystem " &_
						"WHERE SiteGuid = @SiteGuid AND PageSystemGuid = 1" &_
					") " &_
					"THEN Cast(1 AS BIT) " &_
					"ELSE Cast(0 AS BIT) END;"

	set rsDOMFCheck = conn.execute(sqlDOMFCheck)

	if not rsDOMFCheck.eof then
		isDOMF = Cbool(rsDOMFCheck.Fields.Item(0))
	end if
	rsDOMFCheck.close
	set rsDOMFCheck = nothing

	checkDOMFAvailability = isDOMF
end function


sub getDOMFStatus()
	sqlDOMFStatus = "Declare @SiteGuid int = " & siteguid & ";" &_
					"Select A.PageSystemGuid AS [Guid], B.Name From TBLpage_SiteSystem A INNER JOIN TBLpage_system B " &_
						"ON A.PageSystemGuid = B.Guid " &_
						"Where siteguid = @SiteGuid FOR JSON PATH;" &_
					"Select A.PageBlockGroupGuid AS [Guid], B.Name From TBLpage_siteBlockGroup A INNER JOIN TBLpage_BlockGroup B " &_
						"ON A.PageBlockGroupGuid = B.Guid " &_
						"Where SiteGuid = @SiteGuid FOR JSON PATH;" &_
					"Select A.PagePageGroupGuid AS [Guid], B.Name From TBLpage_sitePageGroup A INNER JOIN TBLpage_PageGroup B " &_
						"ON A.PagePageGroupGuid = B.Guid " &_
						"Where SiteGuid = @SiteGuid FOR JSON PATH;" &_
					"Select A.boxTemplateGroupGuid AS [Guid], B.Name From TBLSite_BoxTemplateGroupSite A INNER JOIN TBLSite_BoxTemplateGroup B " &_
						"ON A.boxTemplateGroupGuid = B.Guid " &_
						"Where SiteGuid = @SiteGuid FOR JSON PATH;" &_
					"Select A.websiteSkinGroupGuid AS [Guid], B.Name From TBLSite_WebsiteSkinGroupSite A INNER JOIN TBLSite_WebsiteSkinGroup B " &_
						"ON A.websiteSkinGroupGuid = B.Guid " &_
						"Where SiteGuid = @SiteGuid FOR JSON PATH;"

	set rsDOMFStatus = conn.Execute(sqlDOMFStatus)
	
	if not rsDOMFStatus.eof then
		xPageSystem = rsDOMFStatus.Fields.Item(0)
	else
		xPageSystem = "[]"
	end if

	set rsDOMFStatus = rsDOMFStatus.NextRecordset
	if not rsDOMFStatus.eof then
		xBlockGroup = rsDOMFStatus.Fields.Item(0)
	else
		xBlockGroup = "[]"
	end if

	set rsDOMFStatus = rsDOMFStatus.NextRecordset
	if not rsDOMFStatus.eof then
		xPageGroup = rsDOMFStatus.Fields.Item(0)
	else
		xPageGroup = "[]"
	end if

	set rsDOMFStatus = rsDOMFStatus.NextRecordset
	if not rsDOMFStatus.eof then
		xBoxTemplate = rsDOMFStatus.Fields.Item(0)
	else
		xBoxTemplate = "[]"
	end if

	set rsDOMFStatus = rsDOMFStatus.NextRecordset
	if not rsDOMFStatus.eof then
		xWebsiteSkin = rsDOMFStatus.Fields.Item(0)
	else
		xWebsiteSkin = "[]"
	end if

	rsDOMFStatus.close
	set rsDOMFStatus = nothing
	
	' Process query result
	xSettings = "{""PageSystem"":" & xPageSystem & ", ""BlockGroup"":" & xBlockGroup & ", ""PageGroup"":" & xPageGroup & ", ""BoxTemplate"":" & xBoxTemplate & ", ""WebsiteSkin"":" & xWebsiteSkin & "}"

	isDOMFSite = checkDOMFAvailability()

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

	isDOMFSite = checkDOMFAvailability()

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