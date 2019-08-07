<%
'Janus d, 02-01-2009: Jeg har tilrettet koden så den benytter Application variabler til at cache data i 4 minutter. Den gamle metode hedder nu getUncachedVariableText og bør ikke kaldes medmindre man har en specefik grund til ikke at bruge caching.
function getVariableText(siteGuid, languageGuid, label)
    if isDate(Application("varaible_Cache_Date_" & siteGuid & "_" & languageGuid & "_" & label)) then
        if cdate(Application("varaible_Cache_Date_" & siteGuid & "_" & languageGuid & "_" & label))>now() then
            getVariableText = Application("varaible_Cache_" & siteGuid & "_" & languageGuid & "_" & label)&""
            
        else
            tempResult = getUncachedVariableText(siteGuid, languageGuid, label)
            Application("varaible_Cache_Date_" & siteGuid & "_" & languageGuid & "_" & label) = dateAdd("n", 4, now())
            Application("varaible_Cache_" & siteGuid & "_" & languageGuid & "_" & label) = tempResult
            getVariableText  = tempResult
        end if
    else
        tempResult = getUncachedVariableText(siteGuid, languageGuid, label)
        Application("varaible_Cache_Date_" & siteGuid & "_" & languageGuid & "_" & label) = dateAdd("n", 4, now())
        Application("varaible_Cache_" & siteGuid & "_" & languageGuid & "_" & label) = tempResult
        getVariableText  = tempResult
    end if
end function

function getUncachedVariableText(siteGuid, languageGuid, label)
    siteGuid = int(siteGuid)
    languageGuid = int(languageGuid)
    label = replace(label, "'", "''")
    sqlStr = "SELECT [value] FROM TblSite_VariableValue join TblSite_Variable on TblSite_VariableValue.VariableGuid = TblSite_Variable.Guid where TblSite_VariableValue.siteGuid = " & siteGuid & " and TblSite_VariableValue.languageGuid = " & languageGuid & " and TblSite_Variable.Label = '" & label & "'"
    set rsVariable = conn.execute(sqlStr)
    if(rsVariable.eof) then
        getUncachedVariableText  = ""
    else
        getUncachedVariableText = rsVariable(0)
    end if
end function
%>