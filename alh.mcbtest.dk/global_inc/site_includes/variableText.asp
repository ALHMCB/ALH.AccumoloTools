<%
function getVariableText(siteGuid, languageGuid, label)
    sqlStr = "SELECT [value] FROM TblSite_VariableValue join TblSite_Variable on TblSite_VariableValue.VariableGuid = TblSite_Variable.Guid where TblSite_VariableValue.siteGuid = " & siteGuid & " and TblSite_VariableValue.languageGuid = " & languageGuid & " and TblSite_Variable.Label = '" & label & "'"
    set rsVariable = conn.execute(sqlStr)
    if(rsVariable.eof) then
        getVariableText  = ""
    else
        getVariableText = rsVariable(0)
    end if
end function
 %>