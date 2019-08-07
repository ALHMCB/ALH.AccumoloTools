<%
  function ValidateEmail(sEmail)
    dim regEx
	  set regEx = new RegExp
    
    sEmail = trim(sEmail)
    ValidateEmail = false
    dim retVal

    ' Create regular expression:
    regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,4}$" 

    ' Set pattern:
    regEx.IgnoreCase = true

    ' Set case sensitivity.
    retVal = regEx.Test(sEmail)

    ' Execute the search test.
    if not retVal then
      exit function
    end if
    if inStr(1,sEmail,"..") = 0 and inStr(1,sEmail,".@") = 0 and inStr(1,sEmail,"@.") = 0 and inStr(1,sEmail,"@") > 0 and sEmail & "" <> "1" then
	    ValidateEmail = true
    end if
  end function
%>