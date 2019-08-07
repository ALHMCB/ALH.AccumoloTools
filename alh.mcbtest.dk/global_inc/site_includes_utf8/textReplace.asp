<%
Function Reggex(strString, strPattern, strReplace)
    Set RE = New RegExp
	With RE
		.Pattern = strPattern
		.Global = True
		Reggex = .Replace(strString, strReplace)
	End With
End Function

function filterReplace(strString,strFilter,strBeforeMark,strAfterMark)
	lenFilter = len(strFilter)
	iFR = 1
	done = false
	lenBeforeAfter = len(strBeforeMark)+len(strAfterMark)
	do while not done
		posOfFilter = instr(iFR,strString,strFilter,1)
		lenString = len(strString)
		if posOfFilter>0 then
			strString = left(strString,posOfFilter-1) & strBeforeMark &mid(strString,posOfFilter,lenFilter) & strAfterMark & right(strString,lenString-(posOfFilter+lenFilter)+1)
			iFR = iFR+lenFilter+lenBeforeAfter
		else
			done = true
		end if
	loop
	filterReplace = strString
end function

Function replaceBBCode(strString)
    strString = Replace(Replace(strString&"", "[quote]", "<blockquote>"), "[/quote]", "</blockquote>")
    strString = Replace(Replace(strString, "[code]", "<pre>"), "[/code]", "</pre>")

	strString = Reggex(strString, "\[url=([^\]]+) target=(.*?)\](.*?)\[\/url\]", "<a target=""$2"" href=""$1"">$3</a>")
	strString = Reggex(strString, "\[url=([^\]]+)\](.*?)\[\/url\]", "<a href=""$1"">$2</a>")
	strString = Reggex(strString, "\[url\](.*?)\[\/url\]", "<a href=""$1"">$1</a>")
	strString = Reggex(strString, "\[img\](.*?)\[\/img\]", "<img alt="""" src=""$1""/>")
	strString = Reggex(strString, "\[email=(.*?)\]([^\[]*)\[\/email\]", "<a href=""mailto:$1"" title="""">$2</a>")
	strString = Reggex(strString, "\[email](.*?)\[\/email\]", "<a href=""mailto:$1"" title="""">$1</a>")
	strString = Reggex(strString, "\[color=(.*?)\]([\s\S]*?)\[\/color\]", "<span style=""color: $1;"">$2</span>")
	strString = Reggex(strString, "\[size=(\d+)](.*)\[\/size]", "<span style=""font-size:$1px"">$2</span>")

	strString = Reggex(strString, "\[b\]", "<strong>")
	strString = Reggex(strString, "\[\/b\]", "</strong>")
	strString = Reggex(strString, "\[i\]", "<em>")
	strString = Reggex(strString, "\[\/i\]", "</em>")
	strString = Reggex(strString, "\[u\]", "<u>")
	strString = Reggex(strString, "\[\/u\]", "</u>")
	strString = Reggex(strString, "\[s\]", "<strike>")
	strString = Reggex(strString, "\[\/s\]", "</strike>")
	strString = Reggex(strString, "\[li\]([^\[]*)\[\/li\]", "<li>$1</li>")

	strString = Reggex(strString, "\[google\](.*?)\[\/google\]", "<a href=""http://www.google.com/search?q=$1"" title="""">Google $1</a>")
	strString = Reggex(strString, "\[youtube\](.*?)\[\/youtube\]", "<object width=""425"" height=""355""><param name=""movie"" value=""http://www.youtube.com/v/$1""></param><param name=""wmode"" value=""transparent""></param><embed type=""application/x-shockwave-flash"" wmode=""transparent"" width=""425"" height=""355"" src=""http://www.youtube.com/v/$1""></embed></object>")

	strString = Reggex(strString, "\[br\]", "<br />")
	strString = Reggex(strString, "\[nb\](.*?)\[\/nb\]", "<span style=""white-space: nowrap;"">$1</span>")

	replaceBBCode = strString
End Function


Function removeBBCode(strString)
    strString = Replace(Replace(strString, "[quote]", ""), "[/quote]", "")
    strString = Replace(Replace(strString, "[code]", ""), "[/code]", "")

	strString = Reggex(strString, "\[url=([^\]]+)\](.*?)\[\/url\]", "$2")
	strString = Reggex(strString, "\[url\](.*?)\[\/url\]", "$1")
	strString = Reggex(strString, "\[img\](.*?)\[\/img\]", "$1")
	strString = Reggex(strString, "\[email=(.*?)\]([^\[]*)\[\/email\]", "$2")
	strString = Reggex(strString, "\[email](.*?)\[\/email\]", "$1")
	strString = Reggex(strString, "\[color=(.*?)\]([\s\S]*?)\[\/color\]", "$2")
	strString = Reggex(strString, "\[size=(\d+)](.*)\[\/size]", "$2")

	strString = Reggex(strString, "\[b\]", "")
	strString = Reggex(strString, "\[\/b\]", "")
	strString = Reggex(strString, "\[i\]", "")
	strString = Reggex(strString, "\[\/i\]", "")
	strString = Reggex(strString, "\[u\]", "")
	strString = Reggex(strString, "\[\/u\]", "")
	strString = Reggex(strString, "\[s\]", "")
	strString = Reggex(strString, "\[\/s\]", "")
	strString = Reggex(strString, "\[li\]([^\[]*)\[\/li\]", "$1")
	strString = Reggex(strString, "\[br\]", " ")
	strString = Reggex(strString, "\[nb\]", "")
	strString = Reggex(strString, "\[\/nb\]", "")	

	strString = Reggex(strString, "\[google\](.*?)\[\/google\]", "$1")
	strString = Reggex(strString, "\[youtube\](.*?)\[\/youtube\]", "$1")

	removeBBCode = strString
End Function

function replaceAll(text, openLinkThrough, intTarget, extTarget)
	text = Server.HTMLEncode(text&"")
	text = replaceBBCode(text&"")
	text = findLinks(text&"", "http://" & Request.ServerVariables("SERVER_NAME"), openLinkThrough, intTarget, extTarget)
	text = Replace(text&"", vbCrLf, "<br /> ")
	text = Replace(text&"", vbCr, "<br /> ")
	text = Replace(text&"", vbLf, "<br /> ")
	text = replace(text&"","  ","&nbsp;&nbsp;")
	replaceAll = text
end function

function replaceHtmlCode(text, openLinkThrough, intTarget, extTarget)
	text = replaceBBCode(text&"")
	text = findLinks(text&"", "http://" & Request.ServerVariables("SERVER_NAME"), openLinkThrough, intTarget, extTarget)
	text = Replace(text&"", vbCrLf, "<br /> ")
	text = Replace(text&"", vbCr, "<br /> ")
	text = Replace(text&"", vbLf, "<br /> ")
	text = replace(text&"","  ","&nbsp;&nbsp;")
	replaceHtmlCode = text
end function
%>