<%
'Fast string concatination.
'HACK (Janus 19-04-2010): As the findlinks.asp page somtimes is included several times, the class has been rewritten to not fail.
'Original class: http://www.15seconds.com/howto/pg000929.htm
growthRate = 50'the rate at which the array grows
itemCount = 0'the number of items in the array
ReDim arr(growthRate)

Sub Append(ByVal strValue)
		If itemCount > UBound(arr) Then
			ReDim Preserve arr(UBound(arr) + growthRate)
		End If

		arr(itemCount) = strValue
		itemCount = itemCount + 1
End Sub

Function ToString() 
	ToString = Join(arr, "")
End Function

' text = teksten hvor der skal findes links
' internLinkSite = Hvis der findes et link til dette site, er det et internt link
' openLinkThrough = Eksterne links åbnes gennem denne side
' intTarget = target til interne links
' extTarget = target til eksterne links

'eksempel: findLinks(rsNews("NewsText"),"www.masterpiece.dk", "www.campingferie.dk/link.asp?link=", "", "_blank" )
function removeBadChar(str, endLink)
	Select Case Mid(str, endLink - 1, 1)
		Case ".", "!", "?", ",", ")", "("
			endLink = endLink - 1
			result = true
		Case Else
			result = false
	End Select
	removeBadChar = result
end function

function getFirstNewLine(start, str)
	dim endlineArr(3)
	endlineArr(0) = InStr(start, str, vbCrLf)
	endlineArr(1) = InStr(start, str, vbCr)
	endlineArr(2)= InStr(start, str, vbLf)
	tempEnd = 0
	for endlineCounter=0 to ubound(endlineArr)
		if endlineArr(endlineCounter)>0 and endlineArr(endlineCounter)&""<>"" then
			if endlineArr(endlineCounter)<tempEnd or tempEnd=0 then
				tempEnd = endlineArr(endlineCounter)
			end if
		end if
	next
	getFirstNewLine = tempEnd
end function

Function findLinks(strInput, internLinkSite, openLinkThrough, intTarget, extTarget)
    growthRate = 50'the rate at which the array grows
    itemCount = 0'the number of items in the array
    strInput = strInput &""' handle null inputs
	strInput = Replace(strInput, "src=""//www.youtube.com", "src=""http://www.youtube.com") ' fix to allow new youtube embed
    reDim arr(growthRate)

	Dim iCurrentLocation  ' Our current position in the input string
	Dim iLinkStart        ' Beginning position of the current link
	Dim iLinkEnd          ' Ending position of the current link
	Dim strLinkText       ' Text we're converting to a link
	Dim strOutput         ' Return string with links in it
	if openLinkThrough<>"" then
		if inStr(1,openLinkThrough,"http://")=0 then
			openLinkThrough = "http://" & openLinkThrough
		end if
	end if

	' Start at the first character in the string
	iCurrentLocation = 1
	' Look for http:// in the text from the current position to
	' the end of the string.  If we find it then we start the
	' linking process otherwise we're done because there are no
	' more http://'s in the string.
	Do While InStr(iCurrentLocation, strInput, "mms://", 1)  <> 0 OR InStr(iCurrentLocation, strInput, "http://", 1) <> 0 OR InStr(iCurrentLocation, strInput, "https://", 1) <> 0 OR InStr(iCurrentLocation, strInput, "www.", 1) <> 0 OR (Instr(iCurrentLocation, strInput, "@", 1)<>0)
		' Set the position of the beginning of the link
		httpPos = InStr(iCurrentLocation, strInput, "http://", 1)
		httpsPos = InStr(iCurrentLocation, strInput, "https://", 1)
		if (httpsPos<httpPos and httpsPos>0) or (httpsPos>httpPos and httpPos=0) then
			httpPos = httpsPos
		end if
		wwwPos = InStr(iCurrentLocation, strInput, "www.", 1)
		mailPos = Instr(iCurrentLocation, strInput, "@")
		mmsPos = InStr(iCurrentLocation, strInput, "mms://", 1)
		mailFound = false
		httpFound = false
		wwwFound = false
		mmsFound = false
		if (mailPos>0 and (mailPos<wwwPos or wwwPos=0) and (mailPos<httpPos or httpPos=0) and (mailPos<mmsPos or mmsPos=0)) then
		    mailFound = true
		elseif (wwwPos>0 and (wwwPos<mailPos or mailPos=0) and (wwwPos<httpPos or httpPos=0) and (wwwPos<mmsPos or mmsPos=0)) then
		    wwwFound = true
		elseif (httpPos>0 and (httpPos<mailPos or mailPos=0) and (httpPos<wwwPos or wwwPos=0) and (httpPos<mmsPos or mmsPos=0)) then
		    httpFound = true
		elseif (mmsPos>0 and (mmsPos<mailPos or mailPos=0) and (mmsPos<httpPos or httpPos=0) and (mmsPos<wwwPos or wwwPos=0)) then
		    mmsFound = true
		end if
		
		if mailFound then
			'Find email
			startSpace=inStrRev(strInput," ",mailPos,1)+1
			startEndLine=inStrRev(strInput,vbCrLf,mailPos,1)
			if(startEndLine>0) then
			    startEndLine = startEndLine + 2
			else
                if startEndLine=0 then
                    startEndLine = inStrRev(strInput,vbCr,mailPos,1)
                end if
                if startEndLine=0 then
                    startEndLine = inStrRev(strInput,vbLf,mailPos,1)
                end if
			    if(startEndLine>0) then
			        startEndLine = startEndLine + 1
			    end if
            end if
			if startSpace>startEndLine or startSpace=0 then
				iLinkStart = startSpace
			else
				iLinkStart = startEndLine
			end if
			'Tjek om der allerede er et link her
		  HREFTagFound = false
		  if iLinkStart>6 then
		    if lcase(mid(strInput, iLinkStart-6, 5))="href=" then
		      HREFTagFound = true
		    end if
		  end if

			endSpace = inStr(mailPos,strInput," ",1)
			'endEndLine =  inStr(mailPos,strInput,vbCrLf,1)
			endEndLine = getFirstNewLine(mailPos, strInput)
			if (endSpace>endEndLine and endEndLine>0) or endSpace=0 then
				iLinkEnd = endEndLine
			else
				iLinkEnd = endSpace
			end if
			
			
			If iLinkEnd = 0 Then iLinkEnd = Len(strInput) + 1
			' This adds to the output string all the non linked stuff
			' up to the link we're curently processing.
			Append(Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation))
			'strOutput = strOutput & Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation)
			' Get the text we're linking and store it in a variable
			'Fjerner højest 20 tegn, for at begrænse server belastning ved fejl.
			count = 0
			do while removeBadChar(strInput, iLinkEnd) and count<20
				count = count + 1
			loop
			
			strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)
			'Skriv selve linket.
			
			' Validerer om der ikke allerede findes et tag af samme karakter:
			'Skriv selve linket.
			if ( InStr( 1, LCase(strLinkText), "href=") > 0 ) OR ( InStr(1, strLinkText, """") > 0 ) OR ( InStr( 1, LCase(strLinkText), ">") > 0 ) OR ( InStr( 1, LCase(strLinkText), "<") > 0 ) then
				Append(strLinkText)
				'strOutput = strOutput &strLinkText
			else
				Append("<a href=""mailto:" & Server.URLEncode(strLinkText) & """ style=""" & linkStyle & """>" & strLinkText & "</a>")
				'strOutput = strOutput & "<a href=""mailto:" & Server.URLEncode(strLinkText) & """ style=""" & linkStyle & """>" & strLinkText & "</a>"	
			end if
			
		else
			'Find web link
			if wwwFound then
				iLinkStart = InStr(iCurrentLocation, strInput, "www.", 1)
				tempType = "www"
			elseif httpFound then
				tempType = "http"
				iLinkStart = InStr(iCurrentLocation, strInput, "http://", 1)
				TempHttpsPos = InStr(iCurrentLocation, strInput, "https://", 1)
				if (TempHttpsPos<iLinkStart and TempHttpsPos>0) OR (TempHttpsPos>iLinkStart and iLinkStart=0) then
					iLinkStart = TempHttpsPos
				end if
			else
				tempType = "mms"
				iLinkStart = InStr(iCurrentLocation, strInput, "mms://", 1)
			end if

			'Tjek om der allerede er et link her
	          HREFTagFound = false
	          if iLinkStart>6 then
	            if lcase(mid(strInput, iLinkStart-6, 5))="href=" then
	              HREFTagFound = true
	            end if
	          end if
			  
	          SRCTagFound = false
	          if iLinkStart>5 then
	            if lcase(mid(strInput, iLinkStart-5, 4))="src=" then
	              SRCTagFound = true
	            end if
	          end if				  

			' Set the position of the end of the link.  I use the
			' first space as the determining factor.
			iLinkSpace = InStr(iLinkStart, strInput, " ", 1)
			'iLinkendLine = InStr(iLinkStart, strInput, vbCrLf, 1)
			iLinkendLine = getFirstNewLine(iLinkStart, strInput)
			if (iLinkSpace>iLinkendLine and iLinkendLine>0) or iLinkSpace=0 then
				iLinkEnd = iLinkendLine
			else
				iLinkEnd = iLinkSpace
			end if

			' If we didn't find a space then we link to the
			' end of the string
			If iLinkEnd = 0 Then iLinkEnd = Len(strInput) + 1
			' Take care of any punctuation we picked up
			'Fjerner højest 20 tegn, for at begrænse server belastning ved fejl.
			count = 0
			do while removeBadChar(strInput, iLinkEnd) and count<20
				count = count + 1
			loop
			' This adds to the output string all the non linked stuff
			' up to the link we're curently processing.
			Append(Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation))
			'strOutput = strOutput & Mid(strInput, iCurrentLocation, iLinkStart - iCurrentLocation)
			' Get the text we're linking and store it in a variable
			strLinkText = Mid(strInput, iLinkStart, iLinkEnd - iLinkStart)
					
			' Build our link and append it to the output string
			if inStr(1,strLinkText, internLinkSite)>0 then
				'Internt link
				internal = True
				if len(intTarget)>0 then
					targetStr = " target=""" & intTarget & """"
				else
					targetStr = ""
				end if
			else
				'Ekstern link
				internal = False
				if len(extTarget)>0 then
					targetStr = " target=""" & extTarget & """"
				else
					targetStr = ""
				end if
			end if
			tempLen = inStr(9,strLinkText,"/")
			If tempLen>0 then
				TempLinkName = left(strLinkText,tempLen-1)
			else
				TempLinkName = strLinkText
			end if
			TempLinkName = Replace(TempLinkName,"http://","")
			TempLinkName = Replace(TempLinkName,"https://","")			
			
			' Validerer om der ikke allerede findes et tag af samme karakter:
			'Skriv selve linket.
			if ( InStr( 1, LCase(strLinkText), "href=") > 0 ) OR ( InStr( 1, LCase(strLinkText), ">") > 0 ) OR ( InStr( 1, LCase(strLinkText), "<") > 0 ) or HREFTagFound or SRCTagFound then
				Append(strLinkText)
				'strOutput = strOutput &strLinkText
			else ' Strengen skal laves om til html
		        
		        if (left(strLinkText, 24)="www.youtube.com/watch?v=" or left(strLinkText, 31)="http://www.youtube.com/watch?v=" or left(strLinkText, 27)="http://youtube.com/watch?v=" or left(strLinkText, 20)="youtube.com/watch?v=") and convertYoutubeLinks then
		            if YoutubeHeight&""="" then
		                if YoutubeWidth&""<>"" and isNumeric(YoutubeWidth) then
		                    YoutubeHeight = round((((YoutubeWidth/4)*3)+25),0)
		                else
		                    YoutubeHeight = 355
		                end if
		            end if
		            if YoutubeWidth&""="" then
		                if YoutubeHeight&""<>"" and isNumeric(YoutubeHeight) then
		                    YoutubeWidth = round((((YoutubeHeight-25)/3)*4),0)
		                else
		                    YoutubeWidth = 425
		                end if
		            end if
		            youtubeId = Replace(strLinkText, "http://www.youtube.com/watch?v=", "")
		            youtubeId = Replace(youtubeId, "www.youtube.com/watch?v=", "")
		            youtubeId = Replace(youtubeId, "http://youtube.com/watch?v=", "")
		            youtubeId = Replace(youtubeId, "youtube.com/watch?v=", "")
		            if useSWFObject then
		                randomId = round(rnd(1000)*10000)
		                Append("<div id=""" & randomId & """></div><script type=""text/javascript"">var fo = new SWFObject(""http://www.youtube.com/v/" & youtubeId & """, ""YouTube"", """ & YoutubeWidth & """, """ & YoutubeHeight & """, ""8"", ""#222222""); fo.addParam(""scale"", ""noscale""); fo.addParam(""align"", ""middle""); fo.addParam(""allowScriptAccess"", ""Always""); fo.addParam(""wmode"", ""transparent""); fo.write(""" & randomId & """);</script>")
                        'strOutput = strOutput & "<div id=""" & randomId & """></div><script type=""text/javascript"">var fo = new SWFObject(""http://www.youtube.com/v/" & youtubeId & """, ""YouTube"", """ & YoutubeWidth & """, """ & YoutubeHeight & """, ""8"", ""#222222""); fo.addParam(""scale"", ""noscale""); fo.addParam(""align"", ""middle""); fo.addParam(""allowScriptAccess"", ""Always""); fo.addParam(""wmode"", ""transparent""); fo.write(""" & randomId & """);</script>"
                    elseif useFlashObject then
		                randomId = round(rnd(1000)*10000)
                        Append("<div id=""" & randomId & """></div><script type=""text/javascript"">var fo = new FlashObject(""http://www.youtube.com/v/" & youtubeId & """, ""YouTube"", """ & YoutubeWidth & """, """ & YoutubeHeight & """, ""8"", ""#222222""); fo.addParam(""scale"", ""noscale""); fo.addParam(""align"", ""middle""); fo.addParam(""allowScriptAccess"", ""Always""); fo.addParam(""wmode"", ""transparent""); fo.write(""" & randomId & """);</script>")
                        'strOutput = strOutput & "<div id=""" & randomId & """></div><script type=""text/javascript"">var fo = new FlashObject(""http://www.youtube.com/v/" & youtubeId & """, ""YouTube"", """ & YoutubeWidth & """, """ & YoutubeHeight & """, ""8"", ""#222222""); fo.addParam(""scale"", ""noscale""); fo.addParam(""align"", ""middle""); fo.addParam(""allowScriptAccess"", ""Always""); fo.addParam(""wmode"", ""transparent""); fo.write(""" & randomId & """);</script>"
		            else
		                Append("<object width=""" & YoutubeWidth & """ height=""" & YoutubeHeight & """><param name=""movie"" value=""http://www.youtube.com/v/" & youtubeId & """></param><param name=""wmode"" value=""transparent""></param><embed src=""http://www.youtube.com/v/" & youtubeId & """ type=""application/x-shockwave-flash"" wmode=""transparent"" width=""" & YoutubeWidth & """ height=""" & YoutubeHeight & """></embed></object>")
		                'strOutput = strOutput & "<object width=""" & YoutubeWidth & """ height=""" & YoutubeHeight & """><param name=""movie"" value=""http://www.youtube.com/v/" & youtubeId & """></param><param name=""wmode"" value=""transparent""></param><embed src=""http://www.youtube.com/v/" & youtubeId & """ type=""application/x-shockwave-flash"" wmode=""transparent"" width=""" & YoutubeWidth & """ height=""" & YoutubeHeight & """></embed></object>"
		            end if
		            youtubeId = ""
		        else
				    if tempType = "http" then
				        if internal then
					        Append("<a href=""" & Replace(strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>")
					        'strOutput = strOutput & "<a href=""" & Replace(strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>"
				        else
				            Append("<a href=""" & Replace(openLinkThrough & strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>")
				            'strOutput = strOutput & "<a href=""" & Replace(openLinkThrough & strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>"
				        end if
				    elseif tempType = "www" then
				        if internal then
					        Append("<a href=""" & "http://" & Replace(strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>")
					        'strOutput = strOutput & "<a href=""" & "http://" & Replace(strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>"
				        else
					        Append("<a href=""" & Replace(openLinkThrough & "http://" & strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>")
					        'strOutput = strOutput & "<a href=""" & Replace(openLinkThrough & "http://" & strLinkText, "&", "&amp;") & """"&targetStr&">" & TempLinkName & "</a>"
				        end if
			        elseif tempType = "mms" then
				        if internal then
					        Append("<A HREF=""" & strLinkText & """"&targetStr&">" & TempLinkName & "</A>")
					        'strOutput = strOutput & "<A HREF=""" & strLinkText & """"&targetStr&">" & TempLinkName & "</A>"
				        else
				            Append("<A HREF=""" & openLinkThrough & strLinkText & """"&targetStr&">" & TempLinkName & "</A>")
					        'strOutput = strOutput & "<A HREF=""" & openLinkThrough & strLinkText & """"&targetStr&">" & TempLinkName & "</A>"
				        end if
				    end if
		        end if
			end if
		end if
		' Reset our current location to the end of that link
		iCurrentLocation = iLinkEnd
	Loop

	' Tack on the end of the string.  I need to do this so we
	' don't miss any trailing non-linked text
	Append(Mid(strInput, iCurrentLocation))
	'strOutput = strOutput & Mid(strInput, iCurrentLocation)

	' Set the return value
	findLinks = ToString()
End Function

%>