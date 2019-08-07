<%
' url = url that we want to parse
' key = name of the querystring pameter
' default_ = the value we want returned íf no value is found
function getQuerystringValue(url, key, default_)

  Set re = New RegExp
  re.IgnoreCase = True
  re.Global = True
  
  re.Pattern = "[\\?&]*"+key+"=([^&#]*)"
  
  Set matches = re.Execute(url)
  If matches.Count > 0 Then
    Set match = matches(0)
    If match.SubMatches.Count > 0 Then
      getQuerystringValue  = match.SubMatches(0) 
    else
      getQuerystringValue = default_
    end if 
  else
    getQuerystringValue = default_
  end if
  
  getQuerystringValue = URLDecode(getQuerystringValue)
  
  set re = nothing
end function  

Function getMimeTypeFromFileExtensionInGlobalInc(fileExtension)
	Dim mimeType
	Select Case fileExtension
	Case "avi"
		mimeType = "video/x-msvideo"
	Case "bmp"
		mimeType = "image/bmp"
	Case "doc"
		mimeType = "application/msword"
	Case "docm"
		mimeType = "application/vnd.ms-word.document.macroEnabled.12"
	Case "docx"
		mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	Case "dot"
		mimeType = "application/msword"
	Case "dotx"
		mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
	Case "gif"
		mimeType = "image/gif"
	Case "jpe"
		mimeType = "image/jpeg"
	Case "jpeg"
		mimeType = "image/jpeg"
	Case "jpg"
		mimeType = "image/jpeg"
	Case "mov"
		mimeType = "video/quicktime"
	Case "mp2"
		mimeType = "video/mpeg"
	Case "mp3"
		mimeType = "audio/mpeg"
	Case "mpa"
		mimeType = "video/mpeg"
	Case "mpe"
		mimeType = "video/mpeg"
	Case "mpeg"
		mimeType = "video/mpeg"
	Case "mpg"
		mimeType = "video/mpeg"
	Case "mpp"
		mimeType = "application/vnd.ms-project"
	Case "mpv2"
		mimeType = "video/mpeg"
	Case "pdf"
		mimeType = "application/pdf"
	Case "potm"
		mimeType = "application/vnd.ms-powerpoint.template.macroEnabled.12"
	Case "potx"
		mimeType = "application/vnd.openxmlformats-officedocument.presentationml.template"
	Case "ppam"
		mimeType = "application/vnd.ms-powerpoint.addin.macroEnabled.12"
	Case "pps"
		mimeType = "application/vnd.ms-powerpoint"
	Case "ppsm"
		mimeType = "application/vnd.ms-powerpoint.slideshow.macroEnabled.12"
	Case "ppsx"
		mimeType = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
	Case "ppt"
		mimeType = "application/vnd.ms-powerpoint"
	Case "pptm"
		mimeType = "application/vnd.ms-powerpoint.presentation.macroEnabled.12"
	Case "pptx"
		mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
	Case "qt"
		mimeType = "video/quicktime"
	Case "rtf"
		mimeType = "application/rtf"
	Case "tif"
		mimeType = "image/tiff"
	Case "tiff"
		mimeType = "image/tiff"
	Case "wav"
		mimeType = "audio/x-wav"
	Case "wks"
		mimeType = "application/vnd.ms-works"
	Case "wps"
		mimeType = "application/vnd.ms-works"
	Case "xla"
		mimeType = "application/vnd.ms-excel"
	Case "xlam"
		mimeType = "application/vnd.ms-excel.addin.macroEnabled.12"
	Case "xlc"
		mimeType = "application/vnd.ms-excel"
	Case "xlm"
		mimeType = "application/vnd.ms-excel"
	Case "xls"
		mimeType = "application/vnd.ms-excel"
	Case "xlsb"
		mimeType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12"
	Case "xlsm"
		mimeType = "application/vnd.ms-excel.sheet.macroEnabled.12"
	Case "xlsx"
		mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	Case "xlt"
		mimeType = "application/vnd.ms-excel"
	Case "xltm"
		mimeType = "application/vnd.ms-excel.template.macroEnabled.12"
	Case "xltx"
		mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.template"
	Case "xlw"
		mimeType = "application/vnd.ms-excel"
	Case "zip"
		mimeType = "application/zip"
	Case ""
		mimeType = "application/octet-stream"
	Case Else
		mimeType = "application/octet-stream"
	End Select
	
	getMimeTypeFromFileExtensionInGlobalInc = mimeType
End Function

Function URLDecode(sConvert)
	Dim aSplit
	Dim sOutput
	Dim I
	If IsNull(sConvert) or sConvert&"" = "" Then
	   URLDecode = ""
	   Exit Function
	End If

	' convert all pluses to spaces
	sOutput = REPLACE(sConvert, "+", " ")

	' next convert %hexdigits to the character
	aSplit = Split(sOutput, "%")

	If IsArray(aSplit) Then
	  sOutput = aSplit(0)
	  For I = 0 to UBound(aSplit) - 1
		sOutput = sOutput & _
		  Chr("&H" & Left(aSplit(i + 1), 2)) &_
		  Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
	  Next
	End If

	URLDecode = sOutput
End Function


%>