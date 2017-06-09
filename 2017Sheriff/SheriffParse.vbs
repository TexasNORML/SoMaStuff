baseUrl = "http://www.sheriffstx.org/county_map/county/"
start_idx = 2524
end_idx = 2524 + 254


Set objFSO=CreateObject("Scripting.FileSystemObject")

' How to write file
outFile="sherrif_info.csv"
Set objFile = objFSO.CreateTextFile(outFile,True)



Do While start_idx < end_idx
    url = baseUrl + CStr(start_idx)
	
	' make the post and get the content
	set oHTTP = CreateObject("Microsoft.XMLHTTP")
	oHTTP.open "GET", url,false
	oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oHTTP.setRequestHeader "Content-Length", Len(sRequest)
	oHTTP.send sRequest
	HTTPPost = oHTTP.responseText

	' strip off top chunk
	top_tag = "ptitles"
	top_tag2 = "</p>"
	end_top_idx = InStr(HTTPPost,top_tag)
	end_top_idx = InStr(end_top_idx, HTTPPost, top_tag2)
	HTTPPost = Right(HTTPPost, len(HTTPPost)-end_top_idx+1)
	
	' strip off bottom chunk
	bottom_tag = " <!-- BANNERS -->"
	end_btm_idx = InStr(HTTPPost,bottom_tag)
    HTTPPost = Left(HTTPPost, end_btm_idx)
	
	'get rid of easy unwanted markup
	HTTPPost = Replace(HTTPPost, "</p>", "")
	HTTPPost = Replace(HTTPPost, "<p>", "")
	HTTPPost = Replace(HTTPPost, "</a>", "")
	HTTPPost = Replace(HTTPPost, "</div>", "")
	' Not Working HTTPPost = Trim(HTTPPost)
	
	' Remove white space
	HTTPPost = Replace(HTTPPost, vbCrLf, "")
	HTTPPost = Replace(HTTPPost, vbTab, "")
	
	' Replace any ',' so they don't interfere with our tokens
	HTTPPost = Replace(HTTPPost, ",", "")
	HTTPPost = Replace(HTTPPost, "<br />", ",")
	
	'  remove the anchor for the email
	' estart = InStr(HTTPPost, "<")
	' eend = InStrRev(HTTPPost, ">")
	' ss_len = eend - eestart
	'  substr = Mid(HTTPPost, estart, ss_len)
	HTTPPost = Replace(HTTPPost, "<", ",")
	HTTPPost = Replace(HTTPPost, ">", ",")
	
  
	WScript.Echo HTTPPost
	objFile.Write HTTPPost & vbCrLf
	start_idx = start_idx + 1
Loop

objFile.Close