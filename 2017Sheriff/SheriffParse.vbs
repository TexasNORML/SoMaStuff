baseUrl = "http://www.sheriffstx.org/county_map/county/"
start_idx = 2524
end_idx = start_idx + 254
' start_idx = 2524 + 24
' end_idx = start_idx + 2

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
	' WScript.echo HTTPPost

	
	' strip the data from the MIDdle of the page
	top_tag = "ptitles"
	tag_idx1 = InStr(HTTPPost,top_tag) + len(top_tag)
	bottom_tag = "<!-- BANNERS -->"
	tag_idx2 = InStr(HTTPPost,bottom_tag)
	HTTPPost = Mid(HTTPPost, tag_idx1, tag_idx2 - tag_idx1)
    ' trim a little more
	top_tag = "<div><p>"
	tag_idx1 = InStr(HTTPPost,top_tag) + len(top_tag)
	HTTPPost = Mid(HTTPPost, tag_idx1, tag_idx2 - tag_idx1)

	' get rid of the image
	img_idx = InStr(HTTPPost, "<img")
	If (img_idx > 0) Then
	    img_end = InStr(HTTPPost, ">") + 1
	    img_substr = Mid(HTTPPost, img_idx, img_end - img_idx)
		HTTPPost = Replace(HTTPPost, img_substr, "")
	End If
	
	' get rid of easy unwanted markup
	HTTPPost = Replace(HTTPPost, "</p>", "")
	HTTPPost = Replace(HTTPPost, "<p>", "")
	HTTPPost = Replace(HTTPPost, "</a>", "")
	HTTPPost = Replace(HTTPPost, "</div>", "")
	
	' Remove white space
	HTTPPost = Replace(HTTPPost, vbCrLf, "")
	HTTPPost = Replace(HTTPPost, vbTab, "")
	
	' Replace any ',' so they don't interfere with our tokens
	HTTPPost = Replace(HTTPPost, ",", "")
	HTTPPost = Replace(HTTPPost, "<br />", ",")
	
	'  remove the anchor for the email
	email_atag = "<a title=""Email"
	email_idx = InStr(HTTPPost, email_atag)
	If (email_idx > 0) Then
	    email_end = InStr(HTTPPost, ">") + 1
	    email_substr = Mid(HTTPPost, email_idx, email_end - email_idx)
		HTTPPost = Replace(HTTPPost, email_substr, "")
	End If
	HTTPPost = Replace(HTTPPost, "Email: ", "")
	HTTPPost = Replace(HTTPPost, "Email:", "")
	
    ' add the index
	first_char = Mid(HTTPPost,1,1)
    If Asc(first_char) = 160 Or Asc(first_char) = 44 Then
	    HTTPPost = Right(HTTPPost, len(HTTPPost)-1)
	End If
	' WScript.Echo Asc(first_char)
	' WScript.Echo HTTPPost
	
	HTTPPost = start_idx & "," & HTTPPost
	objFile.Write HTTPPost & vbCrLf
	start_idx = start_idx + 1
Loop

objFile.Close