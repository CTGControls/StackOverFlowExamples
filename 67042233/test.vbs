Set FSO = CreateObject("Scripting.FileSystemObject")

' How To Write To A File
Set File = FSO.CreateTextFile("C:\Users\CTGControls\Desktop\Foobar.html",True)
File.Write cstr(http("GET", "https://www.google.com/search?q=bango+plc&tbm=nws", "text/html; charset=UTF-8", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", ""))
File.Close

Set FSO = Nothing
Set File = Nothing






call MsgBox(http("GET", "https://www.google.com/search?q=bango+plc&tbm=nws", "text/html; charset=UTF-8", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", ""))



''MsgBox(httpGet("https://localhost:5001/api/departments?pageNumber=1&pageSize=1", "application/xml; charset=UTF-8", "application/xml"))
Sub httpGet(sUrl, sRequestHeaderContentType, sRequestHeaderAccept)
	Call http("GET", sUrl, sRequestHeaderContentType, sRequestHeaderAccept, "")
End Sub



''MsgBox(httpPost("https://localhost:5001/api/departments?userfriendlyName=987Junk", "application/xml; charset=UTF-8", "application/xml", ""))
Sub httpPost(sUrl,sRequestHeaderContentType, sRequestHeaderAccept, sbody)
	Call http("POST", sRequestHeaderContentType, sRequestHeaderAccept, sbody)
End Sub

Function http(httpCommand, sUrl, sRequestHeaderContentType, sRequestHeaderAccept, sbody)
		Err.Clear
		Dim oXML 'AS XMLHTTP60
		'Set oXML = CreateObject("msxml2.XMLHTTP.6.0")
		Set oXML = CreateObject("Msxml2.ServerXMLHTTP.6.0")
		Dim aErr
		
	On Error Resume Next
		Call oXML.Open(CStr(httpCommand), CStr(sUrl), False)
		'oXML.setRequestHeader "User-Agent", "Mozilla/4.0"
		oXML.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36"
  		'oXML.setRequestHeader "Authorization", "Basic base64encodeduserandpassword"
		oXML.setRequestHeader "Content-Type", CStr(sRequestHeaderContentType)
		'oXML.setRequestHeader "Content-Type", "text/xml"
		oXML.setRequestHeader "CharSet", "charset=UTF-8"
		'oXML.setRequestHeader "Accept", "*/*"
		oXML.setRequestHeader "Accept", CStr(sRequestHeaderAccept)
		oXML.setRequestHeader "cache-control", "no-cache"
		oXML.setRequestHeader "sec-ch-ua","Google Chrome;v=89, Chromium;v=89, ;Not A Brand;v=99"
		
		aErr = Array(Err.Number, Err.Description)

	On Error Goto 0
		 If 0 = aErr(0) Then
	On Error Resume Next
				Call oXML.send(sbody)
				aErr = Array(Err.Number, Err.Description)
	On Error Goto 0
				Select Case True
					Case 0 <> aErr(0)
						Trace("send failed: " & CStr(aErr(0)) & " " & CStr(aErr(1)))
					Case 200 = oXML.status
						'Trace(sUrl & "    HttpStatusCode:" & oXML.status & "    HttpStatusText:" & oXML.statusText)
						http = oXML.responseText
					Case 201 = oXML.status
						Trace(sUrl & "    HttpStatusCode:" & oXML.status & "    HttpStatusText:" & oXML.statusText)
					Case Else
						Trace("further work needed:")
						Trace("URL:" & CStr(sUrl) & "      Message Status:" & CStr(oXML.status) & "      Message Text:" & CStr(oXML.statusText))
						Trace("further work needed:")
				End Select
		Else
			Trace("open failed: " & CStr(aErr(0)) & " " & CStr(aErr(1)))
		End If
	
	'httpPost.HttpStatusCode = cstr(oXML.status)
	'httpPost.HttpStatusText = cstr(oXML.statusText)
	'httpPost.responseText = cstr(oXML.responseText)
	
	Set oXML = Nothing
End Function

Function Trace(Message1)
	MsgBox(Message1)
End Function
