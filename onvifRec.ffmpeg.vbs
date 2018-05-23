' by mr rekcuf
' vbscript to record live stream from onvif camera
' Usage:
' cscript /nolog onvifQuery.vbs camera_IP userID password profile_token filename_to_save -t HH:YY:SS
'

if wscript.arguments.Count < 6 then 
	wscript.echo
	wscript.echo "Usage: "
	wscript.echo " cscript /nologo " &wscript.scriptName  &" camera_IP userID password profile_token filename_to_save " &chr(34) &"-t HH:YY:SS" &chr(34)
	wscript.echo "                                                              " &chr(34) &"-t 00:00:00" &chr(34) & " for continuously recording " 
	wscript.echo
	wscript.quit
end if

Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
xmlDoc.async = False
xmlDoc.validateOnParse   = False
xmlDoc.resolveExternals  = False
ns = "xmlns:soap=""http://www.w3.org/2003/05/selope/"""
ns2 = "xmlns=""http://schemas.microsoft.com/sharepoint/soap/"""
ns3 = "xmlns:tt=""http://www.onvif.org/ver10/schema"""
ns4 = "xmlns:trt=""http://www.onvif.org/ver10/media/wsdl"""
ns5 = "xmlns:tr2=""http://www.onvif.org/ver20/media/wsdl"""

xmlDoc.setProperty "SelectionNamespaces", ns &" " &ns2 &" " &ns3 &" " &ns4 &" " &ns5

GetProfile=_
"<s:Body " + _
"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
"xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" +_  
"<GetProfiles xmlns='http://www.onvif.org/ver10/media/wsdl'/>" +_
"</s:Body>" 

GetServices=_
"<s:Body " + _
"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
"xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" +_  
"<GetServices xmlns='http://www.onvif.org/ver10/device/wsdl'>" +_
"<IncludeCapability>false</IncludeCapability>" +_
"</GetServices>" +_
"</s:Body>" 


GetStreamUri=_
"<s:Body " + _
"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
"xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" +_  
"<GetStreamUri xmlns='http://www.onvif.org/ver20/media/wsdl'><Protocol>RtspOverHttp</Protocol><ProfileToken>REPLACEPROFILE</ProfileToken></GetStreamUri>" +_
"</s:Body>" 

 



xmlstd = _
"xmlns:s='http://www.w3.org/2003/05/soap-envelope' " + _
"xmlns:a='http://www.w3.org/2005/08/addressing'" '''+_ 


xxxx= _
"xmlns:SOAP-ENC='http://www.w3.org/2003/05/soap-encoding' " +_  
"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
"xmlns:xsd='http://www.w3.org/2001/XMLSchema' " +_  
"xmlns:chan='http://schemas.microsoft.com/ws/2005/02/duplex' " +_  
"xmlns:c14n='http://www.w3.org/2001/10/xml-exc-c14n#' " +_
"xmlns:wsu='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd' " +_
"xmlns:xenc='http://www.w3.org/2001/04/xmlenc#' " +_
"xmlns:wsc='http://schemas.xmlsoap.org/ws/2005/02/sc' " +_
"xmlns:ds='http://www.w3.org/2000/09/xmldsig#' " +_
"xmlns:wsse='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd' " +_
"xmlns:xmime5='http://www.w3.org/2005/05/xmlmime' " +_
"xmlns:xmime='http://tempuri.org/xmime.xsd' " +_
"xmlns:xop='http://www.w3.org/2004/08/xop/include' " +_
"xmlns:wsrfbf='http://docs.oasis-open.org/wsrf/bf-2' " +_
"xmlns:wstop='http://docs.oasis-open.org/wsn/t-1' " +_
"xmlns:wsrfr='http://docs.oasis-open.org/wsrf/r-2' " +_
"xmlns:wsnt='http://docs.oasis-open.org/wsn/b-2' " +_
"xmlns:tt='http://www.onvif.org/ver10/schema' " +_
"xmlns:tds='http://www.onvif.org/ver10/device/wsdl' " +_
"xmlns:tev='http://www.onvif.org/ver10/events/wsdl' " +_
"xmlns:tptz='http://www.onvif.org/ver20/ptz/wsdl' " +_
"xmlns:trt='http://www.onvif.org/ver20/media/wsdl' " +_
"xmlns:timg='http://www.onvif.org/ver20/imaging/wsdl' " +_
"xmlns:tmd='http://www.onvif.org/ver10/deviceIO/wsdl' " +_
"xmlns:tns1='http://www.onvif.org/ver10/topics' " +_
"xmlns:ter='http://www.onvif.org/ver10/error' " +_
"xmlns:tds='http://www.onvif.org/ver10/device/wsdl' " +_
"xmlns:tnsaxis='http://www.axis.com/2009/event/topics' "


requestGetProfile = "<?xml version='1.0' encoding='utf-8'?>" + _
"<s:Envelope " +xmlstd +">" +_
GetProfile +_
"</s:Envelope>" '+_



requestGetStreamUri = "<?xml version='1.0' encoding='utf-8'?>" + _
"<s:Envelope " +xmlstd +">" +_
GetStreamUri +_
"</s:Envelope>" '+_


dim profiles(10), streamUris(10)


index= 0


url = "http://" &wscript.arguments.item(0) &"/onvif/device_service"
user =  wscript.arguments.item(1)
password = wscript.arguments.item(2)
profile = wscript.arguments.item(3)
captrueFile = wscript.arguments.item(4)
duration = wscript.arguments.item(5)

if duration = "-t 00:00:00" then duration = ""

streamUri = ""

with CreateObject("MSXML2.ServerXMLHTTP.6.0")

	.open "POST", url, False , user, password
	.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8" 

	.setRequestHeader "Accept-Encoding", "gzip, deflate"
	.setRequestHeader "Connection", "keep-alive"

	lResolve = 30 * 1000  
	lConnect = 60 * 1000  
	lSend = 30 * 1000 
	lReceive = 120 * 1000
	.setTimeouts lResolve, lConnect, lSend, lReceive

	On Error Resume Next
	.send Replace(requestGetProfile, "'", chr(34))

	If Err.Number = 0 Then 

	   	xmlDoc.loadXML(.responseText)
		Set items = xmlDoc.selectNodes("//trt:Profiles")

		WScript.Echo "Found " & items.length & " Profile(s)."
  	
		x = 0
		y = 0 
  		For Each item In items
			profiles(x) = item.getAttribute("token")

	    		WScript.Echo " Profile " &x &" token: " &item.getAttribute("token") &" Name: " &item.selectNodes("tt:Name")(0).text
	
			.open "POST", url, False , user, password
			.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8" 
			.setRequestHeader "Accept-Encoding", "gzip, deflate"
			.setRequestHeader "Connection", "keep-alive"
			requestGetStreamUri = Replace(requestGetStreamUri, "'", chr(34))
			.send Replace(requestGetStreamUri, "REPLACEPROFILE", profiles(x))


		   	xmlDoc.loadXML(.responseText)
			Set itemStreams = xmlDoc.selectNodes("//tr2:Uri")
	  		For Each itemStream In itemStreams
				streamUris(y) = itemStream.text
	    			WScript.Echo "   Found stream: " &itemStream.text 

				if profile = profiles(x) then
					WScript.Echo "   Token matched. "
    					WScript.Echo "   Using the stream: " &itemStream.text 
					streamUri = streamUris(y)
					exit for
				end if	
				y =y + 1
	  		Next
			if streamUri <> "" then exit for
			x = x +1
	  	Next

	elseif  Err.Number = -2147012889 then

		wscript.echo "Invalid IP Address or Hostname. Error code: " +hex(Err.Number)

	else

		wscript.echo "Error code: " +hex(Err.Number)

	end if

	On Error GoTo 0
end with


wscript.echo

suri = Replace(streamUri, "http://", "rtsp://")
suri = Replace(suri, "HTTP://", "rtsp://")
userpassword=user&":" &password &"@"
suri = Replace(suri, "//", "//" &userpassword)


Set objShell = WScript.CreateObject( "WScript.Shell" )


cmdline = objShell.CurrentDirectory &"\ffmpeg -y -rtsp_transport tcp " &duration &" -i " &chr(34) &suri &chr(34) &" -acodec copy -vcodec copy " &objShell.CurrentDirectory &"/" &captrueFile



wscript.echo
wscript.echo "lauching ffmepg using " &cmdline


result= objShell.Run( cmdline )

Set objShell = Nothing
Set objArgs = Nothing

wscript.quit

