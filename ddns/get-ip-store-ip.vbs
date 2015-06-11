' author: https://github.com/aaaia
' date: 20150609_0914 CEST
' desctription: set actual external ip address into deployed google app script as its property
'     also updates ip in no-ip.pl ddns service

Option Explicit

Dim externalIp

' loads cfg* variables
' example config.vbs.repoignore content:
' '''''''''''''''''''''''''''''''''''
' cfgIpServiceAddressWithQuery="http://ipinfo.io/ip"
' ' here You should put Your script publication address (can be viewed in "Publish"->"Deploy as web app" google app script menu
' cfgKeyValueMapService="https://script.google.com/macros/s/A*********Q/exec?oper=set&key=externalIp"
' cfgDebugMode=true
' cfgNoIpSrvBasicAddressWithQuery="http://update.no-ip.pl/?hostname="
' Dim cfgNoIpSrvHostnames(2)
' cfgNoIpSrvHostnames(0)="t*********m.noip.pl"
' cfgNoIpSrvHostnames(1)="t*********m.no-ip.pl"
' cfgNoIpSrvHostnames(2)="t*********m.no-ip.eu"
' cfgNoIpSrvUser="login"
' cfgNoIpSrvPassword="password"
' '''''''''''''''''''''''''''''''''''
includeFile "config.vbs.repoignore"

' get external ip address
externalIp = makeRestRequest(cfgIpServiceAddressWithQuery)

'set external ip address in prepared service
makeRestRequest(cfgKeyValueMapService & "&val=" & externalIp)

'set external ip address in no-ip.pl 
Dim hostname
For Each hostname In cfgNoIpSrvHostnames
	Call makeRestRequestAuth(cfgNoIpSrvBasicAddressWithQuery & hostname, cfgNoIpSrvUser, cfgNoIpSrvPassword)
Next

' source http://stackoverflow.com/questions/316166/how-do-i-include-a-common-file-in-vbscript-similar-to-c-include
Sub includeFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

' author: https://github.com/aaaia
' date: 20150609_0914 CEST
' description: make http request to rest service
function makeRestRequest(fullAddresWithQuery)
	makeRestRequest = makeRestRequestAuth(fullAddresWithQuery, null, null)
end function

function makeRestRequestAuth(fullAddresWithQuery, user, password) 
	Dim restReq
	Dim restResp
	
	if (cfgDebugMode) then
	    Wscript.Echo("Making request: " & fullAddresWithQuery)
	end if
	
	' You can use Microsoft.XMLHTTP but in some sases it will result:
    '     Error number: 800700
    '     Descrption: Access is denied.
    '     Source : msxml3.dll
	' When you try to read response from page
	Set restReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")

	if (IsNull(user)) then
		restReq.open "GET", fullAddresWithQuery, false
	else
		restReq.open "GET", fullAddresWithQuery, false, user, password
	end if
	restReq.send

	restResp = restReq.responseText
	
	if (cfgDebugMode) then
	    Wscript.Echo("Request result: " & restResp)
	end if
	
	makeRestRequestAuth = restResp
End function
