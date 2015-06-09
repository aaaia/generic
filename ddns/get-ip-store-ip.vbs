' author: https://github.com/aaaia
' date: 20150609_0914 CEST
' desctription: set actual external ip address into deployed google app script as its property

Option Explicit

Dim configFile, configFileObject, ipSrvRsp, keyValueMapSrvRsp, keyValueMapSrvQuery

' loads cfg* variables
' example config.vbs.repoignore content:
' '''''''''''''''''''''''''''''''''''
' cfgIpServiceAddressWithQuery="http://ipinfo.io/ip"
' ' here You should put Your script publication address (can be viewed in "Publish"->"Deploy as web app" google app script menu
' cfgKeyValueMapService="https://script.google.com/macros/s/A***************************lQ/exec?oper=set&key=externalIp"
' cfgDebugMode=false
' '''''''''''''''''''''''''''''''''''
includeFile "config.vbs.repoignore"

' get external ip address
ipSrvRsp = makeRestRequest(cfgIpServiceAddressWithQuery)

keyValueMapSrvQuery = cfgKeyValueMapService&"&val="&ipSrvRsp

'set external ip address in prepared service
keyValueMapSrvRsp = makeRestRequest(keyValueMapSrvQuery)

' source http://stackoverflow.com/questions/316166/how-do-i-include-a-common-file-in-vbscript-similar-to-c-include
Sub includeFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

' author: https://github.com/aaaia
' date: 20150609_0914 CEST
' description: make http request to rest service
function makeRestRequest(fullAddresWithQuery) 
	Dim restReq
	Dim restResp
	
	' You can use Microsoft.XMLHTTP but in some sases it will result:
    '     Error number: 800700
    '     Descrption: Access is denied.
    '     Source : msxml3.dll
	' When you try to read response from page
	Set restReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")

	restReq.open "GET", fullAddresWithQuery, false
	restReq.send

	restResp = restReq.responseText
	
	if (cfgDebugMode) then
	    Wscript.Echo("Request result: " & restResp)
	end if
	
	makeRestRequest = restResp
End function
