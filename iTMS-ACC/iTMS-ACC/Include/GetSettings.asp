<%
	'Program Name				:	GetSettings.asp
	'Module Name				:	General
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 30, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>

<%
' Function to get the General Settings

	Function GetSettings(sIP)
		dim objFS,sPath,objFS1,arrTemp,sStr
		const cReading = 1

		set objFS = Server.CreateObject("Scripting.FileSystemObject")
		sPath = Server.MapPath("/include/Settings.inf")
		set objFS1 = objFS.OpenTextFile (sPath,cReading)
		if sIP = "IP" then
			sStr = objFS1.ReadLine
			arrTemp = split(sStr,"|")
			GetSettings = trim(arrTemp(1))
		end if

	End Function
%>