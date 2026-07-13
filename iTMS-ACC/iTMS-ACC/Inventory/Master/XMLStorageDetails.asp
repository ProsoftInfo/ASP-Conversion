<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLStorageDetails.asp
	'Module Name				:	Inventory (Item creation / Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	October 12, 2004
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

<!-- #include File="../../include/DatabaseConnection.asp" -->

<%
	dim dcrs,dcrs1,sSql,OutData,Root,newElem,newElem1,sOrgID

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	
	sOrgID = trim(Request("sOrgID"))
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		if sOrgID = "ALL" then
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME,OUDEFINITIONID FROM INV_M_STORAGE where ApplicableFor ='IN' ORDER BY 3"
		else
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME,OUDEFINITIONID FROM INV_M_STORAGE WHERE ApplicableFor ='IN' and OUDEFINITIONID = '" & sOrgID & "' ORDER BY 3"
		end if
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("ROOT")												
	OutData.appendChild Root
	if not dcrs.EOF then
		do while not dcrs.EOF
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT DISTINCT BINNUMBER,BINCODE,BINNAME FROM INV_M_STOREBINDETAILS WHERE OUDEFINITIONID = '" & trim(dcrs(3)) & "' AND LOCATIONNUMBER = " & trim(dcrs(0)) & " ORDER BY 1"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			if not dcrs1.EOF then
				do while not dcrs1.EOF
					Set newElem1 = OutData.createElement("STORAGE")
					newElem1.setAttribute "UNIT", trim(dcrs(3))
					newElem1.setAttribute "STNUMBER", trim(dcrs(0))
					newElem1.setAttribute "STCODE", trim(dcrs(1))
					'newElem1.setAttribute "STNAME", trim(dcrs(2)) & " -- " & trim(dcrs1(2))
					newElem1.setAttribute "STNAME", trim(dcrs(2)) & " -- " & trim(dcrs1(1))	'Modified By Ragavendran R,On 2nd June 2011
					
					newElem1.setAttribute "BINNUMBER", trim(dcrs1(0))
					Root.appendChild newElem1
				dcrs1.MoveNext
				loop
			else
				Set newElem1 = OutData.createElement("STORAGE")
				newElem1.setAttribute "UNIT", trim(dcrs(3))
				newElem1.setAttribute "STNUMBER", trim(dcrs(0))
				newElem1.setAttribute "STCODE", trim(dcrs(1))
				newElem1.setAttribute "STNAME", trim(dcrs(2))
					
				newElem1.setAttribute "BINNUMBER", "0"
				Root.appendChild newElem1
			end if
			dcrs1.Close
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml

%>
