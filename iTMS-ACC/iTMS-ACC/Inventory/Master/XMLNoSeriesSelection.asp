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
	'Program Name				:	XMLNoSeriesSelection.asp
	'Module Name				:	Inventory (No Series Selection)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 20, 2004
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

<!--#include virtual="/include/DatabaseConnection.asp"-->

<%
	dim dcrs,dcrs1,sSql,OutData,Root,newElem
	dim sUnit,sItem,sActivity,sQuery,sCatCode,sClassCode
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	sUnit = Request("sUnit")
	sItem = Request("sItem")
	sActivity =Request("sActivity")
	sCatCode = Request("CatCode")
	sClassCode = Request("ClassCode")
	
	if Trim(sCatCode)="0" then sCatCode = ""
	if Trim(sClassCode)="0" then sClassCode = ""
	
	sQuery = "Select H.SeriesNo,H.SeriesCode from INV_M_NUMBERSERIES H left join INV_M_NoSeriesClass D on H.SeriesCode = D.SeriesCode where ORGANISATIONCODE = '" & sUnit & "' and ActivityType ='"& sActivity &"' "
	if Trim(sCatCode)<>"" then
	    sQuery= sQuery&" and CatCode in ("& sCatCode &")"
	end if 
	if Trim(sClassCode)<>"" then
	    sQuery= sQuery&" and ClassCode in ("& sClassCode &")"
	end if 
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		Set Root = OutData.createElement("Root")
		OutData.appendChild Root
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT NUMBER, PREFIX, SUFFIX ,PERIOD FROM APP_R_NOSERIESMODULEENTRY WHERE OUDEFINITIONID = '" & sUnit & "' AND SERIESNO = " & dcrs(0) & " AND SERIESCODE = " & dcrs(1) & " ORDER BY ENTRYNO"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		while not dcrs1.EOF 
			Set newElem = OutData.createElement("DETAILS")
			newElem.setAttribute "NUMBER", trim(dcrs1(0))
			newElem.setAttribute "PREFIX", trim(dcrs1(1))
			newElem.setAttribute "SUFFIX", trim(dcrs1(2))
			newElem.setAttribute "PERIOD", trim(dcrs1(3))

			Root.appendChild newElem

			dcrs1.MoveNext
		wend
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
