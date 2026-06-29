<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLGetNumberEntry.asp
	'Module Name				:	No Series (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 27, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
	dim OutData,Root,newElem,newElem1
	dim objRs,objRs1,sQuery,iSeriesNo

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")

	sQuery="select SeriesNo,Description,CounterType,UsedBy,NumberLength from Ms_NumberSeries order by SeriesNo asc"

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	set objRs.ActiveConnection = nothing

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
	if not objRs.EOF then
		do while not objRs.EOF
			iSeriesNo=objRs(0)
			Set newElem = OutData.createElement("Series")
			newElem.setAttribute "No", trim(objRs(0))
			newElem.setAttribute "Description",trim(objRs(1))
			newElem.setAttribute "Type", trim(objRs(2))
			newElem.setAttribute "UsedBy", trim(objRs(3))
			newElem.setAttribute "NumberLength", trim(objRs(4))
			Root.appendChild newElem
			sQuery="SELECT  EntryNo, Period, Number, IsNull(Prefix,''), Isnull(Suffix,'') FROM Ms_NumberSeriesEntry where SeriesNo="&iSeriesNo&" order by EntryNo asc"
			with objRs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs1.ActiveConnection = nothing
			do while not objRs1.EOF
				Set newElem1 = OutData.createElement("Entry")
				newElem1.setAttribute "EntryNo", trim(objRs1(0))
				newElem1.setAttribute "Period",trim(objRs1(1))
				newElem1.setAttribute "Number",trim(objRs1(2))
				newElem1.setAttribute "Prefix",trim(objRs1(3))
				newElem1.setAttribute "Suffix",trim(objRs1(4))

				newElem.appendChild newElem1
				objRs1.MoveNext
			loop
			objRs1.Close
		objRs.MoveNext
		loop
		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close
%>
