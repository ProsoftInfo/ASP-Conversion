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
	'Program Name				:	XMLGetJobWorkLotDetails.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	September 16, 2004
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
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/Populate.asp" -->

<%
	dim dcrs,dcrs1,sSql,OutData,sType,Root,newElem,NewNode
	dim iRecNo, sFlag, iItem, iClass, sOrgID, sRecType

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	Set Root = OutData.createElement("STOREDLOTDETAILS")
	OutData.appendChild Root

	iRecNo = Request.QueryString("iRecNo")
	sFlag = Request.QueryString("sFlag")
	iItem = Request.QueryString("iItem")
	iClass = Request.QueryString("iClass")
	sOrgID = Request.QueryString("sOrgID")
	sRecType = Request.QueryString("sRecType")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(LOTWEIGHINGTYPE,'') FROM SUB_T_SALEARPACKRECDLOT WHERE ACTUALRECEIPTNO = " & iRecNo & " AND FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass
		.ActiveConnection = con
		.Open
		'Response.Write dcrs.Source 
	end with
	set dcrs.ActiveConnection = nothing
	
	if not dcrs.EOF then sRecType = Trim(dcrs(0))
	dcrs.Close 
	
	if sRecType = "LS" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ACTUALRECEIPTENTRYNO, MILLLOTNO FROM VW_JOBWORKRECEIPT_LOT WHERE ACTUALRECEIPTNO = " & iRecNo & " AND  FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID)
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.EOF
			set NewNode = OutData.createElement("LOT")
			NewNode.setAttribute "LOTNO", trim(dcrs(1))

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT MILLLOTNO, MILLSERIALNO, MILLNETTWEIGHT, MILLTAREWEIGHT, ISNULL(SELLINGUNITID,0), WEIGHTPERSELLINGFORM, PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,''), MILLGROSSWEIGHT FROM VW_JOBWORKRECEIPT_SERIAL WHERE ACTUALRECEIPTNO = " & iRecNo & " AND ACTUALRECEIPTENTRYNO = " & trim(dcrs(0)) & " AND MILLLOTNO = " & pack(trim(dcrs(1))) & " AND  FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID)
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			do while not dcrs1.eof
				Set newElem = OutData.createElement("SERIAL")
				newElem.setAttribute "LOT", trim(dcrs1(0))
				newElem.setAttribute "LOTSERIAL", trim(dcrs1(1))
				newElem.setAttribute "QTYREC", trim(dcrs1(2))
				newElem.setAttribute "TAREREC", trim(dcrs1(3))
				newElem.setAttribute "SELLINGTYPE", trim(dcrs1(4))
				newElem.setAttribute "WEIGHTSTYPE", trim(dcrs1(5))
				newElem.setAttribute "PACKINGTYPE", trim(dcrs1(6))
				newElem.setAttribute "SELLINGFORM", trim(dcrs1(7))
				newElem.setAttribute "PACKNUMBER", trim(dcrs1(8))
				newElem.setAttribute "QTYGRO", trim(dcrs1(9))
				newElem.setAttribute "FLAG", sFlag
				NewNode.appendChild newElem
				dcrs1.MoveNext
			loop
			dcrs1.Close

			root.appendChild NewNode
			dcrs.MoveNext
		loop
		dcrs.Close

	elseif sRecType="S" then
		set NewNode = OutData.createElement("LOT")
		NewNode.setAttribute "LOTNO", ""

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MILLLOTNO,''), MILLSERIALNO, MILLNETTWEIGHT, MILLTAREWEIGHT, ISNULL(SELLINGUNITID,0), ISNULL(WEIGHTPERSELLINGFORM,0), PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,''), MILLGROSSWEIGHT FROM VW_JOBWORKRECEIPT_SERIAL WHERE ACTUALRECEIPTNO = " & iRecNo & " AND  FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID)
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		do while not dcrs1.eof
			Set newElem = OutData.createElement("SERIAL")
			newElem.setAttribute "LOT", trim(dcrs1(0))
			newElem.setAttribute "LOTSERIAL", trim(dcrs1(1))
			newElem.setAttribute "QTYREC", trim(dcrs1(2))
			newElem.setAttribute "TAREREC", trim(dcrs1(3))
			newElem.setAttribute "SELLINGTYPE", trim(dcrs1(4))
			newElem.setAttribute "WEIGHTSTYPE", trim(dcrs1(5))
			newElem.setAttribute "PACKINGTYPE", trim(dcrs1(6))
			newElem.setAttribute "SELLINGFORM", trim(dcrs1(7))
			newElem.setAttribute "PACKNUMBER", trim(dcrs1(8))
			newElem.setAttribute "QTYGRO", trim(dcrs1(9))
			newElem.setAttribute "FLAG", sFlag
			NewNode.appendChild newElem
			dcrs1.MoveNext
		loop
		dcrs1.Close

		root.appendChild NewNode

	elseif sRecType = "L" then
		set NewNode = OutData.createElement("LOT")
		NewNode.setAttribute "LOTNO", ""

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ACTUALRECEIPTENTRYNO, MILLLOTNO, MILLLOTWEIGHT, MILLLOTTAREWT, ISNULL(SELLINGUNITID,''), WEIGHTPERSELLINGFORM, PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,'') from VW_JOBWORKRECEIPT_LOT WHERE ACTUALRECEIPTNO = " & iRecNo & " AND  FROMITEMCODE = " & iItem & " AND FROMCLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID)
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		do while not dcrs1.EOF
			Set newElem = OutData.createElement("SERIAL")
			newElem.setAttribute "LOT", trim(dcrs1(1))
			newElem.setAttribute "LOTSERIAL", ""
			newElem.setAttribute "QTYREC", trim(dcrs1(2))
			newElem.setAttribute "TAREREC", trim(dcrs1(3))
			newElem.setAttribute "SELLINGTYPE", trim(dcrs1(4))
			newElem.setAttribute "WEIGHTSTYPE", trim(dcrs1(5))
			newElem.setAttribute "PACKINGTYPE", trim(dcrs1(6))
			newElem.setAttribute "SELLINGFORM", trim(dcrs1(7))
			newElem.setAttribute "PACKNUMBER", trim(dcrs1(8))
			newElem.setAttribute "QTYGRO",trim(dcrs1(2))
			newElem.setAttribute "FLAG", sFlag
			NewNode.appendChild newElem
			dcrs1.MoveNext
		loop
		dcrs1.Close

		root.appendChild NewNode

	end if
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
