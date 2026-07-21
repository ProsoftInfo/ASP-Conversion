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
	'Program Name				:	XMLGetLotDetails.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	TAJUDEEN S
	'Created On					:	May 25, 2004
	'Modified By				:	KUMAR K A
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
	dim dcrs,dcrs1,dcrs2,sSql,OutData,sType,Root,newElem,NewNode
	dim iRecNo, sFlag, iItem, iClass, sOrgID, sRecType
	dim dActionQty
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")

	Set Root = OutData.createElement("STOREDLOTDETAILS")
	OutData.appendChild Root

	iRecNo = Request.QueryString("iRecNo")
	sFlag = Request.QueryString("sFlag")
	iItem = Request.QueryString("iItem")
	iClass = Request.QueryString("iClass")
	sOrgID = Request.QueryString("sOrgID")
	sRecType = Request.QueryString("sRecType")
	'Response.Write sRecType & "ok"

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT RECEIPTNUMBER FROM RCV_T_ActualRcptLotSerial WHERE RECEIPTNUMBER = " & iRecNo & " and MillLotNo is not null"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	'Response.Write sRecType
	If Not dcrs.EOF Then
		sRecType = "LS"
	End If
	dcrs.Close

	if sRecType = "LS" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ENTRYNO, ISNULL(MILLLOTNO,'') FROM VW_PURCHASE_RCPT_LOT WHERE RECEIPTNUMBER = " & iRecNo & " AND  ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.EOF
			set NewNode = OutData.createElement("LOT")
			NewNode.setAttribute "LOTNO", Trim(dcrs(1))

			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MILLLOTNO,''), MILLSERIALNO, MILLNETTWEIGHT, MILLTAREWEIGHT, ISNULL(SELLINGUNITID,0), ISNULL(WEIGHTPERSELLINGFORM,0) , PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,''), MILLGROSSWEIGHT FROM VW_PURCHASE_RCPT_SERIAL WHERE RECEIPTNUMBER = " & iRecNo & " AND ENTRYNO = " & trim(dcrs(0)) & " AND ISNULL(MILLLOTNO,'') = " & pack(trim(dcrs(1))) & " AND  ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID) & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			do while not dcrs1.eof
			with dcrs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT isNull(Sum(actiononQty),0) FROM RCV_T_ItemActionTaken WHERE ReceiptNumber = " & iRecNo & " AND EntryNo = " & dcrs(0) & " AND actionon='R'"
				.ActiveConnection = con
				.Open
			end with
			'Response.Write dQtyAccep
			set dcrs2.ActiveConnection = nothing
			if not dcrs2.EOF then
				dActionQty = dcrs2(0)
			end if
			dcrs2.close

				Set newElem = OutData.createElement("SERIAL")
				newElem.setAttribute "LOT", trim(dcrs1(0))
				newElem.setAttribute "LOTSERIAL", trim(dcrs1(1))
				newElem.setAttribute "QTYREC", trim(dcrs1(2))
				newElem.setAttribute "QTYACC", cDbl(trim(dcrs1(2)))-cDbl(dActionQty)
				newElem.setAttribute "TAREREC",cDbl(trim(dcrs1(3)))-cDbl(dActionQty)
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
			.Source = "SELECT ISNULL(MILLLOTNO,''), MILLSERIALNO, MILLNETTWEIGHT, MILLTAREWEIGHT, ISNULL(SELLINGUNITID,0), ISNULL(WEIGHTPERSELLINGFORM,0), PACKINGCODE, ISNULL(PACKINGFORM,0), ISNULL(MILLPACKINGNUMBER,''), MILLGROSSWEIGHT FROM VW_PURCHASE_RCPT_SERIAL WHERE RECEIPTNUMBER = " & iRecNo & " AND  ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID) & ""
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
			.Source = "SELECT ENTRYNO, MILLLOTNO, MILLLOTWEIGHT, MILLLOTTAREWT, ISNULL(SELLINGUNITID,''), ISNULL(WEIGHTPERSELLINGFORM,0), PACKINGCODE, ISNULL(WEIGHTPERPACK,0), ISNULL(MILLPACKINGNUMBER,'') from VW_PURCHASE_RCPT_LOT WHERE RECEIPTNUMBER = " & iRecNo & " AND  ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		do while not dcrs1.EOF

			with dcrs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT isNull(Sum(actiononQty),0) FROM RCV_T_ItemActionTaken WHERE ReceiptNumber = " & iRecNo & " AND actionon='R'"
				.ActiveConnection = con
				.Open
			end with
			'Response.Write dQtyAccep
			set dcrs2.ActiveConnection = nothing
			if not dcrs2.EOF then
				dActionQty = dcrs2(0)
			end if
			dcrs2.close

			Set newElem = OutData.createElement("SERIAL")
			newElem.setAttribute "LOT", trim(dcrs1(1))
			newElem.setAttribute "LOTSERIAL", ""
			newElem.setAttribute "QTYREC", trim(dcrs1(2))
			newElem.setAttribute "QTYACC", cDbl(trim(dcrs1(2)))-cDBl(dActionQty)
			newElem.setAttribute "TAREREC", cDbl(trim(dcrs1(3)))'-cDbl(dActionQty)
			newElem.setAttribute "SELLINGTYPE", trim(dcrs1(4))
			newElem.setAttribute "WEIGHTSTYPE", trim(dcrs1(5))
			newElem.setAttribute "PACKINGTYPE", trim(dcrs1(6))
			newElem.setAttribute "SELLINGFORM", "N/A"
			newElem.setAttribute "PACKNUMBER", trim(dcrs1(8))
			newElem.setAttribute "QTYGRO",trim(dcrs1(2))-cDBl(dActionQty)
			newElem.setAttribute "FLAG", sFlag
			NewNode.appendChild newElem
			dcrs1.MoveNext
		loop
		dcrs1.Close

		root.appendChild NewNode

	end if
	Response.Clear 
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
