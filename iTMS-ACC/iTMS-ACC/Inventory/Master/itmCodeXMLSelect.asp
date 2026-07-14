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
	'Program Name				:	itmCodeXMLSelect.asp
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	April 23, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	itmCreationEntry.asp
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
	dim dcrs,dcrs1,Root,newElem,OutData,sOrgCode

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	sOrgCode = trim(Request("sOrgCode"))

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT COMPANYITEMCODE,ISNULL(SHORTDESCRIPTION,''),ITEMDESCRIPTION,ISNULL(ADDITIONALDESCRIPTION,''),ITEMCODE FROM INV_M_ITEMMASTER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEM")
			newElem.setAttribute "ITMCODE", trim(dcrs(0))
			newElem.setAttribute "ITMSHDESC", trim(dcrs(1))
			newElem.setAttribute "ITMDESC", trim(dcrs(2))
			newElem.setAttribute "ITMADDDESC", trim(dcrs(3))
			newElem.setAttribute "ITM", trim(dcrs(4))
			Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	if sOrgCode <> "N" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT CLASSIFICATIONCODE,ISNULL(SHORTDESCRIPTION,''),ITEMDESCRIPTION,ISNULL(ADDITIONALDESCRIPTION,''),ITEMCODE FROM INV_M_ITEMORGMASTER WHERE ORGANISATIONCODE = '" & sOrgCode & "'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			do while not dcrs.EOF
				Set newElem = OutData.createElement("ITEM")
				newElem.setAttribute "CITMCODE", trim(dcrs(0))
				newElem.setAttribute "CITMSHDESC", trim(dcrs(1))
				newElem.setAttribute "CITMDESC", trim(dcrs(2))
				newElem.setAttribute "CITMADDDESC", trim(dcrs(3))
				newElem.setAttribute "CITM", trim(dcrs(4))
				Root.appendChild newElem
			dcrs.MoveNext
			loop
		end if
		dcrs.Close
	end if

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT GENITEMCODE,ISNULL(SHORTDESCRIPTION,''),ITEMDESCRIPTION,ISNULL(ADDITIONALDESCRIPTION,''),TEMPITEMCODE,FINALSTATUS FROM MS_TEMPORARYITEMMASTER"' WHERE FINALSTATUS = 'N'"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		do while not dcrs.EOF
			Set newElem = OutData.createElement("ITEM")
			newElem.setAttribute "TITMCODE", trim(dcrs(0))
			newElem.setAttribute "TITMSHDESC", trim(dcrs(1))
			newElem.setAttribute "TITMDESC", trim(dcrs(2))
			newElem.setAttribute "TITMADDDESC", trim(dcrs(3))
			newElem.setAttribute "TITM", trim(dcrs(4))
			newElem.setAttribute "TSTATUS", trim(dcrs(5))
			if trim(dcrs(5)) = "Y" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT DISTINCT ITEMDESCRIPTION FROM INV_M_ITEMMASTER WHERE ITEMCODE = (SELECT DISTINCT ITEMCODE FROM MS_TEMPFINALITEMDETAIL WHERE TEMPITEMCODE = " & trim(dcrs(4)) & ")"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
						newElem.setAttribute "TDESC", trim(dcrs1(0))
				end if
				dcrs1.Close
			end if

			Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>
