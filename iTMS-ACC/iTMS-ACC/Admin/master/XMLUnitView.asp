<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLUnitView.asp
	'Module Name				:	Inventory (Organization View)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 09, 2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view for Organization)
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
<html>
<title>Organization Units</title>
<head>
<SCRIPT SRC="../scripts/XMLTreeView.js"></SCRIPT>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css"/>
</head>
<BODY leftMargin="20" topMargin="10" MARGINHEIGHT="0" MARGINWIDTH="0"></BODY>
</html>
<%
	dim dcrs,dcrs1,dcrs2,dcrs3,sOrgId,sOrgName,sGCode,sOrgUnitID,sOrgUnitName
	dim sOrgUnitDesID,sOrgUnitDesName,sLocID,sLocName,sBinID,sBinName,sTempOrgUnitDesID,sTempOrgUnitDesName
	dim OutData,Root,newElem,newElem1,newElem2,newElem3
	dim iDivCounter
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")

	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
	set dcrs3 = Server.CreateObject("ADODB.Recordset")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGANIZATIONID,ORGANIZATIONNAME FROM DCS_ORGANIZATION ORDER BY ORGANIZATIONID"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	set sOrgId = dcrs(0)
	set sOrgName = dcrs(1)

	iDivCounter = 0

	if not dcrs.EOF then
		do while not dcrs.EOF
			' ORGANIZATION
			iDivCounter = cint(iDivCounter) + 1
			Set newElem = OutData.createElement("menu")
			newElem.setAttribute "icon1", "../../assets/images/home.gif"
			newElem.setAttribute "caption", trim(sOrgName)
			newElem.setAttribute "DIVID", iDivCounter
			newElem.setAttribute "Description", "ORGANIZATION"
	
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ORGANIZATIONUNITID,ORGANIZATIONUNITNAME FROM DCS_ORGANIZATIONUNITS WHERE ORGANIZATIONID = '" & trim(sOrgId) & "' AND ORGANIZATIONUNITID = 1"
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing
			set sOrgUnitID = dcrs1(0)
			set sOrgUnitName = dcrs1(1)

			Do While Not dcrs1.EOF
				' DIVISION
				iDivCounter = cint(iDivCounter) + 1
				Set newElem1 = OutData.createElement("menuItem")
				newElem1.setAttribute "caption", trim(sOrgUnitName)
				newElem1.setAttribute "DIVID", iDivCounter
				newElem1.setAttribute "icon1", "../../assets/images/folder-closed.gif"
				newElem1.setAttribute "Description", trim(ucase(sOrgUnitName))

				with dcrs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = " & trim(sOrgUnitID) & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs2.ActiveConnection = nothing
				set sTempOrgUnitDesID = dcrs2(0)
				set sTempOrgUnitDesName = dcrs2(1)

				Do While Not dcrs2.EOF
					iDivCounter = cint(iDivCounter) + 1
					Set newElem2 = OutData.createElement("menuItem")
					newElem2.setAttribute "caption", trim(sTempOrgUnitDesName)
					newElem2.setAttribute "opened", "false"
					newElem2.setAttribute "icon1", "../../assets/images/folder-closed.gif"
					newElem2.setAttribute "DIVID", iDivCounter
					newElem2.setAttribute "Description", trim(ucase(sOrgUnitName))

					with dcrs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE (LEFT(OUDEFINITIONID,4) = '" & trim(sTempOrgUnitDesID) & "') AND OUDEFINITIONID <> '" & trim(sTempOrgUnitDesID) & "'"
						.ActiveConnection = con
						.Open
					end with
					set dcrs3.ActiveConnection = nothing
					set sOrgUnitDesID = dcrs3(0)
					set sOrgUnitDesName = dcrs3(1)

					Do While Not dcrs3.EOF
						iDivCounter = cint(iDivCounter) + 1
						Set newElem3 = OutData.createElement("menuItem")
						newElem3.setAttribute "caption", trim(sOrgUnitDesName)
						newElem3.setAttribute "opened", "false"
						newElem3.setAttribute "icon1", "../../assets/images/file.gif"
						newElem3.setAttribute "DIVID", iDivCounter
						newElem3.setAttribute "Description", "UNIT"

					newElem2.appendChild newElem3
					dcrs3.MoveNext
					Loop
					dcrs3.Close

				newElem1.appendChild newElem2
				dcrs2.MoveNext
				Loop
				dcrs2.Close

			newElem.appendChild newElem1
			dcrs1.MoveNext
			Loop
			dcrs1.Close
			
		OutData.appendChild newElem
		dcrs.MoveNext
		Loop
	else
		Set newElem = OutData.createElement("menu")
		newElem.setAttribute "icon1", "../../assets/images/home.gif"
		newElem.setAttribute "caption", "Data Unavailable"
		OutData.appendChild newElem
	end if
	dcrs.Close
	
	dim xsl,xslt,xslProc
	
	set xsl = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
	xsl.async = False
	xsl.Load (Server.MapPath("../xmldata/DisplayDetails.xsl"))

	Set xslt = Server.CreateObject("Msxml2.XSLTemplate")
	xslt.stylesheet = xsl
	set xslProc = xslt.createProcessor()
	xslProc.input = OutData

	'APPLY THE TRANSFORMATION AND WRITE THE RESULTS AS OUTPUT
	xslProc.transform()
	Response.Write(xslProc.output)
%>

