<%@ Language=VBScript %>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetCategoryGroup.asp	
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view for Classification)
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
<title>Item</title>
<head>
<SCRIPT LANGUAGE=javascript SRC="../../Common/XMLTreeView.js"></SCRIPT>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css"/>
</head>
<BODY leftMargin="20" topMargin="10" MARGINHEIGHT="0" MARGINWIDTH="0"></BODY>
</html>

<%
	dim dcrs,dcrs1
	dim iDivCounter
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")

	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	Set newElem = OutData.createElement("menu")
	newElem.setAttribute "icon1", "../../assets/images/home.gif"
	newElem.setAttribute "caption", "CLASS"
	newElem.setAttribute "Description", "CLASS"
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME,CHILDCOUNT,PARENTGROUP FROM INV_M_CLASSIFICATION ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	Do While Not dcrs.EOF

		sGCode = dcrs(0)
		sGName = dcrs(1)
		sGChildCount = dcrs(2)
		sPGroup = dcrs(3)
        stempkey = Trim(sGCode)
        iDivCounter=CDbl(iDivCounter)+1
        
		If stempkey = Trim(sPGroup) Then
			set newElem1 = OutData.createElement("menuItem")
			newElem1.setAttribute "caption", trim(ucase(sGName))
			newElem1.setAttribute "icon1", "../../assets/images/folder-closed.gif"
			newElem1.setAttribute "opened","false"
			newElem1.setAttribute "DIVID", iDivCounter
			newElem1.setAttribute "Description", ""
			newElem.appendChild newElem1
			
            stempkey = Trim(sGCode)
            child stempkey,newElem1

		end if
	
	dcrs.MoveNext
	Loop
	dcrs.Close
			
	OutData.appendChild newElem
	
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

<%
Private Sub child(sid,oNodParent)
    Dim sptempkey, sctempkey,dcrs1
    dim stGCode,stGName

	set dcrs1 = Server.CreateObject("ADODB.Recordset")
    
	with dcrs1
		.Source = "SELECT GROUPCODE,GROUPNAME,CHILDCOUNT,PARENTGROUP FROM INV_M_CLASSIFICATION WHERE PARENTGROUP = " & sid & " AND GROUPCODE <> " & sid & " ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with

    Do While Not dcrs1.EOF
		stGCode = dcrs1(0)
		stGName = dcrs1(1)
		iDivCounter=CDbl(iDivCounter)+1
		set newElem2 = OutData.createElement("menuItem")
		newElem2.setAttribute "caption", trim(ucase(stGName))
		newElem2.setAttribute "opened","false"
		newElem2.setAttribute "icon1", "../../assets/images/folder-closed.gif"
		newElem2.setAttribute "DIVID", iDivCounter
		newElem2.setAttribute "Description","" 
        oNodParent.appendChild(newElem2)
        child Trim(stGCode),newElem2
    dcrs1.MoveNext
    Loop
    dcrs1.Close
End Sub

%>
