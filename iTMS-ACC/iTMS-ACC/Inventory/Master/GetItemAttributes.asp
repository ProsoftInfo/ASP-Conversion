<%@ Language=VBScript %>
<%	option explicit	%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
	Dim dcrs,oDOM,Node,Root,sOrgID
	sOrgID = Session("organizationcode")

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	set oDOM = server.CreateObject("Microsoft.XMLDOM")

	set Root = oDOM.createElement("ROOT")
	oDOM.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ITEMTYPEID, ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.Source = "SELECT ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME,ClassificationCode FROM INV_M_ITEMTYPEATTRIBUTES"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM.createElement("ATTRIBUTES")
		'Node.setAttribute "ITEMTYPEID",trim(dcrs(0))
		Node.setAttribute "ATTRID",trim(dcrs(0))
		Node.setAttribute "ATTRNAME",trim(dcrs(1))
		Node.setAttribute "ClassCode",trim(dcrs(2))
		Root.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close
	Response.ContentType = "text/xml"
	Response.Write oDOM.xml
%>
