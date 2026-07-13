<%@ Language=VBScript %>
<%	option explicit	%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
	Dim dcrs,oDOM,Node,Root,sOrgID,sCategoryCode
	sOrgID = Session("organizationcode")
	sCategoryCode = Replace(Request("Code"),",","','")

	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	set oDOM = server.CreateObject("Microsoft.XMLDOM")

	set Root = oDOM.createElement("ROOT")
	oDOM.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select CategoryCode,CategoryName from Inv_M_ClassificationCategory where CategoryCode in('"& sCategoryCode &"')"
		'.Source = "SELECT ITEMTYPEID, ITEMTYPEATTRIBUTEID, ITEMTYPEATTRIBUTENAME FROM INV_M_ITEMTYPEATTRIBUTES ORDER BY 2"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	do while Not dcrs.EOF
		set Node = oDOM.createElement("CATEGORY")
		Node.setAttribute "CODE",trim(dcrs(0))
		Node.setAttribute "NAME",trim(dcrs(1))
		Root.appendChild Node
		dcrs.movenext
	loop
	dcrs.Close
	Response.ContentType = "text/xml"
	Response.Write oDOM.xml
%>
