<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgStorageBinAmendInsert.asp
	'Module Name				:	Inventory (Storage Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 20, 2004
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 07,2011
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do you want to Delete another Storage Location")) 
				window.location.href = "STORELOCATIONS.asp"
			else
				window.location.href = "../welcome_Inventory.asp"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>

<%
'XML DOM Variables
Dim oDOM,RootNode,StoreNode,DeleteNode,oNode
dim sExp,sExp1

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

dim dcrs,sSql
dim sorgID,iSLNo

sOrgID = Session("organizationcode")
iSLNo = trim(Request.Form("hPara"))

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

if oDOM.Load(server.MapPath("../xmldata/Storage.xml")) then
	Set RootNode = oDOM.documentElement
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT LOCATIONNUMBER FROM INV_M_ITEMSTORAGE WHERE LOCATIONNUMBER = " & iSLNo & ""
		'Response.Write dcrs.Source
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if dcrs.EOF then
		sSql = "DELETE INV_M_STOREBINDETAILS WHERE LOCATIONNUMBER = " & iSLNo & " AND OUDEFINITIONID = " & Pack(sOrgID) & ""
		'Response.Write sSql & "<BR>"
		con.Execute sSql

		sSql = "DELETE INV_M_STORAGE WHERE LOCATIONNUMBER = " & iSLNo & " AND OUDEFINITIONID = " & Pack(sOrgID) & ""
		'Response.Write sSql & "<BR>"
		con.Execute sSql

		sExp ="//Organization [@OUDEFINITIONID = '"&sorgID&"']"

		Set StoreNode = RootNode.Selectnodes(sExp)
	
		sExp1 ="//Organization [@OUDEFINITIONID = '"&sorgID&"'] /Storage [ @LOCATIONNUMBER = "&iSLNo&"]"

		Set DeleteNode = RootNode.Selectnodes(sExp1)

		if DeleteNode.Length > 0 then
			Set oNode = StoreNode.Item(0).RemoveChild(DeleteNode.Item(0))
		end if			

		oDOM.Save server.MapPath("../xmldata/Storage.xml")
%>
	<BODY onLoad = "msgbox('Storage Location has been Deleted Successfully','Y')">
<%
	else
%>
	<BODY onLoad = "msgbox('Storage Location could not be deleted since Item exits.','N')">
<%

	end if
else
%>
	<BODY onLoad = "msgbox('XML File Not Found','N')">
<%

end if

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
