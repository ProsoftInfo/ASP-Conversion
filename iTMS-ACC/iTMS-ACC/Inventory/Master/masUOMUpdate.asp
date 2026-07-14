<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	masUOMUpdate.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 17, 2003
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
<!--#include virtual="/include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to Amend another Unit Of Measurement")) 
				window.location.href = "MasUOMAmendEntry.asp"
			else
				window.location.href = "MasCategoryAmendEntry.asp"
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
Dim oDOM,newElem,Root,objfs,tempNode,bFlag

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql
dim sUOMCode,sUOMName,sUOmShName,arrTemp,sDecimalAllowed

arrTemp = split(trim(Request.Form("selUoM")),"|")
sUOMCode = trim(arrTemp(0))
sDecimalAllowed = trim(Request.Form("radDecimal"))

sUOMName = trim(Request.Form("txtUOMName"))
sUOmShName = sUOMCode

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT UOMCODE FROM MS_UNITOFMEASUREMENT WHERE LOWER(UOMDESCRIPTION) = " & Pack(lcase(sUOMName)) & " AND DECIMALALLOWED = " & Pack(sDecimalAllowed) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if dcrs.EOF then
	sSql = "UPDATE MS_UNITOFMEASUREMENT SET UOMDESCRIPTION = " & Pack(ucase(sUOMName)) & "," &_
		"UOMSHORTDESCRIPTION = " & Pack(ucase(sUOmShName)) & ",DECIMALALLOWED = " & Pack(sDecimalAllowed) & " " &_
		"WHERE UOMCODE = " & Pack(sUOMCode) & ""
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
	bFlag = true
	if objfs.FileExists(Server.MapPath("../xmldata/UoM.xml")) then
		oDOM.Load server.MapPath("../xmldata/UoM.xml")
		Set Root = oDOM.documentElement
		for each tempNode in Root.childNodes
			if ucase(tempNode.attributes.getNamedItem("UOMCODE").value) = ucase(sUOMCode) then
				tempNode.attributes.getNamedItem("UOMDESCRIPTION").value = ucase(sUOMName)
				tempNode.attributes.getNamedItem("UOMSHORTDESCRIPTION").value = ucase(sUOmShName)
				tempNode.attributes.getNamedItem("DECIMALSALLOWED").value = sDecimalAllowed
				bFlag = false
				exit for
			end if
		next
	else	
		Set Root = oDOM.createElement("Root")
		oDOM.appendChild Root
		bFlag = true
	end if
	if bFlag then
		Set newElem = oDOM.createElement("UoM")
		newElem.setAttribute "UOMCODE", ucase(sUOMCode)
		newElem.setAttribute "UOMDESCRIPTION", ucase(sUOMName)
		newElem.setAttribute "UOMSHORTDESCRIPTION", ucase(sUOmShName)
		newElem.setAttribute "DECIMALSALLOWED", sDecimalAllowed
		Root.appendChild newElem
	end if
	
	oDOM.Save server.MapPath("../xmldata/UoM.xml")
%>
	<BODY onLoad = "msgbox('Unit Of Measurement <%=replace(sUOMName,"'","\'")%> has been Updated Successfully','Y')">
<%
else
%>
	<BODY onLoad = "msgbox('Unit Of Measurement <%=replace(sUOMName,"'","\'")%> has already created','N')">
<%
End If

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
