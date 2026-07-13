<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasUOMInsert.asp	
	'Module Name				:	Inventory (Master Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 16, 2002
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,sAppCode) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to define another Unit Of Measurement")) 
			{
				window.location.href = "MasUOMEntry.asp?AppCode="  + sAppCode
			}	
			else
			{
				if (sAppCode == 6)
					window.location.href = "../../Production/welcome_Production.asp"
				else
					window.location.href = "MasClassificationEntry.asp"
			}	
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
Dim oDOM,newElem,Root,objfs

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,sSql
dim sUOMCode,sUOMName,sDecimalAllowed,sUOmShName,sAppCode

sUOMCode = trim(Request.Form("txtUOMCode"))
sUOMName = trim(Request.Form("txtUOMName"))
sAppCode = trim(Request.Form("hdAppCode"))
sDecimalAllowed = trim(Request.Form("radDecimal"))
sUOmShName = sUOMCode

Set dcrs = Server.CreateObject("ADODB.RecordSet")

with dcrs
	.Source = "SELECT UOMCODE FROM MS_UNITOFMEASUREMENT WHERE LOWER(UOMCODE) = " & Pack(lcase(sUOMCode)) & " OR LOWER(UOMDESCRIPTION) = " & Pack(lcase(sUOMName)) & ""
	.ActiveConnection = con
	.Open
end with
if dcrs.EOF then
	sSql = "INSERT INTO MS_UNITOFMEASUREMENT (UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION,DECIMALALLOWED,NOOFDECIMALS) VALUES" &_
		"(" & Pack(ucase(sUOMCode)) & "," & Pack(ucase(sUOMName)) & "," & Pack(ucase(sUOmShName)) & "," & Pack(sDecimalAllowed) & ",0)"
	'Response.Write sSql & "<BR>"
	con.Execute sSql
	
	
	if objfs.FileExists(Server.MapPath("../xmldata/UoM.xml")) then
		oDOM.Load server.MapPath("../xmldata/UoM.xml")
		Set Root = oDOM.documentElement
	else	
		Set Root = oDOM.createElement("Root")
		oDOM.appendChild Root
	end if

	Set newElem = oDOM.createElement("UoM")

	newElem.setAttribute "UOMCODE", ucase(sUOMCode)
	newElem.setAttribute "UOMDESCRIPTION", ucase(sUOMName)
	newElem.setAttribute "UOMSHORTDESCRIPTION", ucase(sUOmShName)
	newElem.setAttribute "DECIMALSALLOWED", sDecimalAllowed
	Root.appendChild newElem
	
	oDOM.Save server.MapPath("../xmldata/UoM.xml")
%>
	<BODY onLoad = "msgbox('Unit Of Measurement <%=replace(sUOMName,"'","\'")%> has been Created Successfully','Y','<%=sAppCode%>')">
<%
else
%>
	<BODY onLoad = "msgbox('Unit Of Measurement Code <%=sUOMCode%> or Unit Of Measurement Name <%=replace(sUOMName,"'","\'")%> already created','N')">
<%
End If
con.close
set con = nothing
%>
