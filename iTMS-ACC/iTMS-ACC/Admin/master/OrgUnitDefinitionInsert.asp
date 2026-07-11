<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgUnitDefinitionInsert.asp
	'Module Name				:	Inventory (Organization Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 13, 2002
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
<SCRIPT>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to define another Unit"))
				window.location.href = "OrgUnitDefinitionEntry.asp"
			else
				window.location.href = "OrgCreationEntry.asp"
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
Dim oDOM,newElem,Root,objfs,adoCmd

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,iCount,sSql,sunitDefId
dim sorgID,iorgUnitID,sorgParUnitID,sorgName,sorgShName,sorgAddr1,sorgAddr2,sorgPIN,sorgCity
dim sorgState,norgCountry,sorgCountry,sorgPhone,sorgFax,sorgEmail,sorgURL,sorgContactPerson
dim sorgTNGSTRCNo,sorgAreaCode,sorgCSTRCNo,dorgCSTRCDate,sorgRange
dim sorgDivision,sorgCollectorate,sorgCtrlExNo,sorgRegNo,sorgLANo
dim stempsorgID,sorgRangeAdd1,sorgRangeAdd2,sorgRangeAdd3
dim sorgDivisionAdd1,sorgDivisionAdd2,sorgDivisionAdd3

sorgID = trim(Request.Form("selOrg"))
iorgUnitID = trim(Request.Form("selOrgUnit"))
sorgParUnitID = trim(Request.Form("selParOrgUnit"))
sorgName = trim(Request.Form("txtUnitName"))
sorgShName = trim(Request.Form("txtUnitShName"))
sorgAddr1 = trim(Request.Form("txtUnitAddr1"))
sorgAddr2 = trim(Request.Form("txtUnitAddr2"))
sorgPIN = trim(Request.Form("txtUnitPIN"))
sorgCity = trim(Request.Form("txtUnitCity"))
sorgState = trim(Request.Form("txtUnitState"))
norgCountry = trim(Request.Form("selUnitCountry"))
sorgCountry = trim(Request.Form("hcountry"))
sorgPhone = trim(Request.Form("txtUnitPhone"))
sorgFax = trim(Request.Form("txtUnitFax"))
sorgEmail = trim(Request.Form("txtUnitEmail"))
sorgURL = trim(Request.Form("txtUnitURL"))
sorgContactPerson = trim(Request.Form("txtUnitContactPerson"))
sorgTNGSTRCNo = trim(Request.Form("txtUnitTNGSTNo"))
sorgAreaCode = trim(Request.Form("txtUnitAreaCode"))
sorgCSTRCNo = trim(Request.Form("txtUnitCSTRCNo"))
dorgCSTRCDate = trim(Request.Form("txtUnitCSTRCDate"))
sorgRange = trim(Request.Form("txtUnitRange"))
sorgDivision = trim(Request.Form("txtUnitDivision"))
sorgCollectorate = trim(Request.Form("txtUnitCollectorate"))
sorgCtrlExNo = trim(Request.Form("txtUnitCentralENo"))
sorgRegNo = trim(Request.Form("txtUnitRegNo"))
sorgLANo = trim(Request.Form("txtUnitLANo"))
sorgRangeAdd1 = trim(Request.Form("txtRangeAdd1"))
sorgRangeAdd2 = trim(Request.Form("txtRangeAdd2"))
sorgRangeAdd3 = trim(Request.Form("txtRangeAdd3"))
sorgDivisionAdd1 = trim(Request.Form("txtDivisionAdd1"))
sorgDivisionAdd2 = trim(Request.Form("txtDivisionAdd2"))
sorgDivisionAdd3 = trim(Request.Form("txtDivisionAdd3"))

con.beginTrans

if isnull(sorgParUnitID) or IsEmpty(sorgParUnitID) or sorgParUnitID = "" then
	stempsorgID = sorgID
	if objfs.FileExists(Server.MapPath("../xmldata/Division.xml")) then
		oDOM.Load server.MapPath("../xmldata/Division.xml")
		Set Root = oDOM.documentElement
	else
		Set Root = oDOM.createElement("Root")
		oDOM.appendChild Root
	end if
	Set newElem = oDOM.createElement("Division")
else
	stempsorgID = sorgParUnitID
	if objfs.FileExists(Server.MapPath("../xmldata/Unit.xml")) then
		oDOM.Load server.MapPath("../xmldata/Unit.xml")
		Set Root = oDOM.documentElement
	else
		Set Root = oDOM.createElement("Root")
		oDOM.appendChild Root
	end if
	Set newElem = oDOM.createElement("Unit")
end if

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ISNULL(MAX(SUBSTRING(OUDEFINITIONID," & Cint(iorgUnitID) * 2 + 1 & ",248)),0)+ 1 FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = " & iorgUnitID & " AND LEFT(OUDEFINITIONID," & Cint(iorgUnitID)* 2 & " ) = '"&stempsorgID&"'"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
set iCount = dcrs(0)

if iCount < 10 then
	sunitDefId = trim(stempsorgID & "0" & trim(iCount))
Else
	sunitDefId = trim(stempsorgID & trim(iCount))
End If
dcrs.Close

with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LOWER(ORGUNITDESCRIPTION) = " & Pack(lcase(sorgName)) & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if dcrs.EOF then
	sSql = "INSERT INTO DCS_ORGANIZATIONUNITDEFINITIONS (OUDEFINITIONID,ORGANIZATIONUNITID,ORGUNITDESCRIPTION,ORGUNITSHORTDESCRIPTION," &_
		"ADDRESS1,ADDRESS2,POSTCODE,CITY,STATE,COUNTRY,PHONENUMBER,FAXNUMBER,EMAILID,WESITEURL,CONTACTPERSON," &_
		"TNGSTRCNUMBER,CSTRCNUMBER,CSTRCDATE,AREACODE,RANGE,DIVISION,COLLECTORATE," &_
		"CENTRALEXCISECODE,REGISTRATIONNUMBER,LANUMBER,RANGEADDRESS1,RANGEADDRESS2," &_
		"RANGEADDRESS3,DIVADDRESS1,DIVADDRESS2,DIVADDRESS3) VALUES " &_
		"(" & Pack(sunitDefId) & "," & iorgUnitID & "," & Pack(ucase(sorgName)) & "," & Pack(sorgShName) & ", " &_
		" " & Pack(sorgAddr1) & "," & Pack(sorgAddr2) & ", " &_
		" " & Pack(sorgPIN) & "," & Pack(sorgCity) & "," & Pack(sorgState) & ", " &_
		" " & norgCountry & "," & Pack(sorgPhone) & "," & Pack(sorgFax) & ", " &_
		" " & Pack(sorgEmail) & "," & Pack(sorgURL) & "," & Pack(sorgContactPerson) & ", " &_
		" " & Pack(sorgTNGSTRCNo) & "," & Pack(sorgCSTRCNo) & ", " &_
		" CONVERT(DATETIME," & Pack(dorgCSTRCDate) & ",103)," & Pack(sorgAreaCode) & ", " &_
		" " & Pack(sorgRange) & "," & Pack(sorgDivision) & "," & Pack(sorgCollectorate) & ", " &_
		" " & Pack(ucase(sorgCtrlExNo)) & "," & Pack(ucase(sorgRegNo)) & "," & Pack(ucase(sorgLANo)) & ", " &_
		" " & Pack(sorgRangeAdd1) & "," & Pack(sorgRangeAdd2) & "," & Pack(sorgRangeAdd3) & ", " &_
		" " & Pack(sorgDivisionAdd1) & "," & Pack(sorgDivisionAdd2) & "," & Pack(sorgDivisionAdd3) & ")"
	'Response.Write sSql & "<BR>"
	con.Execute sSql


	newElem.setAttribute "OUDEFINITIONID", sunitDefId
	newElem.setAttribute "ORGANIZATIONUNITID", iorgUnitID
	newElem.setAttribute "ORGUNITDESCRIPTION", ucase(sorgName)
	newElem.setAttribute "ORGUNITSHORTDESCRIPTION", sorgShName
	newElem.setAttribute "ADDRESS1", sorgAddr1
	newElem.setAttribute "ADDRESS2", sorgAddr2
	newElem.setAttribute "POSTCODE", sorgPIN
	newElem.setAttribute "CITY", sorgCity
	newElem.setAttribute "STATE", sorgState
	newElem.setAttribute "COUNTRYCODE", norgCountry
	newElem.setAttribute "COUNTRYNAME", sorgCountry
	newElem.setAttribute "PHONENUMBER", sorgPhone
	newElem.setAttribute "FAXNUMBER", sorgFax
	newElem.setAttribute "EMAILID", sorgEmail
	newElem.setAttribute "WESITEURL", sorgURL
	newElem.setAttribute "CONTACTPERSON", sorgContactPerson
	newElem.setAttribute "TNGSTRCNUMBER", sorgTNGSTRCNo
	newElem.setAttribute "CSTRCNUMBER", sorgCSTRCNo
	newElem.setAttribute "CSTRCDATE", dorgCSTRCDate
	newElem.setAttribute "AREACODE", sorgAreaCode
	newElem.setAttribute "RANGE", sorgRange
	newElem.setAttribute "DIVISION", sorgDivision
	newElem.setAttribute "COLLECTORATE", sorgCollectorate
	newElem.setAttribute "CENTRALEXCISECODE", ucase(sorgCtrlExNo)
	newElem.setAttribute "REGISTRATIONNUMBER", ucase(sorgRegNo)
	newElem.setAttribute "LANUMBER", ucase(sorgLANo)
	newElem.setAttribute "RANGEADDRESS1", sorgRangeAdd1
	newElem.setAttribute "RANGEADDRESS2", sorgRangeAdd2
	newElem.setAttribute "RANGEADDRESS3", sorgRangeAdd3
	newElem.setAttribute "DIVISIONADDRESS1", sorgDivisionAdd1
	newElem.setAttribute "DIVISIONADDRESS2", sorgDivisionAdd2
	newElem.setAttribute "DIVISIONADDRESS3", sorgDivisionAdd3
	newElem.setAttribute "PARTYCODE", ""
	newElem.setAttribute "PARTYMAPPEDBY", ""
	newElem.setAttribute "PARTYMAPPEDON", ""

	Root.appendChild newElem

	if isnull(sorgParUnitID) or IsEmpty(sorgParUnitID) or sorgParUnitID = "" then
		oDOM.Save server.MapPath("../xmldata/Division.xml")
	else
		oDOM.Save server.MapPath("../xmldata/Unit.xml")
	end if

	'blocked
	''Declaration of Objects
	'Set adoCmd = Server.CreateObject("ADODB.Command")
	'Set adoCmd.ActiveConnection = con

	'sSql = "mapUnitMaster"

	'adoCmd.CommandText = sSql
	'adoCmd.CommandType = 4

	'adoCmd.Parameters.Refresh
	'adoCmd.Parameters.Item(1).Value = sunitDefId
	'adoCmd.Parameters.Item(2).Value = sorgShName

	'Set dcrs = adoCmd.Execute()

%>
	<BODY onLoad = "msgbox('Organization Unit <%=replace(sorgName,"'","\'")%> has been Created Successfully','Y')">
<%
else
%>
	<BODY onLoad = "msgbox('Organization Unit <%=replace(sorgName,"'","\'")%> already created','N')">
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
