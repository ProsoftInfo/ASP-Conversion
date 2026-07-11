<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgCreationInsert.asp	
	'Module Name				:	Inventory (Organization Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 11, 2002
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
Dim oDOM,newElem,Root

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

dim dcrs,iCount,sSql
dim sorgID,sorgName,sorgShName,iorgNoUnits,sorgAddr1,sorgAddr2,sorgPIN,sorgCity
dim sorgState,norgCountry,sorgCountry,sorgPhone,sorgFax,sorgEmail,sorgURL,sorgContactPerson
dim norgCurrency,sorgTNGSTRCNo,sorgAreaCode,sorgCSTRCNo,dorgCSTRCDate

sorgID = trim(Request.Form("hOrgID"))
sorgName = trim(Request.Form("txtOrgName"))
sorgShName = trim(Request.Form("txtOrgShortName"))
iorgNoUnits = trim(Request.Form("txtOrgNoUnits"))
sorgAddr1 = trim(Request.Form("txtOrgAddress1"))
sorgAddr2 = trim(Request.Form("txtOrgAddress2"))
sorgPIN = trim(Request.Form("txtOrgPIN"))
sorgCity = trim(Request.Form("txtOrgCity"))
sorgState = trim(Request.Form("txtOrgState"))
norgCountry = trim(Request.Form("hcountryValue"))
sorgCountry = trim(Request.Form("hcountry"))
sorgPhone = trim(Request.Form("txtOrgPhone"))
sorgFax = trim(Request.Form("txtOrgFax"))
sorgEmail = trim(Request.Form("txtOrgEmail"))
sorgURL = trim(Request.Form("txtOrgURL"))
sorgContactPerson = trim(Request.Form("txtOrgContactPerson"))
norgCurrency = trim(Request.Form("hcurrency"))
sorgTNGSTRCNo = trim(Request.Form("txtOrgTNGSTRCNo"))
sorgAreaCode = trim(Request.Form("txtOrgAreaCode"))
sorgCSTRCNo = trim(Request.Form("txtOrgCSTRCNo"))
dorgCSTRCDate = trim(Request.Form("txtOrgCSTRCDate"))

con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT COUNT(ORGANIZATIONID) FROM DCS_ORGANIZATION"
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
set iCount = dcrs(0)
if iCount = 0 then
	sSql = "INSERT INTO DCS_ORGANIZATION (ORGANIZATIONID,ORGANIZATIONNAME,ORGANIZATIONSHORTNAME," &_
		"ORGANIZATIONUNITS,OPERATINGCURRENCY,ADDRESS1,ADDRESS2,POSTCODE,CITY,STATE,COUNTRY," &_
		"PHONENUMBER,FAXNUMBER,EMAILID,WESITEURL,CONTACTPERSON,TNGSTRCNUMBER,CSTRCNUMBER," &_
		"CSTRCDATE,AREACODE) VALUES " &_
		"('01'," & Pack(ucase(sorgName)) & "," & Pack(sorgShName) & "," & iorgNoUnits & ", " &_
		" " & norgCurrency & "," & Pack(sorgAddr1) & "," & Pack(sorgAddr2) & ", " &_
		" " & Pack(sorgPIN) & "," & Pack(sorgCity) & "," & Pack(sorgState) & ", " &_
		" " & norgCountry & "," & Pack(sorgPhone) & "," & Pack(sorgFax) & ", " &_
		" " & Pack(sorgEmail) & "," & Pack(sorgURL) & "," & Pack(sorgContactPerson) & ", " &_
		" " & Pack(sorgTNGSTRCNo) & "," & Pack(sorgCSTRCNo) & ", " &_
		" CONVERT(DATETIME," & Pack(dorgCSTRCDate) & ",103)," & Pack(sorgAreaCode) & ")"
	'Response.Write sSql & "<BR>"
	con.Execute sSql

	Set Root = oDOM.createElement("OrganizationDetails")
	oDOM.appendChild Root
	Set newElem = oDOM.createElement("Organization")
	newElem.setAttribute "ORGANIZATIONID", "01"
	newElem.setAttribute "ORGANIZATIONNAME", ucase(sorgName)
	newElem.setAttribute "ORGANIZATIONSHORTNAME", sorgShName
	newElem.setAttribute "ORGANIZATIONUNITS", iorgNoUnits
	newElem.setAttribute "OPERATINGCURRENCY", norgCurrency
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
	Root.appendChild newElem
	oDOM.Save server.MapPath("../xmldata/Organization.xml")

%>
	<BODY BGCOLOR="#336699" onLoad = "msgbox('Organization <%=replace(sorgName,"'","\'")%> has been Created Successfully','Y')">
<%
else
	sSql = "UPDATE DCS_ORGANIZATION SET ORGANIZATIONNAME = " & Pack(ucase(sorgName)) & ",ORGANIZATIONSHORTNAME = " & Pack(sorgShName) & "," &_
		"ORGANIZATIONUNITS = " & iorgNoUnits & ",OPERATINGCURRENCY = " & norgCurrency & "," &_
		"ADDRESS1 = " & Pack(sorgAddr1) & ",ADDRESS2 = " & Pack(sorgAddr2) & ",POSTCODE = " & Pack(sorgPIN) & "," &_
		"CITY = " & Pack(sorgCity) & ",STATE = " & Pack(sorgState) & ",COUNTRY = " & norgCountry & "," &_
		"PHONENUMBER = " & Pack(sorgPhone) & ",FAXNUMBER = " & Pack(sorgFax) & ",EMAILID = " & Pack(sorgEmail) & "," &_
		"WESITEURL = " & Pack(sorgURL) & ",CONTACTPERSON = " & Pack(sorgContactPerson) & "," &_
		"TNGSTRCNUMBER = " & Pack(sorgTNGSTRCNo) & ",CSTRCNUMBER = " & Pack(sorgCSTRCNo) & "," &_
		"CSTRCDATE = CONVERT(DATETIME," & Pack(dorgCSTRCDate) & ",103),AREACODE = " & Pack(sorgAreaCode) & " WHERE ORGANIZATIONID = " & Pack(sorgID) & ""
	con.Execute sSql

	Set Root = oDOM.createElement("OrganizationDetails")
	oDOM.appendChild Root
	Set newElem = oDOM.createElement("Organization")
	newElem.setAttribute "ORGANIZATIONID", sorgID
	newElem.setAttribute "ORGANIZATIONNAME", ucase(sorgName)
	newElem.setAttribute "ORGANIZATIONSHORTNAME", sorgShName
	newElem.setAttribute "ORGANIZATIONUNITS", iorgNoUnits
	newElem.setAttribute "OPERATINGCURRENCY", norgCurrency
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
	Root.appendChild newElem
	oDOM.Save server.MapPath("../xmldata/Organization.xml")

%>
	<BODY onLoad = "msgbox('Organization Updated Sucessfully','Y')">
<%
end if
dcrs.Close

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
