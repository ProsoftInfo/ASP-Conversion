<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AmendOrgUnitDefinitionInsert.asp	
	'Module Name				:	Inventory (Organization Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 29, 2004
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
			if (confirm("Do You want to amend another Unit")) 
				window.location.href = "AmendOrgUnitDefinitionEntry.asp"
			else
				window.location.href = "../welcome_admin.asp"
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
	dim sorgID,iorgUnitID,sorgName,sorgShName,sorgAddr1,sorgAddr2,sorgPIN,sorgCity
	dim sorgState,norgCountry,sorgCountry,sorgPhone,sorgFax,sorgEmail,sorgURL,sorgContactPerson
	dim sorgTNGSTRCNo,sorgAreaCode,sorgCSTRCNo,dorgCSTRCDate,sorgRange
	dim sorgDivision,sorgCollectorate,sorgCtrlExNo,sorgRegNo,sorgLANo
	dim stempsorgID,sorgRangeAdd1,sorgRangeAdd2,sorgRangeAdd3
	dim sorgDivisionAdd1,sorgDivisionAdd2,sorgDivisionAdd3

	sorgID = trim(Request.Form("selOrg"))
	iorgUnitID = trim(Request.Form("selOrgUnit"))
	sunitDefId = trim(Request.Form("selDivisionUnit"))
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

	sSql = "UPDATE DCS_ORGANIZATIONUNITDEFINITIONS SET ORGUNITDESCRIPTION = " & Pack(ucase(sorgName)) _
	& ",ORGUNITSHORTDESCRIPTION = " & Pack(sorgShName) & ",ADDRESS1 = " & Pack(sorgAddr1) _
	& ",ADDRESS2 = " & Pack(sorgAddr2) & ",POSTCODE = " & Pack(sorgPIN) & ",CITY = " _
	& Pack(sorgCity) & ",STATE = " & Pack(sorgState) & ",COUNTRY = " & norgCountry _
	& ",PHONENUMBER = " & Pack(sorgPhone) & ",FAXNUMBER = " & Pack(sorgFax) _
	& ",EMAILID = " & Pack(sorgEmail) & ",WESITEURL = " & Pack(sorgURL) & ",CONTACTPERSON = " _
	& Pack(sorgContactPerson) & ",TNGSTRCNUMBER = " & Pack(sorgTNGSTRCNo) & ",CSTRCNUMBER = " _
	& Pack(sorgCSTRCNo) & ",CSTRCDATE = CONVERT(DATETIME," & Pack(dorgCSTRCDate) & ",103), " _
	& "AREACODE = " & Pack(sorgAreaCode) & ",RANGE = " & Pack(sorgRange) & ",DIVISION = " _
	& Pack(sorgDivision) & ",COLLECTORATE = " & Pack(sorgCollectorate) & ",CENTRALEXCISECODE = " _
	& Pack(ucase(sorgCtrlExNo)) & ",REGISTRATIONNUMBER = " & Pack(ucase(sorgRegNo)) _
	& ",LANUMBER = " & Pack(ucase(sorgLANo)) & ",RANGEADDRESS1 = " & Pack(sorgRangeAdd1) _
	& ",RANGEADDRESS2 = " & Pack(sorgRangeAdd2) & ",RANGEADDRESS3 = " & Pack(sorgRangeAdd3) _
	& ",DIVADDRESS1  = " & Pack(sorgDivisionAdd1) & ",DIVADDRESS2 = " & Pack(sorgDivisionAdd2) _
	& ",DIVADDRESS3 = " & Pack(sorgDivisionAdd3) & " WHERE OUDEFINITIONID = " & Pack(sunitDefId) _
	& " AND ORGANIZATIONUNITID = " & iorgUnitID

	con.Execute sSql
	
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

	if iorgUnitID = "1" then
		server.Execute ("XMLDivisionDefault.asp")
	else
		server.Execute ("XMLUnitDefault.asp")
	end if

%>
	<BODY onLoad = "msgbox('Organization Unit <%=replace(sorgName,"'","\'")%> has been Amended Successfully','Y')">

