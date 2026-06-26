<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetPartyDet.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	July 23,2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Nov 20,2010
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
dim OutData,Root,newElem,sType,sType2

dim objRs,sQuery,iPartyCode,sUnit

Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")

iPartyCode = Request("PartyCode")
sQuery = "Select distinct OUDefinitionID From APP_R_OrgParty Where PartyCode = "&iPartyCode&" and isNull(Useable,'0') = '0' "


with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

set objRs.ActiveConnection = nothing
Do While Not objRs.EOF
	sUnit = sUnit&":"&objRs(0)
	objRs.MoveNext
Loop
objRs.Close

sUnit = Mid(sUnit,2)


sQuery = "Select Count(1) From Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader H "&_
		 "Where V.PartyCode = "&iPartyCode&" and H.PartyCode = "&iPartyCode&" "

'Response.Write sQuery

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

Set objRs.ActiveConnection = nothing
IF Not objRs.EOF Then
	sType = objRs(0)
End IF

objRs.Close

sQuery = "Select Distinct OUDefinitionID from Acc_T_VoucherHeader Where PartyCode = "&iPartyCode&" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

Set objRs.ActiveConnection = nothing
Do while not objRs.EOF
	sType = sType&":"&objRs(0)
	objRs.MoveNext
loop

objRs.Close

if CStr(sType) <> ":" Then
	sType = Mid(sType,2)
End IF

sQuery = "Select Distinct OUDefinitionID from Acc_T_CreatedVoucherHeader Where PartyCode = "&iPartyCode&" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

Set objRs.ActiveConnection = nothing
Do while not objRs.EOF
	sType2 = sType2&":"&objRs(0)
	objRs.MoveNext
loop

objRs.Close

if CStr(sType2) <> ":" Then
	sType2 = Mid(sType2,2)
End IF

sType = sType&":"&sType2


sQuery="SELECT PartyCode, OrgnPartyCode, PartyName, isNull(AddressLine1,''), isNull(AddressLine2,''),"&_
 "isNull(City,''), isNull(State,''), isNull(Country,''), isNull(PhoneNos,''), isNull(MobileNos,''), isNull(FaxNos,''), isNull(Email,''), "&_
 "isNull(WebsiteURL,''), isNull(Pincode,''), isNull(ExciseControlCode,''), isNull(LocalSTNoandDT,''), isNull(CentralSTNoandDT,''), isNull(IncomeTaxPANNo,''),"&_
 " isNull(PartyGroupCoyType,''),isNull(TINNumber,''),Useable FROM APP_M_PartyMaster where PartyCode="&iPartyCode

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	set objRs.ActiveConnection = nothing

	Set Root = OutData.createElement("Root")
	OutData.appendChild Root

	if not objRs.EOF then
		do while not objRs.EOF
			Set newElem = OutData.createElement("Party")
			newElem.setAttribute "OrgnPartyCode", objRs(1)
			newElem.setAttribute "AddressLine1", objRs(3)
			newElem.setAttribute "AddressLine2", objRs(4)
			newElem.setAttribute "City", objRs(5)
			newElem.setAttribute "State", objRs(6)
			newElem.setAttribute "Country", objRs(7)
			newElem.setAttribute "PhoneNos", objRs(8)
			newElem.setAttribute "MobileNos", objRs(9)
			newElem.setAttribute "FaxNos", objRs(10)
			newElem.setAttribute "Email", objRs(11)
			newElem.setAttribute "WebsiteURL", objRs(12)
			newElem.setAttribute "Pincode", objRs(13)
			newElem.setAttribute "ExciseControlCode", objRs(14)
			newElem.setAttribute "LocalSTNoandDT", objRs(15)
			newElem.setAttribute "CentralSTNoandDT", objRs(16)
			newElem.setAttribute "IncomeTaxPANNo", objRs(17)
			newElem.setAttribute "PartyGroupCoyType", objRs(18)
			newElem.setAttribute "PartyName", objRs(2)
			newElem.text= objRs(1)
			newElem.setAttribute "Units",sUnit
			newElem.setAttribute "InTrans",sType
			newElem.setAttribute "TINNumber",objRs(19)
			newElem.setAttribute "Useable",objRs(20)
			Root.appendChild newElem
			objRs.MoveNext
		loop
	end if
	Response.ContentType="text/xml"
	Response.Write OutData.xml
	objRs.Close
%>
