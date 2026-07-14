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
	'Program Name				:	XMLGetContactData.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Kalaiselvi R
	'Created On					:	sep 28,2011
	'Modified By				:	
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
<%
dim OutData,Root,newElem

dim objRs,rsTemp,sQuery,iContactNo,iPartyCode,sPartyName

Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set rsTemp = Server.CreateObject("ADODB.Recordset")

iContactNo = Request.QueryString("ContactNo")


sQuery="SELECT ContactName,Designation,ContactPersonFor,isNull(AddressLine1,''), isNull(AddressLine2,''),"&_
	"isNull(City,''), isNull(State,''), isNull(Country,''), isNull(PhoneNos,''), isNull(MobileNos,''), isNull(FaxNos,''), isNull(Email,''), "&_
	"isNull(WebsiteURL,''), isNull(Pincode,''),isNull(Useable,0),isNull(PartyCode,0) FROM APP_M_Contacts where ContactNumber="& iContactNo
'Response.Write "<p>" & sQuery
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
		
			iPartyCode = "0"
			sPartyName = ""
			
			if trim(objRs(15)) <> "0" then
			
				iPartyCode = objRs(15)
				
				sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode =" & iPartyCode & ""
				' Response.Write sQuery
				with rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = con
					.Source = sQuery
					.Open
				end with
															
				if not rsTemp.EOF then
					sPartyName = rsTemp(0)
				end if 
				rsTemp.Close
			end if 'if trim(objRs(15)) <> "0" then
	
	
			Set newElem = OutData.createElement("Contact")			
			newElem.setAttribute "ContactName", objRs(0)
			newElem.setAttribute "Designation", objRs(1)
			newElem.setAttribute "ContactPersonFor", objRs(2)
			newElem.text= objRs(0)
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
			newElem.setAttribute "Useable",objRs(14)
			newElem.setAttribute "PartyCode",iPartyCode
			newElem.setAttribute "PartyName",sPartyName
			
			Root.appendChild newElem
			
			objRs.MoveNext
		loop
	end if
	Response.ContentType="text/xml"
	Response.Write OutData.xml
	objRs.Close
%>
