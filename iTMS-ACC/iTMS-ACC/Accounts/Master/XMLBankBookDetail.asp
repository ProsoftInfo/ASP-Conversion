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
	'Program Name				:	BankBookDetailXML
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	January 20, 2003
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
dim dcrs,sSql,OutData,Root,newElem
dim objRs,objRs1,sQuery,sorgID,iBookNo,iChargeHead,iDisHead
	
Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")
	
sorgID = Request("orgID")
iBookNo= Request("BookNo")
sQuery="SELECT isNull(BankAddress1,''), isNull(BankAddress2,''), isNull(City,''), isNull(State,''), isNull(Country,''),isNull(PinCode,''), isNull(PhoneNos,''),"&_
		"isNull(MobileNos,''), isNull(FaxNos,''), isNull(EMailId,''), isNull(WebSiteURL,''),isNull( PrintCheques,''), isNull(PrintPayInSlip,''), isNull(AccountType,''),"&_
		"isNull(AccountNo,''), isNull(CreditLimit,0), isNull(OverDraftLimit,0),isNull( DiscountingLimit,0),isNull(LCLimit,0),isNull( SwiftCode,''),isnull(BankChargesHead,0),"&_
		"isnull(BillDiscountingHead,0),BankName FROM Acc_M_BankDetails where "&_	
		" OUDefinitionID='"&sorgID&"' and BookCode=02 and BookNumber="&iBookNo
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
				
			Set newElem = OutData.createElement("Book")	
			
			newElem.setAttribute "BankName",trim(objRs(22))		
			newElem.setAttribute "BankAddress1", trim(objRs(0))
			newElem.setAttribute "BankAddress2",trim(objRs(1))
			newElem.setAttribute "City", trim(objRs(2))
			newElem.setAttribute "State", trim(objRs(3))
			newElem.setAttribute "Country", trim(objRs(4))
			newElem.setAttribute "PinCode",trim(objRs(5))
			newElem.setAttribute "PhoneNos", trim(objRs(6))
			newElem.setAttribute "MobileNos",trim(objRs(7))
			newElem.setAttribute "FaxNos", trim(objRs(8))
			newElem.setAttribute "EMailId", trim(objRs(9))
			newElem.setAttribute "WebSiteURL",trim(objRs(10))
						
			newElem.setAttribute "PrintCheques", trim(objRs(11))
			newElem.setAttribute "PrintPayInSlip",trim(objRs(12))
			newElem.setAttribute "AccountType", trim(objRs(13))
			newElem.setAttribute "AccountNo", trim(objRs(14))
			newElem.setAttribute "CreditLimit",trim(objRs(15))
			newElem.setAttribute "OverDraftLimit", trim(objRs(16))
			newElem.setAttribute "DiscountingLimit",trim(objRs(17))
			newElem.setAttribute "LCLimit", trim(objRs(18))
			newElem.setAttribute "SwiftCode",trim(objRs(19))
			newElem.setAttribute "ChargeHead",trim(objRs(20))
			newElem.setAttribute "DiscountHead",trim(objRs(21))
			
			if cint(objRs(20)) >0 then
				sQuery="select AccountDescription from Acc_M_GLAccountHead where AccountHead="&objRs(20)
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				newElem.setAttribute "ChargeHeadName",objRs1(0)
				objRs1.Close
			else
				newElem.setAttribute "ChargeHeadName",""			
			end if
			newElem.setAttribute "DiscountHead",trim(objRs(21))
			if cint(objRs(21)) >0 then
				sQuery="select AccountDescription from Acc_M_GLAccountHead where AccountHead="&objRs(21)
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				newElem.setAttribute "DiscountHeadName",objRs1(0)
				objRs1.Close
			else
				newElem.setAttribute "DiscountHeadName",""			
			end if		
			Root.appendChild newElem
	objRs.MoveNext
	loop

	Response.ContentType="text/xml"
	Response.Write OutData.xml
end if
objRs.Close
%>
