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
	'Program Name				:	BankBookDetailsUpdate.asp
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
dim sBankAddress1,sBankAddress2,sCity,sState,sCountry,sPinCode,sPhoneNos,sMobileNos,sFaxNos
dim sEMailId,sWebSiteURL,bPrintCheques,bPrintPayInSlip,sAccountType,sAccountNo,dCreditLimit
dim dOverDraftLimit,dDiscountLimit,dLCLimit,sSwiftCode,iCharges,iDiscount,sBankName

dim sQuery,sorgID,iBookNo,bActionFlag
dim oDOM,Root,Node

set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.async = False
oDOM.load(Request)

set Root = oDOM.documentElement

If Root.hasChildNodes Then
	For Each Node in Root.childNodes
		sorgID = Node.getAttribute("UnitID")
		iBookNo= Node.getAttribute("BookNo")
		bActionFlag= Node.getAttribute("ActionFlag")
		
		sBankName= Node.getAttribute("BankName")
		sBankAddress1= Node.getAttribute("Address1")
		sBankAddress2= Node.getAttribute("Address2")
		sCity= Node.getAttribute("City")
		sState= Node.getAttribute("State")
		sCountry= Node.getAttribute("Country")
		sPinCode= Node.getAttribute("Pincode")
		sPhoneNos= Node.getAttribute("Phone")
		sMobileNos= Node.getAttribute("MobileNo")
		sFaxNos= Node.getAttribute("Fax")
		sEMailId= Node.getAttribute("EMail")
		sWebSiteURL= Node.getAttribute("WebSite")
		
		bPrintCheques= Node.getAttribute("PrintCheque")
		bPrintPayInSlip= Node.getAttribute("PrintPayInSlip")
		sAccountType= Node.getAttribute("AccountType")
		sAccountNo= Node.getAttribute("AccountNo")
		dCreditLimit= Node.getAttribute("CreditLimit")
		dOverDraftLimit= Node.getAttribute("ODLimit")
		dDiscountLimit= Node.getAttribute("DiscountLimit")
		dLCLimit=Node.getAttribute("LCLimit")
		sSwiftCode= Node.getAttribute("SwiftCode")
		iCharges= Node.getAttribute("ChargesHead")
		iDiscount= Node.getAttribute("DiscountHead")
		
	Next
End IF

if bActionFlag="I" then
	sQuery="insert into Acc_M_BankDetails (OUDefinitionID,BookCode,BookNumber,BankName,"&_
		"BankAddress1, BankAddress2,City, State, Country, PinCode, PhoneNos,"&_
		"MobileNos, FaxNos, EMailId, WebSiteURL, PrintCheques, PrintPayInSlip, AccountType,"&_
		"AccountNo, CreditLimit, OverDraftLimit, DiscountingLimit, LCLimit, SwiftCode,BankChargesHead,BillDiscountingHead)"&_

		"Values ('"&sorgID&"','02',"&iBookNo&",'"&sBankName&"','"&sBankAddress1&"','"&sBankAddress2&"','"&sCity&"','"&_
		""&sState&"','"&sCountry&"','"&sPinCode&"','"&sPhoneNos&"','"&sMobileNos&"','"&sFaxNos&"','"&_
		""&sEMailId&"','"&sWebSiteURL&"',"&bPrintCheques&","&bPrintPayInSlip&",'"&sAccountType&"','"&_
		""&sAccountNo&"',"&dCreditLimit&","&dOverDraftLimit&","&dDiscountLimit&","&dLCLimit&",'"&sSwiftCode&"',"
		if 	CInt(iCharges)>0 then
			sQuery=sQuery&iCharges&","
		else
			sQuery=sQuery&"NULL,"
		end if
		if 	CInt(iDiscount)>0 then
			sQuery=sQuery&iDiscount&")"
		else
			sQuery=sQuery&"NULL)"
		end if

else
	sQuery="update Acc_M_BankDetails set BankName='"&sBankName&"',BankAddress1='"&sBankAddress1&"', BankAddress2='"&sBankAddress2&"',"&_
		"City='"&sCity&"',State='"&sState&"',Country='"&sCountry&"', PinCode='"&sPinCode&"',"&_
		"PhoneNos='"&sPhoneNos&"',MobileNos='"&sMobileNos&"',FaxNos='"&sFaxNos&"',EMailId='"&sEMailId&"',"&_
		"WebSiteURL='"&sWebSiteURL&"', PrintCheques="&bPrintCheques&", PrintPayInSlip="&bPrintPayInSlip&","&_
		"AccountType='"&sAccountType&"',AccountNo='"&sAccountNo&"',CreditLimit="&dCreditLimit&","&_
		"OverDraftLimit="&dOverDraftLimit&", DiscountingLimit="&dDiscountLimit&", LCLimit="&dLCLimit&", SwiftCode='"&sSwiftCode&"',"

		if 	CInt(iCharges)>0 then
			sQuery=sQuery&" BankChargesHead="&iCharges&","
		else
			sQuery=sQuery&" BankChargesHead=NULL,"
		end if
		if 	CInt(iDiscount)>0 then
			sQuery=sQuery&" BillDiscountingHead="&iDiscount
		else
			sQuery=sQuery&"BillDiscountingHead=NULL"
		end if
		sQuery=sQuery&" where OUDefinitionID='"&sorgID&"'and BookCode='02'and BookNumber="&iBookNo

end if

con.Execute(sQuery)
%>

