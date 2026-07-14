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
	'Program Name				:	XMLGetOrgBook.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 21, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	VouCashBookSelection.asp
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
<!--#include virtual="/include/Accpopulate.asp"-->

<%
	dim objRs,objRs1,sQuery,OutData,Root,newElem
	dim sorgID,iBookCode,sClosingCRDR

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")

	sorgID = Request("orgID")
	iBookCode= Request("BkCode")
	sQuery="select BookNumber,Upper(BookName),isnull(BookAccountHead,0),OtherUnitTransaction from "&_
		"vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode="&iBookCode

''blocked by ragav on Jan 05,2012 for To Avoid Without Account Head mapping  Books
'	if iBookCode="01" or iBookCode="02"  then
'		sQuery=sQuery&" and BookAccountHead is not null "
'	end if
''end	

''blocked by ragav on Jan 05,2012 for To Avoid Without Account Head mapping  Books
	if iBookCode="01" or iBookCode="02" or iBookCode="04" or iBookCode="05"  then
		sQuery=sQuery&" and BookAccountHead is not null "
	end if
''end	
	
	sQuery=sQuery&" Order By BookName "

	'Response.Write sQuery
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
			if iBookCode="01" or iBookCode="02"  then
				sQuery="select ClosingCDIndication from Acc_T_GLAccOpeningAmt where "&_
					" AccountHead="& trim(objRs(2)) &" and ClosingMonthYear='"&getToFinYear&"'"

 	'		Response.Write sQuery
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing
				if not objRs1.EOF then
					sClosingCRDR=objRs1(0)
				end if
				objRs1.Close
			end if

			dim bPrintCheques,bPrintPayInSlip,dCreditLimit,dOverDraftLimit,dDiscountingLimit,dLCLimit

			if iBookCode="02"  then

				sQuery=" select PrintCheques,PrintPayInSlip,isnull(CreditLimit,0),isnull(OverDraftLimit,0),isnull(DiscountingLimit,0),isnull(LCLimit,0) "&_
				" from Acc_M_BankDetails where OUDefinitionID = '" & sorgID & "' and BookCode='02' and BookNumber="&trim(objRs(0))
				
     '           Response.Write sQuery

				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing
				if not objRs1.EOF then
					bPrintCheques=objRs1(0)
					bPrintPayInSlip=objRs1(0)
					dCreditLimit=objRs1(0)
					dOverDraftLimit=objRs1(0)
					dDiscountingLimit=objRs1(0)
					dLCLimit=objRs1(0)
				end if
				objRs1.Close
			end if
				Set newElem = OutData.createElement("Book")
				newElem.setAttribute "BookNumber", trim(objRs(0))
				newElem.setAttribute "BookName", trim(objRs(1))
				newElem.setAttribute "AccHead", trim(objRs(2))
				newElem.setAttribute "OtherUnit", trim(objRs(3))
				newElem.setAttribute "ClosingCRDR", sClosingCRDR
				newElem.setAttribute "PrintCheques", bPrintCheques
				newElem.setAttribute "PrintPayInSlip", bPrintPayInSlip
				newElem.setAttribute "CreditLimit", dCreditLimit
				newElem.setAttribute "OverDraftLimit", dOverDraftLimit
				newElem.setAttribute "DiscountingLimit", dDiscountingLimit
				newElem.setAttribute "LCLimit", dLCLimit

				Root.appendChild newElem
		objRs.MoveNext
		loop
		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close
%>

