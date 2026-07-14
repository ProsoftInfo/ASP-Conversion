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
	'Program Name				:	XMLGetBookSalPurType.asp
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
	dim objRs,sQuery,OutData,Root,newElem
	dim sorgID,iBookCode,sClosingCRDR,iBookNo,iAccHead

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")


	'sorgID = Request("orgID")
	sorgID = Session("organizationcode")
	iBookCode= Request("BkCode")
	iBookNo = Request("BookNo")


	'Response.Write sorgID &" " & iBookCode &" " & iBookNo
	sQuery = "Select isNUll(BookAccountHead,0) From Acc_R_ApplicableAccountHeads Where  "&_
			 "BookCode = '"&iBookCode&"' and OUDefinitionID = '"&sorgID&"' and BookNumber = "&iBookNo
'Response.Write sQuery
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iAccHead = objRs(0)
	Else
		iAccHead = 0
	End if
	objRs.Close

	IF CStr(iBookCode) = "05" Then
		sQuery = "Select InvoiceType,InvTypeShortName,InvoiceTypeName From VwSalInvTypes "&_
				 "Where OUDefinitionID = '"&sorgID&"' and AccountHead = "&iAccHead&"  "&_
				 "Order By InvoiceTypeName "


	End IF
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
	Do While Not objRs.EOF

		Set newElem = OutData.createElement("Type")
		newElem.setAttribute "TypeNumber", trim(objRs(0))
		newElem.setAttribute "TypeShortName", trim(objRs(1))
		newElem.setAttribute "TypeName", trim(objRs(2))
		Root.appendChild newElem
		objRs.MoveNext
	loop
	objRs.Close

	Response.ContentType="text/xml"
	Response.Write OutData.xml

%>

