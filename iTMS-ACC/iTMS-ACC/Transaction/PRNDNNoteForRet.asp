<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PRNDNNoteForRet.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	29 March 2006
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->

<%
'------------------------Declaration Constants -----------------------------
	dim aiHeadColWidth(5,8)

	'WIDTH SPECIFICATION FOR PAGE TITLE 1
	aiHeadColWidth(0,0)=80

	'WIDTH SPECIFICATION FOR PAGE TITLE 2
	aiHeadColWidth(1,0)=0 '3
	aiHeadColWidth(1,1)=4
	aiHeadColWidth(1,2)=45
	aiHeadColWidth(1,3)=13
	aiHeadColWidth(1,4)=15

	'WIDTH SPECIFICATION FOR OPENING/CLOSING  LINE 
	aiHeadColWidth(2,0)=5
	aiHeadColWidth(2,1)=80
	aiHeadColWidth(2,2)=5

	'WIDTH SPECIFICATION FOR HEADING DETAIL LINE 
	aiHeadColWidth(3,0)=0 '3
	aiHeadColWidth(3,1)=15
	aiHeadColWidth(3,2)=20
	aiHeadColWidth(3,3)=20
	aiHeadColWidth(3,4)=15
	aiHeadColWidth(3,5)=10
	aiHeadColWidth(3,6)=3

	'WIDTH SPECIFICATION FOR HEADING DETAIL LINE 
	aiHeadColWidth(4,0)=0 '3
	aiHeadColWidth(4,1)=4
	aiHeadColWidth(4,2)=10
	aiHeadColWidth(4,3)=22
	aiHeadColWidth(4,4)=22
	aiHeadColWidth(4,5)=7
	aiHeadColWidth(4,6)=15
	aiHeadColWidth(4,7)=3	

	dim objFSO,objTxt
	dim iPgNo,iLineNo,sTStr,i,sFinalText,sUnit,sUnitName,sVouDate,sNewNarr
	Dim sSalInfo,AccVoucherNo,iEnNo

	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"&Session.SessionID&"_CNNote_View.txt"))

	dim iTransNo,strText,XmlData,Root,Node,subNode
	dim rs,Au,Acccode,AcName,Drcr,Amount,TotAmt,Narration,sQuery,sUnitAdd,sUnitCity,sUnitPin
	Dim iPartyCode,sTemp,sExp,TempNode,sPurTyFor,sPurInfo,sRetVal
	

	iTransNo=Request("iTransNo")
	'Response.Write iTransNo

	set XmlData=server.CreateObject("Microsoft.XMLDOM")
	set rs=server.CreateObject("ADODB.Recordset")

	'XmlData.load server.MapPath("../XmlData/Voucher/" & iTransNo & ".xml")
	'XmlData.load server.MapPath("../XmlData/Voucher/84.xml")
	sRetVal = GetVouchXML(iTransNo)
	XmlData.Load server.MapPath(sRetVal)

	set Root=XmlData.documentElement
	
	sExp = "//Party"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		iPartyCode = TempNode.Item(0).attributes.getNamedItem("ParCode").value
	End IF
	
	
	with rs
		.CursorLocation=3
		.CursorType=3
		.Source ="SELECT PARTYNAME, ADDRESSLINE1, ADDRESSLINE2, CITY, STATE, COUNTRY, PINCODE FROM APP_M_PARTYMASTER WHERE PARTYCODE = "&iPartyCode&" "
		.ActiveConnection=con
		.Open 
	end with
	set rs.ActiveConnection=nothing

	dim Name,Address1,Address2,City,State,Country,Pincode
	if not rs.EOF then
		Name=rs(0)
		Address1=rs(1)
		Address2=rs(2)
		city=rs(3)
		State=rs(4)
		Country=rs(5)
		Pincode=rs(6)
	end if
	rs.Close 

	sExp = "//Organization"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sUnit = TempNode.Item(0).attributes.getNamedItem("OrgId").value
	End IF
	
	sExp = "//Details"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sVouDate = TempNode.Item(0).attributes.getNamedItem("VouDate").value
	End IF
	
	'sUnit = Root.Attributes.Item(0).nodeValue
	
	sQuery = "Select OrgUnitDescription,Address1,isNull(Address2,0),City,PostCode From  "&_
			 "DCS_OrganizationUnitDefinitions Where OUDefinitionID = '"&sUnit&"' "
	
	rs.Open sQuery,con
	IF Not rs.EOF Then
		sUnitName = rs(0)
		sUnitAdd = rs(1)&" " & rs(2)
		sUnitCity = rs(3)
		sUnitPin = rs(4)
	End IF
	rs.Close		 
	
	ConsPgHeader()
	totamt=0
	
	sExp = "//PurInvoice"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		'sPurTyFor = TempNode.Item(0).Attributes.getNamedItem("CRNoteType").Value
		sPurInfo = TempNode.Item(0).Attributes.getNamedItem("PurInvNo").Value
		sPurInfo = sPurInfo&" DTD "
		sPurInfo = sPurInfo&TempNode.Item(0).Attributes.getNamedItem("PurInvDate").Value
	End IF
	
	IF CStr(sPurInfo) = "" Then
		sExp = "//SaleInvoice"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			'sPurTyFor = TempNode.Item(0).Attributes.getNamedItem("CRNoteType").Value
			sSalInfo = TempNode.Item(0).Attributes.getNamedItem("InvNo").Value
			sSalInfo = sSalInfo&" DTD "
			sSalInfo = sSalInfo&TempNode.Item(0).Attributes.getNamedItem("InvDate").Value
		End IF
	End IF
	
	IF CStr(sPurTyFor) = "Q" Then
		sPurTyFor = "Quality "
	Elseif CStr(sPurTyFor) = "R" Then
		sPurTyFor = "Rate "
	Elseif CStr(sPurTyFor) = "D" Then	
		sPurTyFor = "Discount "
	Else
		sPurTyFor = "Quantity "
	End IF
	
	sExp = "//TaxDetails"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		Amount = TempNode.Item(0).attributes.Item(0).Nodevalue
	End IF
	
	sExp = "//Narration"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sNewNarr = TempNode.Item(0).Text
	End IF
	
	IF CStr(sPurInfo) <> "" Then
		Narration = "Being For the Purchase Inv No:"&sPurInfo&" Returned For Quality "&sNewNarr
	Else
		Narration = "Being For the Sales Inv No:"&sSalInfo&" "&sNewNarr
	End IF
	
	strText=strText & myalign(Narration,aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber( Amount),aiHeadColWidth(4,6),"R") & vbcrlf
	strText=strText & myalign("" & Trim(Mid(Narration,60)),aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & vbcrlf
	
	iLineNo=iLineNo+1
	iEnNo=iEnNo+1

	'for each node in root.childNodes
	'	au=node.attributes.getNamedItem("AccUnit").value
	'	drcr=node.attributes.getNamedItem("CRDR").value
	'	Amount=node.attributes.getNamedItem("Amount").value
	'	TotAmt =TotAmt+amount
	'	for each SubNode in node.ChildNodes
	'		if subnode.nodeName="AccHead" then
	'		'acccode=subNode.attributes.getNamedItem("No").value
	'		'with rs
	'		'	.CursorLocation=3
	'		'	.CursorType=3
	'		'	.Source="Select AccountDescription from Acc_M_GLAccountHead where AccountHead=" & Acccode 
	'		'	.ActiveConnection=con
	'		'	.Open 
	'		'end with
	'		'set rs.ActiveConnection =nothing
	'		'if not rs.EOF then Acccode=rs(0)
	'		'rs.Close 
'
'				'strText=strText & string(aiHeadColWidth(4,0)," ") & myalign(right(au,2),aiHeadColWidth(4,1),"L") & myalign(left(acccode,9) ,aiHeadColWidth(4,2),"L") & myalign(subNode.attributes.getNamedItem("Name").value,aiHeadColWidth(4,3),"L") & myalign("",aiHeadColWidth(4,4),"L") & centerAlign( drcr,aiHeadColWidth(4,5)) & myalign( FormatNumber( Amount,2,0,0,0 ),aiHeadColWidth(4,6),"R") & vbcrlf
'			'	strText=strText & myalign("Being " & subNode.attributes.getNamedItem("Name").value,aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber( Amount,2,0,0,0 ),aiHeadColWidth(4,6),"R") & vbcrlf
'			'	iLineNo=iLineNo+1
'			'	CheckNew
'			elseif subNode.nodeName="Narration" then
'				if isnull(subNode.Text) then
'					Narration=""
'				else
'					Narration=subNode.Text
'				end if
'				'strText=strText & string(aiHeadColWidth(4,0)," ") & myalign("",aiHeadColWidth(4,1),"L") & myalign("",aiHeadColWidth(4,2),"R") & myalign(Narration,aiHeadColWidth(4,3),"L") & vbcrlf
'				strText=strText & myalign("Being " & Narration,aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber( Amount,2,0,0,0 ),aiHeadColWidth(4,6),"R") & vbcrlf
'				strText=strText & myalign("" & Trim(Mid(Narration,60)),aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & vbcrlf
'				
'				iLineNo=iLineNo+1
'				CheckNew
'			end if
'		next
'	next

	for i=1 to 6-CInt(iEnNo)
		strText =strText & string(80," ") & vbcrlf
	next

	strText = strtext & string(80,"-") & vbcrlf
	'strText =strText & string(50," ") & "Total   " & myalign(FormatNumber(TotAmt,2,0,0,0),15,"R") & vbcrlf
	strText = strText & myalign("Total ",aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber(Amount),aiHeadColWidth(4,6),"R") & vbcrlf
	strText = strtext & string(80,"-") & vbcrlf
	strText = strText & "(" & amountwords(Amount) & ")" & vbcrlf			
	strText =strText & " "& vbcrlf
	strText =strText & " "& vbcrlf
	strText =strText & " "& vbcrlf

	strText = strText &myalign("PreparedBy ",15,"L")
	strText = strText &myalign(" ",15,"L")
	strText = strText &myalign("AO / FM",15,"L")
	strText = strText &myalign(" ",15,"L")
	strText = strText &myalign("M.D/Director ",30,"L") & vbcrlf	
	sFinalText = sFinalText & strtext
	objTxt.write sFinalText
	if strText<>"" then
		Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Reports/"&Session.SessionID&"_CNNote_View.txt&exitpath=/accounts/reports/CreditNoteSelection.asp&frame=_parent")
	else
		Response.Clear
	end if
%>
                          
<%
	function ConsPgHeader()

		sQuery="Select CreatedVouchStatus,VoucherNumber from Acc_T_CreatedVoucherHeader H , Acc_T_VoucherHeader v where H.CreatedTransNo=v.CreatedTransNo and " _
		& "right(H.CreatedVouchStatus,2)=04  and H.CreatedTransNo="&iTransNo
		With Rs
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		End With
		Set Rs.ActiveConnection = nothing
		if not rs.EOF then AccVoucherNo=Rs(1)
		Rs.Close
		
		strText=string(80," ") & vbcrlf
		strText=strText & centerAlign(sUnitName,aiHeadColWidth(0,0)) & vbcrlf
		strText=strText & centerAlign(sUnitAdd,aiHeadColWidth(0,0)) & vbcrlf
		strText=strText & centerAlign(sUnitCity &" - " & sUnitPin,aiHeadColWidth(0,0)) & vbcrlf
		strText=strtext & string(80," ") & vbcrlf
		strText=strText & centerAlign("Debit Note",aiHeadColWidth(0,0)) & vbcrlf

		strText=strText & string(aiHeadColWidth(1,0)," ") & myalign("To",aiHeadColWidth(1,1),"L") & myalign(name,aiHeadColWidth(1,2),"L") &  myalign("Date ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sVouDate,aiHeadColWidth(1,4),"L") & vbcrlf
		strText=strText & string(aiHeadColWidth(1,0)," ") & string(aiHeadColWidth(1,1)," ") & myalign(Address1,aiHeadColWidth(1,2),"L") &  myalign("Ref No ",aiHeadColWidth(1,3),"L") & ": " &  myalign(root.attributes.getNamedItem("CreatedVouNo").value,aiHeadColWidth(1,4),"L") & vbcrlf
		strText=strText & string(aiHeadColWidth(1,0)," ") & string(aiHeadColWidth(1,1)," ") & myalign(Address2,aiHeadColWidth(1,2),"L") &  myalign("Vou No ",aiHeadColWidth(1,3),"L") & ": " &  myalign(AccVoucherNo ,aiHeadColWidth(1,4),"L") & vbcrlf
		strText=strText & string(aiHeadColWidth(1,0)," ") & string(aiHeadColWidth(1,1)," ") & city & vbcrlf
		strText=strText & string(aiHeadColWidth(1,0)," ") & string(aiHeadColWidth(1,1)," ")& State & vbcrlf
		strText=strText & string(aiHeadColWidth(1,0)," ") & string(aiHeadColWidth(1,1)," ")& Country & " - " & Pincode & vbcrlf
		
		strText=strtext & string(80," ") & vbcrlf
		strText=strtext & string(aiHeadColWidth(1,0)," ") & "Sir / Sirs We have today debited your account with us as detailed below : " & vbcrlf
		strText=strtext & string(80,"-") & vbcrlf
		strText=strText & string(aiHeadColWidth(4,0)," ") & myalign("Particulars",aiHeadColWidth(4,1) + aiHeadColWidth(4,2),"L") & myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",aiHeadColWidth(4,4),"L") & myalign(" ",aiHeadColWidth(4,5),"L") & myalign("Amount (Rs.Ps)",aiHeadColWidth(4,6),"R") & vbcrlf
		strText=strtext & string(80,"-") & vbcrlf
		iLineNo=18
	end Function

	Function CheckNew
		if iLineNo>=22 then
			strText =strtext & string(80,"-") & vbcrlf
			strText =strText & string(50," ") & "NET AMOUNT :   " & myalign(FormatNumber(TotAmt,2,0,0,0),15,"R") & vbcrlf
			strText =strtext & string(80,"-") & vbcrlf
			strText =strText & "AMOUNT : (" & amountwords(totamt) & ")" & vbcrlf
			strText =strText & myalign("Kindly acknowledge the receipt and arrange to pass necessary ",80,"L") & vbcrlf
			strText =strText & myalign("entries in your books.  ",80,"L") & vbcrlf
			strText =strText & myalign("Thanking You ",40,"L") 
			strText =strText & myalign("Yours faithfully ",40,"R") & vbcrlf
			
			strText =strText & myalign(sUnitName,80,"R") & vbcrlf
			strText =strText & string(80," ") & vbcrlf
			strText =strText & string(80," ") & vbcrlf
			strText =strText & myalign("Authorised Signatory",80,"R") & vbcrlf
			strText =strtext & string(80,"-") & vbcrlf & chr(12)

			sFinalText = sFinalText & sTstr & chr(12)
			sTstr = ConsPgHeader()
		end if	
	end Function

	function centerAlign(str1,width)
		dim diff,strlen,val, i, str, newstr, blank
			str = str1
			strlen = len(str)
			diff = width - strlen
			for i=0 to (diff-1)/2
				blank = blank & " "
			next
			newstr = blank & str & blank
		centerAlign = newstr
	end function
%>
