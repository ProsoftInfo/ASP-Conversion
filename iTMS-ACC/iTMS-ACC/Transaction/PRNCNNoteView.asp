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
	'Program Name				:	PRNCNNoteView.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	Tajudeen.S
	'Created On					:	29 March 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:2
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


	'------------------------End of Declaration Constants ----------------------

	'-----------------------------Declaration for Printing--------------------
	dim objFSO,objTxt
	dim iPgNo,iLineNo,sTStr,i,sFinalText,sUnit,sUnitName
	Dim sExp,TempNode,iEntryNo,iCtr


	'----------------------Intialization of file Object-------------------------

	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"&Session.SessionID&"_CRNote_View.txt"))

	'----------------------End of file Object Intialization --------------------

	'----------- To Get The Values From the Selection Page ----------------

	dim iTransNo,strText,XmlData,Root,Node,subNode,strTextC,sAccName
	dim rs,Au,Acccode,AcName,Drcr,Amount,TotAmt,Narration,sQuery,sUnitAdd,sUnitCity,sUnitPin
	Dim iPartyCode,sTemp,AccVoucherNo,iEnNo,sCrVouNo,sAccVouNo,iCrVouNo,iVouAmt,sNarrText
	Dim iVouStatus,iCreatedBy,sCreatedOn,dVouDate,sItemDesc,iInvQty,sInvUOM,iInvRate
	Dim sEmpName,sPreparedBy
	iTransNo=Request.QueryString("iTransNo")
	'Response.Write iTransNo

	set XmlData=server.CreateObject("Microsoft.XMLDOM")
	set rs=server.CreateObject("ADODB.Recordset")

	XmlData.load server.MapPath("../XmlData/Voucher/" & iTransNo & ".xml")
	'XmlData.load server.MapPath("../XmlData/Voucher/84.xml")
	set Root=XmlData.documentElement
	
	'iPartyCode = Root.attributes.getNamedItem("PartyCode").value
	'sTemp = Split(iPartyCode,"?")
	'iPartyCode = sTemp(3)
	sQuery = "Select CreatedVoucherNo,PartyCode,OUDefinitionID,CreatedVouchStatus,CreatedBy,convert(Varchar,CreatedOn,103) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
	rs.open sQuery,Con
	IF Not rs.EOF Then
		sCrVouNo = rs(0)	
		iPartyCode = rs(1)	
		sUnit = rs(2)	
		iVouStatus = rs(3)
		iCreatedBy = rs(4)
		sCreatedOn = rs(5)
	End IF
	rs.Close  

	with rs
		.CursorLocation=3
		.CursorType=3
		.Source ="SELECT PARTYNAME,IsNull(ADDRESSLINE1,''), IsNull(ADDRESSLINE2,''), CITY, STATE, COUNTRY, PINCODE FROM APP_M_PARTYMASTER WHERE PARTYCODE = "&iPartyCode&" "
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
	sQuery = "SELECT LoginId FROM DCS_User WHERE EmployeeNumber ="&iCreatedBy
	rs.open sQuery,Con	
	IF not rs.EOF  Then
		sEmpName = Trim(rs(0))
	End IF
	rs.Close
	IF Cstr(iVouStatus) <> "010104" Then
		sPreparedBy	= sEmpName &"/" & sAccVouNo &" - " & sCreatedOn

	Else
		sPreparedBy	= sEmpName & "-" & sCreatedOn
	End IF
	sQuery = "Select CreatedVoucherNo From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo
	rs.open sQuery,Con
	IF Not rs.EOF Then
		sCrVouNo = rs(0)		
	End IF
	rs.Close
	
	sQuery = "Select VoucherNumber From Acc_T_VoucherHeader Where CreatedTransNo = "&iTransNo
	rs.open sQuery,Con
	IF Not rs.EOF Then
		sAccVouNo = rs(0)		
	End IF
	rs.Close
	


	sQuery = "Select M.AccountDescription From Acc_T_CreatedVoucherDetails T,Acc_M_GLAccountHead M Where "&_
			 "M.AccountHead = T.AccUnitAccountHead and T.CreatedTransNo = "&iTransNo
			 
	rs.Open sQuery,con
	IF Not rs.EOF Then
		sAccName = rs(0)		
	End IF
	rs.Close


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
	iLineNo=17
	'for each node in root.childNodes
	sNarrText = "Being the amount credited to your Account towards the"&Narration&" Invoice No:"&iCrVouNo&"/"&iVouAmt&" Dt."&dVouDate  	
	sNarrText = BreakString(sNarrText,30)

	sQuery = "Select VoucherEntryNumber,AccountingUnit,TransCrDrIndication,ItemDescription,InvoicedQuantity,InvoicedUOM,InvoicedRate,Amount,Isnull(VoucherNarration,'')  From Acc_T_CreatedVoucherdetails Where CreatedTransNo = "&iTransNo
	rs.open sQuery,Con	
	'Response.Write iCrVouNo
	'Response.End 
	do while not rs.EOF 
		iEnNo = iEnNo + 1
		iEntryNo = rs(0)
		au = rs(1)
		drcr = rs(2)
		sItemDesc = rs(3)
		iInvQty = rs(4)
		sInvUOM = rs(5)
		iInvRate = rs(6)
		Amount = rs(7)
		Narration = rs(8)	 
		TotAmt = CDbl(TotAmt) + CDbl(Amount) 
		
		'strText=strText & myalign(sNarrText & Narration,aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber( Amount),aiHeadColWidth(4,6),"R") & vbcrlf
		strText=strText & myalign(sNarrText(0)&Narration&" Invoice No:"&iCrVouNo&"/"&iVouAmt&" Dt."&dVouDate  ,58,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber( Amount),aiHeadColWidth(4,6),"R") & vbcrlf
		
		rs.MoveNext 
	loop
	rs.Close			
	
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(FormatNumber(TotAmt),aiHeadColWidth(4,6),"R") & vbcrlf						
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf												
	
	for i=1 to 6-CInt(iEnNo)
		strText =strText & string(80," ") & vbcrlf
	next

	strText =strtext & string(80,"-") & vbcrlf
	'strText =strText & string(50," ") & "Total   " & myalign(FormatNumber(TotAmt,2,0,0,0),15,"R") & vbcrlf
	strText=strText & myalign("Total ",aiHeadColWidth(4,3) + aiHeadColWidth(4,0) + aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber(TotAmt),aiHeadColWidth(4,6),"R") & vbcrlf
	strText =strtext & string(80,"-") & vbcrlf
	strText =strText & "(" & amountwords(totamt) & ")" & vbcrlf

	strText =strText & " "& vbcrlf
	strText =strText & " "& vbcrlf
	strText =strText & " "& vbcrlf
	strText =strText & " "& vbcrlf

	strText = strText &myalign("PreparedBy ",15,"L")
	strText = strText &myalign(" ",15,"L")
	'strText = strText &myalign("AO / FM",15,"L")
	strText = strText &myalign(" ",30,"L")
	strText = strText &myalign("M.D/Director ",30,"L") & vbcrlf
	
	sFinalText = sFinalText & strtext & " " & Vbcrlf &" " & vbcrlf 
	
	
	objTxt.write sFinalText
	if strText<>"" then
		'Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Reports/"&Session.SessionID&"_DBNote_View.txt&exitpath=/accounts/reports/CreditNoteSelection.asp&frame=_parent")
		Response.Redirect("../../Components/FormattPrintNew.asp?server=server&filepath=/accounts/temp/Reports/"&Session.SessionID&"_CRNote_View.txt&exitpath=/accounts/reports/CreditNoteSelection.asp&frame=_parent")
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
		strText=strText & centerAlign("Credit Note",aiHeadColWidth(0,0)) & vbcrlf
		'strText=strText & myalign("To",aiHeadColWidth(1,1),"L") & myalign(name,aiHeadColWidth(1,2),"L") &  myalign("Date   ",aiHeadColWidth(1,3),"L") & ": " &  myalign(dVouDate,aiHeadColWidth(1,4),"L") & vbcrlf
		strText=strText & myalign("To",aiHeadColWidth(1,1),"L") & vbCrLf 
		strText=strText & myalign(" ",4,"L") & myalign(name,aiHeadColWidth(1,2),"L") &  myalign("Date   ",aiHeadColWidth(1,3),"L") & ": " &  myalign(dVouDate,aiHeadColWidth(1,4),"L") & vbcrlf
		'strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address1,aiHeadColWidth(1,2),"L") &  myalign("Ref No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sCrVouNo,aiHeadColWidth(1,4),"L") & vbcrlf		
		strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address2,aiHeadColWidth(1,2),"L") &  myalign("Vou No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sAccVouNo ,aiHeadColWidth(1,4),"L") & vbcrlf		
		strText=strText & string(aiHeadColWidth(1,1)," ")& city & " - " & Pincode & vbcrlf
		strText=strText & string(aiHeadColWidth(1,1)," ")& State & " " & Country & vbcrlf
		'strText=strText & string(aiHeadColWidth(1,1)," ")& Country & " - " & Pincode & vbcrlf		
		strText=strtext & string(80," ") & vbcrlf
		'strText=strtext & string(aiHeadColWidth(1,0)," ") & "Sir / Sirs We have today credited your account with us as detailed below : " & vbcrlf
		strText=strtext & string(80,"-") & vbcrlf
		strText=strText & string(aiHeadColWidth(4,0)," ") & myalign("Particulars",aiHeadColWidth(4,1) + aiHeadColWidth(4,2),"L") & myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",aiHeadColWidth(4,4),"L") & myalign(" ",aiHeadColWidth(4,5),"L") & myalign("   Amount",aiHeadColWidth(4,6),"R") & vbcrlf
		strText=strtext & string(80,"-") & vbcrlf
		
		strTextC=string(80," ") & vbcrlf
		strTextC=strTextC & centerAlign(sUnitName,aiHeadColWidth(0,0)) & vbcrlf
		strTextC=strTextC & centerAlign(sUnitAdd,aiHeadColWidth(0,0)) & vbcrlf
		strTextC=strTextC & centerAlign(sUnitCity &" - " & sUnitPin,aiHeadColWidth(0,0)) & vbcrlf
		strTextC=strtextC & string(80," ") & vbcrlf
		strTextC=strTextC & centerAlign("Credit Note",aiHeadColWidth(0,0)) & vbcrlf
		strTextC=strTextC & myalign("To",aiHeadColWidth(1,1),"L") & myalign(name,aiHeadColWidth(1,2),"L") &  myalign("Date   ",aiHeadColWidth(1,3),"L") & ": " &  myalign(dVouDate,aiHeadColWidth(1,4),"L") & vbcrlf
		'strTextC=strTextC & string(aiHeadColWidth(1,1)," ") & myalign(Address1,aiHeadColWidth(1,2),"L") &  myalign("Ref No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sCrVouNo,aiHeadColWidth(1,4),"L") & vbcrlf		
		strTextC=strTextC & string(aiHeadColWidth(1,1)," ") & myalign(Address2,aiHeadColWidth(1,2),"L") &  myalign("Vou No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sAccVouNo ,aiHeadColWidth(1,4),"L") & vbcrlf		
		strTextC=strTextC & string(aiHeadColWidth(1,1)," ")& city & " - " & Pincode & vbcrlf
		strTextC=strTextC & string(aiHeadColWidth(1,1)," ")& State & " "  & Country & vbcrlf
		'strTextC=strTextC & string(aiHeadColWidth(1,1)," ")& Country & " - " & Pincode & vbcrlf		
		
		strTextC=strtextC & string(aiHeadColWidth(1,0)," ") & "A/C Head: " & sAccName & vbcrlf
		strTextC=strtextC & string(80," ") & vbcrlf
		'strTextC=strtextC & string(aiHeadColWidth(1,0)," ") & "Sir / Sirs We have today credited your account with us as detailed below : " & vbcrlf
		strTextC=strtextC & string(80,"-") & vbcrlf
		strTextC=strTextC & string(aiHeadColWidth(4,0)," ") & myalign("Particulars",aiHeadColWidth(4,1) + aiHeadColWidth(4,2),"L") & myalign(" ",aiHeadColWidth(4,3),"L") & myalign(" ",aiHeadColWidth(4,4),"L") & myalign(" ",aiHeadColWidth(4,5),"L") & myalign("   Amount",aiHeadColWidth(4,6),"R") & vbcrlf
		strTextC=strtextC & string(80,"-") & vbcrlf
		
		
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
