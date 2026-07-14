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
	'Program Name				:	PRNCNoteNew.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	FEB 22,2010
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
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/PrintFunctions.asp"-->

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
	Dim sTempArr,sTempVal


	'----------------------Intialization of file Object-------------------------

	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"&Session.SessionID&"_CRNote_View.txt"))

	'----------------------End of file Object Intialization --------------------

	'----------- To Get The Values From the Selection Page ----------------

	dim iTransNo,strText,XmlData,Root,Node,subNode,strTextC,sAccName
	dim rs,rs1,Au,Acccode,AcName,Drcr,Amount,TotAmt,Narration,sQuery,sUnitAdd,sUnitCity,sUnitPin
	Dim iPartyCode,sTemp,AccVoucherNo,iEnNo,sCrVouNo,sAccVouNo,iBankInsNo
	Dim iCrVouNo,dVouDate,iVouAmt,sItemDesc,iInvQty,sInvUOM,iInvRate,iVouStatus
	Dim iTaxCode,iCatCode,iTaxAccHead,iTaxEntNo,iTaxAmt,sTaxName,sNarrFlag,sNarrText
	dim Name,Address1,Address2,City,State,Country,Pincode,sAmtWords
	dim TempTaxName,sTempNarr,TempTaxAmt,sTaxFlag,RebFlag,iRebTaxAmt,iTotTaxAmt
	dim sPreparedBy,sEmpName,iCreatedBy,sCreatedOn,j,sTempItemDesc
	
	
	iTransNo=Request.QueryString("iTransNo")
	iBankInsNo = Request.QueryString("BankInsNo")
	 'Response.Write "iTransNo="& iTransNo
	 'Response.Write "iBankInsNo="&iBankInsNo
	 'Response.End
	set XmlData=server.CreateObject("Microsoft.XMLDOM")
	set rs=server.CreateObject("ADODB.Recordset")
	set rs1=server.CreateObject("ADODB.Recordset")
	XmlData.load server.MapPath("../XmlData/Voucher/" & iTransNo & ".xml")
	''XmlData.load server.MapPath("../XmlData/Voucher/84.xml")
	set Root=XmlData.documentElement

	'iPartyCode = Root.attributes.getNamedItem("PartyCode").value
	
	sQuery = "Select H.CreatedVoucherNo,H.PartyCode,H.OUDefinitionID,H.CreatedVouchStatus,H.CreatedBy,"&_
			"convert(Varchar,H.CreatedOn,103),H.CreatedTransNo,D.AccUnitPartyCode From Acc_T_CreatedVoucherHeader H,Acc_T_CreatedVoucherDetails D"&_
			" Where H.CreatedTransNo = D.CreatedTransNo and H.CreatedTransNo = "&iTransNo &" and D.AccUnitPartyCode is not null"

	Response.Write sQuery
		
	rs.open sQuery,Con
	IF Not rs.EOF Then
		sCrVouNo = rs(0)	
		iPartyCode = rs(7)	
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
	
	sQuery = "Select CreatedVoucherNo,convert(Varchar,VoucherDate,103),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iTransNo &" " 
	Response.Write sQuery 
	rs.Open sQuery,con
	If not rs.EOF then	
		iCrVouNo = rs(0)
		dVouDate = rs(1)
		iVouAmt  = rs(2)
	End If
	rs.Close
	
	sQuery = "Select VoucherNumber From Acc_T_VoucherHeader Where CreatedTransNo = "&iTransNo
	rs.open sQuery,Con
	IF Not rs.EOF Then
		sAccVouNo = rs(0)		
	End IF
	rs.Close
	
	sQuery = "SELECT LoginId FROM DCS_User WHERE EmployeeNumber ="&iCreatedBy
	rs.open sQuery,Con	
	IF not rs.EOF  Then
		sEmpName = Trim(rs(0))
	End IF
	rs.Close
	
	sPreparedBy	= sEmpName &"/"& iCrVouNo & "-" & sCreatedOn
	
	sQuery = "Select M.AccountDescription From Acc_T_CreatedVoucherDetails T,Acc_M_GLAccountHead M Where "&_
			 "M.AccountHead = T.AccUnitAccountHead and T.CreatedTransNo = "&iTransNo
			 
	rs.Open sQuery,con
	IF Not rs.EOF Then
		sAccName = rs(0)		
	End IF
	rs.Close

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
	Response.Write "sCR="& sCrVouNo
	
	iEnNo = 0	
	sQuery = "Select VoucherEntryNumber,AccountingUnit,TransCrDrIndication,ItemDescription,InvoicedQuantity,InvoicedUOM,InvoicedRate,Amount,Isnull(VoucherNarration,'')  From Acc_T_CreatedVoucherdetails Where CreatedTransNo = "&iTransNo &" and AccUnitPartyCode is null"
	'Response.Write sQuery
	'Response.End 
	rs.open sQuery,Con	
	do while not rs.EOF 
	'''''''''''''''''''''''''''''''''''''''''
		iEnNo = iEnNo + 1
		iEntryNo = rs(0)
		au = rs(1)
		drcr = rs(2)
		iInvRate = rs(6)
		Amount = rs(7)
		Narration = rs(8)	 
		sItemDesc = rs(8)
		
		sTempArr =  split(sItemDesc,"&")
		if iEnNo = 1 then
		TotAmt = CDbl(TotAmt) + CDbl(Amount) 
			strText=strText & myalign("",1,"L") & myalign(iEnNo,5,"L") &  myalign(" ",2,"L")
			if UBound(sTempArr) > 1 then
				 strText = strText & myalign(sTempArr(0),aiHeadColWidth(4,3)+15,"L")& myalign("",2,"L") &myalign(sTempArr(1),5,"L")& myalign(sTempArr(2),5,"R") & myalign("",5,"R")& myalign(FormatNumber(sTempArr(3),2),8,"R") & myalign("",3,"L") & myalign( FormatNumber(Amount),aiHeadColWidth(4,6),"R") & vbcrlf						
			else
				strText=strText & myalign(sItemDesc,aiHeadColWidth(4,3)+10,"L") &  myalign("",33,"L") & myalign( FormatNumber(Amount),aiHeadColWidth(4,6),"R") & vbcrlf						
			end if 'if UBound(sTempArr) > 1 then
		end if
					
		rs.MoveNext 
	loop
	rs.Close
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(FormatNumber(TotAmt),aiHeadColWidth(4,6),"R") & vbcrlf						
	strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf												
	
	sNarrText = "Being the amount credited to your Account towards the yarn return vide our Invoice No:"&iCrVouNo&"/"&iVouAmt&" Dt."&dVouDate  	
	sNarrText = BreakString(sNarrText,30)
	
	'Check----Have to display details from taxdet table if tax amount <> 0
	sQuery = "Select VoucherEntryNumber,AccountingUnit,TransCrDrIndication,ItemDescription,InvoicedQuantity,InvoicedUOM,InvoicedRate,Amount,Isnull(VoucherNarration,'')  From Acc_T_CreatedVoucherdetails Where CreatedTransNo = "&iTransNo &" and AccUnitPartyCode is null"
	rs.Open sQuery,con
	iCtr = 0 
	iTotTaxAmt = 0
	iEnNo = 0
	Amount = 0
	drcr=""
	If not rs.EOF then 
		do while not rs.EOF
			iEnNo = iEnNo + 1
			Amount = cdbl(rs(7))
			drcr = rs(2)
			if iEnNo >1 then
				iCtr = iCtr + 1
				if drcr="C" then
					Amount= cdbl(Amount) * -1
				end if
				iTotTaxAmt = CDbl(iTotTaxAmt) + Amount
				iTaxAmt = iTaxAmt & ":" &  Amount
				sTaxName = sTaxName & ":"& rs(8)
				sTaxFlag = true
			end if 'if iEnNo >1 then
			rs.MoveNext 
		loop
		rs.Close
	End If	 
	iCtr = iCtr - 1
	iTaxAmt=mid(iTaxAmt,2)
	sTaxName=mid(sTaxName,2)
	TempTaxAmt = split(iTaxAmt,":")
	TempTaxName = split(sTaxName,":")
	Response.Write iTaxAmt
	Response.Write iCtr &"****" &sTaxName&vbCrLf
	
	Response.Write "sNarrText = "& UBound(sNarrText)
	Response.Write "sTempTaxName = "& UBound(TempTaxName)
'Response.End 
	IF sTaxFlag = True then
		For i =  0 to UBOUND(sNarrText) - 1
		'For i =  0 to UBOUND(TempTaxName) 
			if i > UBound(TempTaxName)  then
				strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(sNarrText(i),aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(" ",aiHeadColWidth(4,2)+1,"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R") & vbcrlf		
			else
				IF i > 0 then
					IF UBound(TempTaxName)  <> 0 then 
						strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(sNarrText(i),aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(TempTaxName(i),aiHeadColWidth(4,2)+1,"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(formatNumber(TempTaxAmt(i)),aiHeadColWidth(4,6),"R") & vbcrlf		
					Else
						strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(sNarrText(i),aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(" ",aiHeadColWidth(4,2)+1,"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R") & vbcrlf		
					End IF
				Else
					strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(sNarrText(i),aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(TempTaxName(i),aiHeadColWidth(4,2)+1,"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(formatNumber(TempTaxAmt(i)),aiHeadColWidth(4,6),"R") & vbcrlf		
				End IF
			end if
			sNarrFlag = true
		Next		
		'--Code for tax--last tax value is not displaying-- 
		if iCtr > i then 
			for j = i to iCtr -1
				strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(TempTaxName(j),aiHeadColWidth(4,2)+1,"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(formatNumber(TempTaxAmt(j)),aiHeadColWidth(4,6),"R") & vbcrlf		
			next
		End if
		'---------------------------------------------------
		strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf
		strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(FormatNumber(iTotTaxAmt),aiHeadColWidth(4,6),"R") & vbcrlf						
		strText=strText & myalign(" ",5,"L") &  myalign(" ",3,"L") &myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",1,"L") &  myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf												
		TotAmt = cdbl(TotAmt) + cdbl(iTotTaxAmt)
	'Response.End 	
	Else
		For i =  0 to UBOUND(sNarrText)  		
			strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(sNarrText(i),aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign(" ",aiHeadColWidth(4,6),"L")  &  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R")&  myalign("",2,"L") & myalign(" ",aiHeadColWidth(4,6),"R") & vbcrlf		
			sNarrFlag = True 
			
		Next
	End IF
	'--code for Rebate if it is available have to less from total amt----
	IF RebFlag = True then
			strText = strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & myalign(" ",aiHeadColWidth(4,3)+8,"L") & myalign(" ",5,"L") & myalign("Less : Rebate",aiHeadColWidth(4,6),"L")  & myalign(" ",aiHeadColWidth(4,6),"R") & myalign(FormatNumber(iRebTaxAmt),aiHeadColWidth(4,6),"R") & vbcrlf		
			TotAmt = cdbl(TotAmt) - cdbl(iRebTaxAmt)
	End If
	'-----------------------------------------------------------------------
	 
	strText =strText & string(90," ") & vbcrlf
	strText =strText & string(90," ") & vbcrlf
	strText =strText & string(90," ") & vbcrlf
	 
	'strText =strtext & string(80,"-") & vbcrlf
	strText=strText & myalign(" ",aiHeadColWidth(4,3) +  aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4)+8 ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf
	'strText =strText & string(50," ") & "Total   " & myalign(FormatNumber(TotAmt,2,0,0,0),15,"R") & vbcrlf
	'strText=strText & myalign("Total ",aiHeadColWidth(4,3) +  aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4) ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber(TotAmt),aiHeadColWidth(4,6),"R") & vbcrlf
	strText=strText & myalign(" ",aiHeadColWidth(4,3) +  aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4)+8 ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign( FormatNumber(TotAmt),aiHeadColWidth(4,6),"R") & vbcrlf
	'strText =strtext & string(80,"-") & vbcrlf
	strText=strText & myalign(" ",aiHeadColWidth(4,3) +  aiHeadColWidth(4,1) + aiHeadColWidth(4,2)+aiHeadColWidth(4,4)+8 ,"L") & myalign(" " ,aiHeadColWidth(4,5),"R") & myalign("-----------",aiHeadColWidth(4,6),"R") & vbcrlf
	
	sAmtWords = "(" &amountwords(totamt)& ")" 
	sAmtWords = BreakString(sAmtWords,30)
	
	for i = 0 to UBOUND(sAmtWords)
		strText =strText & myalign(" ",5,"L") &  myalign(" ",3,"L") & sAmtWords(i)  & vbcrlf	
	Next
	
	strText =strText & " " & vbcrlf
	strText =strText &myalign(" ",55,"L")&formattprint("CONDENSESTART","") & sUnitName &formattprint("CONDENSEEND","")& vbcrlf
	strText =strText & " " & vbcrlf	
	strText =strText & " " & vbcrlf
	strText =strText & " " & vbcrlf
	strText =strText & myalign(sPreparedBy,20,"L") & vbcrlf
	strText =strText & myalign("PreparedBy ",15,"L") & myalign(" ",15,"L") & myalign(" ",15,"L")& myalign(" ",20,"L") & myalign("Authorised Signatory",25,"L")& vbcrlf
	
	'strText = strText &myalign("PreparedBy ",15,"L")
	'strText = strText &myalign(" ",15,"L")
	'strText = strText &myalign("AO / FM",15,"L")
	'strText = strText &myalign(" ",15,"L")
	'strText = strText &myalign("M.D/Director ",30,"L") & vbcrlf  
	sFinalText = sFinalText & strtext & " " & Vbcrlf &" " & vbcrlf 
	'sFinalText = sFinalText & strTextC  
	
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
		strText=strText & myalign("To",aiHeadColWidth(1,1),"L")  & vbcrlf
		strText=strText & myalign("",aiHeadColWidth(1,1),"L") & myalign(name,aiHeadColWidth(1,2),"L") &  myalign("Date   ",aiHeadColWidth(1,3),"L") & ": " &  myalign(dVouDate,aiHeadColWidth(1,4),"L") & vbcrlf
		'strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address1,aiHeadColWidth(1,2),"L") &  myalign("Ref No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sCrVouNo,aiHeadColWidth(1,4),"L") & vbcrlf
		strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address1,aiHeadColWidth(1,2),"L")  & vbcrlf
		IF Cstr(iVouStatus) = "010104" Then		
			strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address2,aiHeadColWidth(1,2),"L") &  myalign("Vou No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign(sAccVouNo ,aiHeadColWidth(1,4),"L") & vbcrlf		
		Else
			strText=strText & string(aiHeadColWidth(1,1)," ") & myalign(Address2,aiHeadColWidth(1,2),"L") &  myalign("Vou No  ",aiHeadColWidth(1,3),"L") & ": " &  myalign("",aiHeadColWidth(1,4),"L") & vbcrlf		
		End IF
		strText=strText & string(aiHeadColWidth(1,1)," ")& city & " - " & Pincode & vbcrlf
		strText=strText & string(aiHeadColWidth(1,1)," ")& State & " " & Country & vbcrlf
		'strText=strText & string(aiHeadColWidth(1,1)," ")& Country & " - " & Pincode & vbcrlf		
		strText=strtext & string(90," ") & vbcrlf
		'strText=strtext & string(aiHeadColWidth(1,0)," ") & "Sir / Sirs We have today credited your account with us as detailed below : " & vbcrlf
		strText=strtext & string(90,"-") & vbcrlf
		strText=strText & centerAlign("S.No",5) & myalign(" ",3,"L") & myalign("Particulars", aiHeadColWidth(4,3)+8,"L") & myalign(" ",3,"L") & centerAlign("  Quantity ",aiHeadColWidth(4,2)+1) &  myalign("",4,"L") & centerAlign("Rate",aiHeadColWidth(4,6)) &  myalign("",2,"L") & centerAlign("  Amount",aiHeadColWidth(4,6)) & vbcrlf
		strText=strtext & string(90,"-") & vbcrlf 
		iLineNo=18
	end Function
'**********************************************************************************************
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
