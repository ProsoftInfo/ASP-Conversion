<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AmdAccGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Maheshwari  S.
	'Created On					:	Oct 28, 2006
	'Description				:	Amedment For Account Cash and Bank Voucher
	'Modified By				:	UmaMaheswari S
	'Modified On				:	12 April 2011
%>
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/AccPopulate.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<!--#include virtual="/Accounts/Transaction/GenerateVou.asp"-->

<%
	Dim rs,rs1,sQuery,iTransNo,sDDrAmt,sDCrAmt,sFinPeriod,sFinFrm,sFinTo,sFinStDate,sFinEndDate
	Dim sSummTrType,sSummCngTy,iRecCnt,sCrVouNo,sTrVouNo,iOldBookNo,iAccTrNo,iAccHead,iUnitHead
	Dim sUnit,sRetValChk,iParAccHead,sVouMon,sVouYr,sVouMonYr,sOldVouDate,sSummChk,sBkCode,sFormVal
	Dim sBkNo,sTbGLVal,sTempVal,dDrAmt,dCrAmt,sContraChk,sRevContraChk,sParChk,sTempConNo,sConToDel
	Dim sTDSChk,sTDSCrDrIndi
	Dim sConVouNo,sCrConVouNo,sNorToContra
	Dim  oDOMNew,Root
	Set oDOMNew = Server.CreateObject("Microsoft.XMLDOM")
	
	iTransNo = Request.Form("hTransNo")
	Response.Write "<p>"& Request.QueryString 
	
	Response.Write "<p>Trans No = "& iTransNo
	
	if trim(iTransNo) ="" then
		iTransNo = Request.QueryString("hTransNo")
	end if 	
	Response.Write "<p>Trans No = "& iTransNo
	
	sNorToContra = Request("hNorToContra")

	IF CStr(sNorToContra) = "" Then
		sNorToContra = "Y"
	End IF
	
	' Response.Write "sNorToContra="&sNorToContra
	' Response.End 
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs1 = server.CreateObject("ADODB.RecordSet")
	
	'To Get Date
	sFinPeriod = Session("FinPeriod")
	sFinFrm = Left(sFinPeriod,4)
	sFinTo = Right(sFinPeriod,4)

	sFinStDate = sFinFrm &"04"
	sFinEndDate = sFinTo &"03"
	
	sFormVal = Session("AmdCas")
	Session("AmdCas") = " "
	
	sQuery =  "Select CreatedVoucherNo,VoucherNumber,BookNumber,TransactionNumber, "&_
			  "TransactionType,OUDefinitionID,Convert(Varchar,VoucherDate,103),BookCode,CrDrIndication "&_
			  "From Acc_T_VoucherHeader Where CreatedTransNo = "&iTransNo&" "
		Response.Write sQuery &"<BR><BR>"	
	rs.Open sQuery,Con
	IF Not rs.EOF Then
		sCrVouNo = rs("CreatedVoucherNo")
		sTrVouNo = rs("VoucherNumber")
		iOldBookNo = rs("BookNumber")
		iAccTrNo = rs("TransactionNumber")
		sSummTrType = rs("TransactionType")
		sUnit = rs("OUDefinitionID")
		sOldVouDate = Trim(rs(6))
		sBkCode = Trim(rs("BookCode"))
		sTDSCrDrIndi = Trim(rs("CrDrIndication"))
	End IF
	rs.Close
	
	sVouMon = Mid(sOldVouDate,4,2)
	sVouYr = Right(sOldVouDate,4)
	sVouMonYr = sVouMon&sVouYr
	
	sSummCngTy = GetSummTransTy(sSummTrType)
	sBkNo = iOldBookNo
	
	Con.BeginTrans
	
	iRecCnt = 0
		
	if trim(iAccTrNo) <> ""  then
		sQuery = "Select Count(1) From Acc_T_VoucherDetails Where AccUnitPartyType Is Not Null "&_
			 "And TransactionNumber = "&iAccTrNo
			 
		Response.Write sQuery 
	

		rs.Open sQuery,Con
		IF Not rs.EOF Then
			iRecCnt = rs(0)
		Else
			iRecCnt = 0
		End IF
		rs.Close
	end if 
		
	IF CStr(iRecCnt) = "0" Then
		sParChk = "F"
		sContraChk = "T" 'It is a Contra 
	Else
		sParChk = "T"
		sContraChk = "F" 'It is Not a Contra
	End IF
	
	
	
	Response.Write "sContraChk="&sContraChk
	Response.Write "<BR><BR>sTDSCrDrIndi="&sTDSCrDrIndi&"<BR><BR>"
	 
	If trim(sContraChk) = "T" and trim(sTDSCrDrIndi) = "D" then
		Set Root = oDOMNew.documentElement
		Set Root = oDOMNew.CreateElement("Root")
		oDOMNew.AppendChild Root
		dim iConTrNo ,Elem
			'Response.Write iTransNo
		sQuery = "Select isNull(ContraTransactionNumber,0) from  Acc_T_Voucherheader where  Createdtransno =  "& iTransNo &" "
		rs.Open sQuery,Con
		IF Not rs.EOF Then
			iConTrNo = rs(0)
		End IF
		rs.Close
		If trim(iConTrNo) <> "0" then
			sQuery = "Select BankInstrumentEntryNo,BankInstrumentType,BankInstrumentNo,convert(VarChar,BankInstrumentDate,103),"&_
					 "InstrumentAmount,PayableAt,DrawnOnBank,isNull(InstrumentEntryNo1,0) from Acc_T_CreatedVoucherInstrumentDet where Createdtransno = "&_
					 "(Select Createdtransno  from Acc_T_Voucherheader where transactionnumber = "& iConTrNo &" ) "
					 Response.Write "<b>"&sQuery &"</b>"
			rs.Open sQuery,Con
			Do while Not rs.EOF 
		
				Set Elem = oDOMNew.createElement("BankInstrumentDet")
				Elem.setAttribute "InsType", rs(1)
				Elem.setAttribute "InsNo",  rs(2)
				Elem.setAttribute "InsDate", rs(3)
				Elem.setAttribute "InsAmt", rs(4)			
				Elem.setAttribute "PayAt",rs(5)
				Elem.setAttribute "DrawnOn", rs(6)
				Elem.setAttribute "Option","N"
				Elem.setAttribute "Action","0"
				Elem.setAttribute "BankInsEntNo", rs(0)
				Root.Appendchild Elem
				rs.MoveNext 
			loop
		
			rs.Close	
			oDOMNew.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-BankInsDet.xml")
		End if	
	End If
	
	
'	Response.End 
'============================ Entry Checking is Over Contra Checking Starts Here ============================================
	
	IF CStr(iRecCnt) = "0" Then
		if trim(iAccTrNo) <> ""  then
			sQuery = "Select Count(1) From Acc_T_VoucherDetails Where TransactionNumber = "&iAccTrNo
			rs.Open sQuery,Con
			IF Not rs.EOF Then
				iRecCnt = rs(0)
			Else
				iRecCnt = 1
			End IF
			rs.Close
		end if 	
	End IF

	IF CStr(iRecCnt) <> "1" Then
		sContraChk = "F" 'Not a Contra
	End IF
	
	IF CStr(iRecCnt) = "1" Then
		sQuery = "Select AccountHead,isNull(AccUnitAccountHead,0) From CashBankView Where TransactionNumber = "&iAccTrNo
		rs.Open sQuery,con
		IF Not rs.EOF Then
			iAccHead = rs(0)
			iUnitHead = rs(1)
		Else
			iAccHead = 0
			iUnitHead = 0
		End IF
		rs.Close
		
		IF CStr(iAccHead) <> "0" and CStr(iUnitHead) <> "0" Then
			sQuery = "Select Count(1) From Acc_M_ContraEntries Where OUDefinitionID = '"&sUnit&"' "&_
					 "And FromAccountHead = "&iAccHead&" And ToAccountHead = "&iUnitHead&" "
					 
			Response.Write sQuery
			rs.Open sQuery,con		 
			IF Not rs.EOF Then
				iRecCnt = rs(0)
			Else
				iRecCnt = 1
			End IF
			rs.Close
			
			IF CStr(iRecCnt) <> "0" Then
				sContraChk = "T"
			Else
				sContraChk = "F"
			ENd IF
			
			IF CStr(sContraChk) = "T" Then 'This is To Check Wheather the Wise Versa Contra is There
				sQuery = "Select Count(1) From Acc_M_ContraEntries Where OUDefinitionID = '"&sUnit&"' "&_
						 "And FromAccountHead = "&iUnitHead&" And ToAccountHead = "&iAccHead&" "
				rs.Open sQuery,con
				IF Not rs.EOF Then
					iRecCnt = rs(0)
				Else
					iRecCnt = 0
				End IF
				rs.Close
			End IF
			
			IF CStr(iRecCnt) <> "0" Then
				sRevContraChk = "T"
			Else
				sRevContraChk = "F"
			ENd IF
			
		End IF	
	End IF



	'store some field information of Acc_T_CreatedVoucherHeader table into xml
	Dim oDOMTemp,RootTemp,sTempVouName
	Dim sOtherApplnTableName,nOtherApplnTransNo,nFromApplication,dtClearedOn
	
	sOtherApplnTableName	= ""
	nOtherApplnTransNo		= ""
	nFromApplication		= ""
	
	sQuery = "Select isNull(OtherApplnTableName,''),isNull(OtherApplnTransNo,0),isNull(FromApplication,0),convert(varchar,ClearedOn,103) From Acc_T_CreatedVoucherHeader  "&_
			 "Where CreatedTransNo = " & iTransNo & " "
	'Response.Write "<p> sQuery = " & sQuery		 
	rs.Open sQuery,Con
	IF Not rs.EOF Then
		sOtherApplnTableName = rs(0)
		
		if trim(rs(1)) = "0" then
			nOtherApplnTransNo = ""
		else
			nOtherApplnTransNo = rs(1)
		end if
		
		
		if trim(rs(2)) = "0" then
			nFromApplication = ""
		else
			nFromApplication = rs(2)
		end if
		
		if trim(rs(3)) = "" or isNull(rs(3)) then
			dtClearedOn = ""
		else
			dtClearedOn = rs(3)
		end if 	
	End IF
	rs.Close
	
	Set oDOMTemp = Server.CreateObject("Microsoft.XMLDOM")
	
	
	if CStr(sBkCode) = "02" Then
		sTempVouName = "BA"
	elseif CStr(sBkCode) = "01" Then
		sTempVouName = "CA"
	end if 
	
	oDOMTemp.Load server.MapPath("../temp/transaction/Voucher Entry_"&sTempVouName&"_"&Session.SessionID&".xml")
	Set RootTemp = oDOMTemp.documentElement
		
	RootTemp.SetAttribute "OtherApplnTableName",sOtherApplnTableName
	RootTemp.SetAttribute "OtherApplnTransNo",nOtherApplnTransNo
	RootTemp.SetAttribute "FromApplication",nFromApplication
	RootTemp.SetAttribute "ClearedOn",dtClearedOn
	
	'Added By UmaMaheswari S,on April 12 2011
	RootTemp.SetAttribute "CreatedVoucherNo",""
	RootTemp.SetAttribute "VoucherNo",""
	RootTemp.SetAttribute "Approver",""

	
	oDOMTemp.Save server.MapPath("../temp/transaction/Voucher Entry_"&sTempVouName&"_"&Session.SessionID&".xml")
	'end
		
'============================ Contra Checking Ends Here =====================================================================	
	
	Response.Write "<b>" & sTDSChk & "=================</b>"
	
	'IF CStr(sTDSChk) = "T" Then
	'	TDSType()
	'End IF
	Response.Write "<p> PartyChk = "& sParChk
	Response.Write "<p> Contra Check = "& sContraChk 
	'Response.End 
	
	IF CStr(sParChk) = "T" Then
		PartyType()
	Elseif CStr(sContraChk) = "F" Then
		AccHeadDeltion()
	Else
		ContraDeletion()
	End IF
	
	CommonDel()
	
	'Response.Clear
	Response.Write "<br><br><b>*****************Deletion is Over*********************</b><br><br>"
	
	
	IF CStr(Trim(sCrConVouNo)) = "" Then
		sCrConVouNo = "0"
	End IF

	IF CStr(Trim(sConVouNo)) = "" Then
		sConVouNo = "0"
	End IF
	
	
'=================== Deletion Of Records is Over ====================================================
	IF CStr(sBkCode) = "02" Then
	Response.Write "<p> Welcome to Bank To Cash"
		sRetValChk = CreateVou("02","BA",sCrVouNo,sTrVouNo,iOldBookNo,sOldVouDate,iTransNo,iAccTrNo,sCrConVouNo,sConVouNo,sNorToContra)
	Elseif CStr(sBkCode) = "01" Then
	Response.Write "<p> Welcome to Cash To Bank"
		sRetValChk = CreateVou("01","CA",sCrVouNo,sTrVouNo,iOldBookNo,sOldVouDate,iTransNo,iAccTrNo,sCrConVouNo,sConVouNo,sNorToContra)
	End IF 
	
	'Response.Write "Test..."
	'Response.Write "<br>"&sRetValChk
	'Response.End 
	
	sTempVal = Split(sRetValChk,":")
	IF CStr(sTempVal(0)) = "T" and CStr(sRevContraChk) = "T" Then
		Dim adoCmd
		sQuery = "GLContraUpdate"
		Set adoCmd = Server.CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection =con
		adoCmd.CommandText = sQuery
		adoCmd.CommandType = 4
		adoCmd.Parameters.Append adoCmd.CreateParameter("@FromTransNo",3,1,3,sTempVal(1))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@ToTransNo",3,1,3,sTempVal(2))
		adoCmd.Execute()
	End IF
	sTbGLVal = CheckTBGL(sUnit)
	sTempVal = Split(sTbGLVal,":")
	'Response.Write sTbGLVal
	dDrAmt = sTempVal(0)
	dCrAmt = sTempVal(1)

	dDrAmt = FormatNumber(dDrAmt,2,,,0)
	dCrAmt = FormatNumber(dCrAmt,2,,,0)
	IF CStr(dDrAmt) = CStr(dCrAmt) Then
			
	'	Con.RollbackTrans
	'	Response.Write "<p>OK</p>"
	'	Response.End
		
	    con.CommitTrans
		
		IF CStr(sBkCode) = "02" Then
			'Response.Redirect "ACCBANKVOUCHERS.ASP?hFormVal="&sFormVal
			Response.Redirect "MANAGEBANKVOUCHERS.ASP"
		Elseif CStr(sBkCode) = "01" Then
			'Response.Redirect "ACCCASHVOUCHERS.ASP?hFormVal="&sFormVal
			Response.Redirect "MANAGECASHVOUCHERS.ASP"
		End IF
	Else
		con.RollbackTrans
		Response.write  "<b>Debit Amount and Credit Amount does'nt Match Transaction Rolled back </b><br><br>"
		Response.write  "DR AMOUNT   : " & dDrAmt &"<br>"
		Response.write  "CR AMOUNT   : " & dCrAmt &"<br>"
		Response.write  "<B>DIFFERENCES : " & FormatNumber(Round(CDbl(dDrAmt) - CDbl(dCrAmt),2),2,,,0) &"<br>"
		IF CStr(sTempVal(2)) = "L" Then
			Response.write  "<B>Differences In : Ledger </b></br>"
		Elseif CStr(sTempVal(2)) = "T" Then
			Response.write  "<B>Differences In : Trial Balance </b></br>"
		End IF
	End IF
%>

<%
	Function TDSType()
		'Response.Clear
		Response.Write "Inside TDS  Type <br><br>"
		
		sQuery = "Select T.PartyType,T.PartySubType,T.PartyCode,V.AccountHead,Sum(T.Amount) Amount, " & _
				 "Convert(VarChar,T.VoucherDate,103) VoucherDate " & _
				 "From Acc_T_PartyTransactions T,vwOrgPartyType V Where T.PartyType = V.PartyType " & _
				 "And T.PartySubType = V.PartySubType And T.OUDefinitionID = V.OUDefinitionID " & _
				 "And T.TransCrDrIndication = '"&sTDSCrDrIndi&"' And T.TransactionNumber = "&iAccTrNo&" " & _
				 "Group By T.PartyType,T.PartySubType,T.PartyCode,V.AccountHead,T.VoucherDate "
				 
		Response.Write sQuery &"<br><br>"
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF
			IF CStr(sTDSCrDrIndi) = "C" Then
				sQuery = "Update Acc_T_PartyTransactAmt Set MonthCrAmount = MonthCrAmount - "&rs("Amount")&" "&_
						 "Where PartyType = '"&rs("PartyType")&"' and PartySubType = "&rs("PartySubType")&" and PartyCode = "&rs("PartyCode")&" and " & _
						 "MonthYear = '"&sVouMonYr&"' and OUDefinitionID = '"&sUnit&"' "
			Else
				sQuery = "Update Acc_T_PartyTransactAmt Set MonthDrAmount = MonthDrAmount - "&rs("Amount")&" "&_
						 "Where PartyType = '"&rs("PartyType")&"' and PartySubType = "&rs("PartySubType")&" and PartyCode = "&rs("PartyCode")&" and " & _
						 "MonthYear = '"&sVouMonYr&"' and OUDefinitionID = '"&sUnit&"' "
			End IF
			
			Response.Write sQuery &"<br><br>"
			Con.Execute (sQuery)
			
			IF CStr(sTDSCrDrIndi) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&rs("Amount")&" "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" And MonthYear = '"&sVouMonYr&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&rs("Amount")&" "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" And MonthYear = '"&sVouMonYr&"' "
			End IF
			Con.Execute (sQuery)
			
			sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&rs("Amount")&" "&_
					 "Where AccountHead = "&rs("AccountHead")&" And TransCrDrIndication = '"&sTDSCrDrIndi&"' " &_
					 "And OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" " &_
					 "And TransactionType = '"&sSummCngTy&"' And "&_
					 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" and Amount - "&CDbl(rs("Amount")) & " >= 0"		
					
			Response.Write sQuery &"<br><br>" 
			Con.Execute (sQuery)
			
			rs.MoveNext
		loop
		rs.Close
		
	End Function
%>

<%
	Function PartyType()
		'Response.Clear
		Response.Write "Ïnside Party<br><br>  "
		sQuery = "Select isNull(AccUnitAccountHead,0) AccHead,isNull(AccUnitPartyType,'0') ParType, "&_
				 "isNull(AccUnitPartySubType,0) ParSubType,isNull(AccUnitPartycode,0) ParCode, "&_
				 "Amount,TransCrDrIndication,VoucherEntryNumber From  "&_
				 "Acc_T_VoucherDetails  Where TransactionNumber = "&iAccTrNo
				 
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF
			IF CStr(rs("AccHead")) = "0" Then
				iParAccHead = GetPartyAccHead(rs("ParType"),rs("ParSubType"),rs("ParCode"),sUnit)
				IF CStr(rs("TransCrDrIndication")) = "C" Then
					sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&iParAccHead&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				Else
					sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&iParAccHead&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				End IF	
				Response.Write "<br> Party Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
				
				IF CStr(rs("TransCrDrIndication")) = "C" Then
					sQuery = "Update Acc_T_PartyTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And PartyType = '"&rs("ParType")&"' "&_
							 "And PartySubType = "&rs("ParSubType")&" And PartyCode = "&rs("ParCode")&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				Else
					sQuery = "Update Acc_T_PartyTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And PartyType = '"&rs("ParType")&"' "&_
							 "And PartySubType = "&rs("ParSubType")&" And PartyCode = "&rs("ParCode")&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				End IF	
				Response.Write "<br> Party Reduction IN TB:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
				
				sQuery = "Update Acc_T_PartyTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "TransactionNumber = "&iAccTrNo&" and VoucherEntryNumber = "&rs("VoucherEntryNumber")&" and "&_
						 "Amount - "&CDbl(rs("Amount")) & " >= 0"	
				Response.Write "<br> Party Reduction IN Party Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery	
				
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
						 "TransCrDrIndication = '"&rs("TransCrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&iParAccHead&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" and Amount - "&CDbl(rs("Amount")) & " >= 0"		
				
				Response.Write "<br> Control A/C Reduction IN GL Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery	
				
			Else
				IF CStr(rs("TransCrDrIndication")) = "C" Then
					sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				Else
					sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("Amount")) & " "&_
							 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
							 "And MonthYear = '"&sVouMonYr&"' "
				End IF	
				Response.Write "<br> Party Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
				
				sSummChk = SummaryCheck(rs("AccHead"),sUnit)
				IF CStr(sSummChk) = "Y" Then 'The Account head is Of Summary Type Only Then
					sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
							 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
							 "TransCrDrIndication = '"&rs("TransCrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
							 "TransactionNumber = 0 And AccountHead = "&rs("AccHead")&" And  "&_
							 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
				Else
					sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
							 "TransactionNumber = "&iAccTrNo&" And Accounthead = "&rs("AccHead")&" "&_
							 "And TransCrDrIndication = '"&rs("TransCrDrIndication")&"' and Amount - "&CDbl(rs("Amount")) & " >= 0 "		
				End IF
				 Response.Write "<br> Party Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
			End IF 'rs("AccHead")) Check 
			rs.MoveNext
		Loop
		rs.Close
		
		sQuery = "Select AccountHead,VoucherAmount,CrDrIndication From  "&_
				 "Acc_T_VoucherHeader Where TransactionNumber = "&iAccTrNo&" "
				 
		Response.write sQuery & "<br><br>"
		
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		IF Not rs.EOF Then
			IF CStr(rs("CrDrIndication")) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			End IF	
			Response.Write "<br> GL Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
				
			sSummChk = SummaryCheck(rs("AccountHead"),sUnit)
			Response.Write rs("AccountHead") &"  " & sUnit &"  " & sSummChk &"<br><br><br>"
			
			IF CStr(sSummChk) = "Y" Then 'The Account head is Of Summary Type Only Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
						 "TransCrDrIndication = '"&rs("CrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&rs("AccountHead")&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
			Else
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccountHead")&" "&_
						 "And TransCrDrIndication = '"&rs("CrDrIndication")&"' "	
			End IF
			Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
		End IF	
		rs.Close		
	End Function
%>

<%
	Function AccHeadDeltion()
		Dim iRecAff,iTempAccNo
		'Response.Clear
		Response.Write "<b>Inside Account Head Deletion </b><br><br><br>"
		sQuery = "Select AccUnitAccountHead AccHead,Amount,TransCrDrIndication,VoucherEntryNumber From  "&_
				 "Acc_T_VoucherDetails  Where TransactionNumber = "&iAccTrNo
				 
		'Response.Write sQuery
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF
			IF CStr(rs("TransCrDrIndication")) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("Amount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("Amount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			End IF	
			'Response.Write "<br> GL Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
				
			sSummChk = SummaryCheck(rs("AccHead"),sUnit)
			IF CStr(sSummChk) = "Y"  Then 'The Account head is Of Summary Type Only Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
						 "TransCrDrIndication = '"&rs("TransCrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&rs("AccHead")&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
				Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
			Else
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccHead")&" "&_
						 "And TransCrDrIndication = '"&rs("TransCrDrIndication")&"' And VoucherEntryNumber = "&rs("VoucherEntryNumber")				
				Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery,iRecAff
				
				Response.Write "Recond Count  " & iRecAff
				iTempAccNo = 0
				IF CStr(iRecAff) = "0" Then
					iTempAccNo = Cint(iAccTrNo) + 1
					sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
							 "TransactionNumber = "&iTempAccNo&" And AccountHead = "&rs("AccHead")&" "&_
							 "And TransCrDrIndication = '"&rs("TransCrDrIndication")&"' "						
					Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
					Con.Execute sQuery,iRecAff
				End IF
				
			End IF
			
			rs.MoveNext
		Loop
		rs.Close
		
		sQuery = "Select AccountHead,VoucherAmount,CrDrIndication,Convert(Datetime,VoucherDate,103) VoucherDate From  "&_
				 "Acc_T_VoucherHeader Where TransactionNumber = "&iAccTrNo&" "
		
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		IF Not rs.EOF Then
			IF CStr(rs("CrDrIndication")) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
						 
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			End IF	
			
			'Response.Write "<br> GL Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
				
			sSummChk = SummaryCheck(rs("AccountHead"),sUnit)
			IF CStr(sSummChk) = "Y" Then 'The Account head is Of Summary Type Only Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
						 "TransCrDrIndication = '"&rs("CrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&rs("AccountHead")&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
				Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
			Else
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccountHead")&" "&_
						 "And TransCrDrIndication = '"&rs("CrDrIndication")&"' "						
				Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery,iRecAff
				
				Response.Write "Recond Count  " & iRecAff
				
				IF CStr(iRecAff) = "0" Then
					iTempAccNo = iAccTrNo + 1
					sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
							 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccountHead")&" "&_
							 "And TransCrDrIndication = '"&rs("CrDrIndication")&"' "						
					Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
					Con.Execute sQuery,iRecAff
				End IF
			End IF
			
			
			
		End IF
		rs.Close
	End Function
%>

<%
	Function SummaryCheck(iRecAccHead,sOrgUnit)
		Dim rs3,sQry,sSummTy
		Set rs3 = Server.CreateObject("ADODB.Recordset")
		sQry = "Select SummaryPosting From VwOrgGLHeads Where AccountHead = "&iRecAccHead&" And OUDefinitionID = '"&sOrgUnit&"' "		
		rs3.Open sQry,con
		IF Not rs3.EOF Then
			sSummTy = rs3("SummaryPosting")
		End IF			
		rs3.Close
		IF CStr(sSummTy) = "1" Then
			SummaryCheck = "Y"
		Else
			SummaryCheck = "N"
		End IF
	End Function
%>

<%
	Function GetSummTransTy(sSumTrType)
		Dim sChgSumTy
		Select Case Cstr(sSumTrType)
			Case "CAR" sChgSumTy = "CSR"
			Case "CAP" sChgSumTy = "CSP"
			Case "BAR" sChgSumTy = "BSR"
			Case "BAP" sChgSumTy = "BSP"
			Case "PJR" sChgSumTy = "PJS"
			Case "SJP" sChgSumTy = "SJS"
			Case "CNR" sChgSumTy = "CNS"
			Case "DNR" sChgSumTy = "DNS"
			Case "GJR" sChgSumTy = "GJS"
		End Select
		
		GetSummTransTy = sChgSumTy
	End Function
%>

<%
	Function CommonDel()
		Response.Write "<BR><br>-----------CONTRA DELETION STARTS HERE ------------------<br>"
		
		Dim iFrmAccHead,iToAccHead,iFrmAmt,sFrmDate,sFrmCrDrIndi,iContraTrNo,iCrContraTrNo
		Dim iTempTrNo,sToCrDrIndi
		
		sQuery = "Select PayablesNumber From Acc_T_PybleAdjustmentDetails  "&_
				 "Where AdjustType Is Null And PaidByTransactionNo = "&iAccTrNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Delete From Acc_T_PybleAdjustmentDetails Where PayablesNumber = "&rs("PayablesNumber")&"  "&_
					 "And PaidByTransactionNo = "&iAccTrNo
					 Response.Write sQuery &"<BR>"
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Payables Over **********************************************************************************
		sQuery = "Select ReceivableNumber From Acc_T_RcvblAdjustmentDetails  "&_
				 "Where AdjustType Is Null And RecdByTransactionNo = "&iAccTrNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Delete From Acc_T_RcvblAdjustmentDetails Where ReceivableNumber = "&rs("ReceivableNumber")&"  "&_
					 "And RecdByTransactionNo = "&iAccTrNo
					 
			Response.Write sQuery &"<br> "
			
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Receivables Over **********************************************************************************		
		sQuery = "Select PayablesNumber,AmountPaid From Acc_T_PybleAdjustmentDetails  "&_
				 "Where AdjustType Is NOT Null And PaidByTransactionNo = "&iAccTrNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = AdvanceAdjusted - "&rs("AmountPaid")&"  "&_
					 "Where AdvanceNumber = "&rs("PayablesNumber")&" "
					 Response.Write sQuery &"<BR>"
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Advance Payables Over **********************************************************************************				
		sQuery = "Select ReceivableNumber,AmountReceived From Acc_T_RcvblAdjustmentDetails  "&_
				 "Where AdjustType Is Null And RecdByTransactionNo = "&iAccTrNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Update Acc_T_AdvancePayments Set AdvanceAdjusted = AdvanceAdjusted - "&rs("AmountReceived")&"  "&_
					 "Where AdvanceNumber = "&rs("ReceivableNumber")&" "
					 Response.Write sQuery &"<BR>"
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
	
'***************************** Advance Receivables Over **********************************************************************************				
		IF CStr(sContraChk) = "T" Then
			sQuery = "Select AccountHead,AccUnitAccountHead,Amount,Convert(Varchar,VoucherDate,103), "&_
					 "CrDrIndication,ContraTransactionNumber From CashBankView Where TransactionNumber = "&iAccTrNo
			rs.Open sQuery,Con
			IF Not rs.EOF Then
				iFrmAccHead = rs(0)
				iToAccHead = rs(1)
				iFrmAmt = rs(2)
				sFrmDate = Trim(rs(3))
				sFrmCrDrIndi = rs(4)
				iTempTrNo =  rs(5)
			End IF
			rs.Close
			
			'iTempTrNo = Cint(iAccTrNo) + 1
			IF CStr(sFrmCrDrIndi) = "C" Then
				sToCrDrIndi = "D"
			Else
				sToCrDrIndi = "C"
			End IF
			
			sQuery = "Select TransactionNumber From Acc_T_VoucherHeader Where TransactionNumber = "&iTempTrNo&" "&_
					 "And AccountHead = "&iToAccHead&" And VoucherAmount = "&iFrmAmt&" And  "&_
					 "Convert(Varchar,VoucherDate,103)= '"&sFrmDate&"' And CrDrIndication = '"&sToCrDrIndi&"' "
					 
			sTempConNo = GetContraNumbers(iAccTrNo)
			sConToDel = Split(sTempConNo,"|")
		'Response.Write "<BR>sTempConNo="&sTempConNo&"<BR>"
			iTempTrNo = sConToDel(0)
	'Response.Write "<BR>"&iTempTrNo &"="&sTempConNo&"<BR>"
			IF Cstr(iTempTrNo) = "0" Then
				'Response.Clear
				Con.RollbackTrans
				Response.Write "<br><b>Error On Contra Deletion</b>"
				Response.End
			End IF

			sQuery = "Delete from Acc_T_GLTransSummaryBreakup where TransactionNumber = "&iTempTrNo&" "
			Response.write  sQuery &"<BR>"
			con.execute sQuery
			
			'Newly added on 21st feb 09 bec of Additional entry in GLTrans table with 0 value
			sQuery = "Delete from Acc_T_GLTransactions where TransactionNumber  = "&iTempTrNo&" "
			Response.write  sQuery &"<BR>"
			con.execute sQuery
			

			sQuery = "Delete from Acc_T_VoucherDetails where TransactionNumber = "&iTempTrNo&" "
			Response.write  sQuery &"<BR>"
			con.execute sQuery

			sQuery = "Delete from Acc_T_VoucherHeader where TransactionNumber = "&iTempTrNo&" "
			Response.write  sQuery&"<BR>"
			con.execute sQuery

			iTempTrNo = sConToDel(1)

			IF Cstr(iTempTrNo) = "0" Then
				'Response.Clear
				Con.RollbackTrans
				Response.Write "<br><b>Error On Contra Deletion</b>"
				Response.End
			End IF
			
			sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet  where CreatedTransNo =  "&iTempTrNo&" "
			Response.write  sQuery &"<BR>"
			con.execute sQuery
			

			sQuery = "Delete from Acc_T_CreatedVoucherDetails where CreatedTransNo =  "&iTempTrNo&" "
			Response.write  sQuery &"<BR>"
			con.execute sQuery

			sQuery = "Delete from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iTempTrNo&" "
			Response.write  sQuery &"<BR><BR><BR>"
			con.execute sQuery
		End IF	
'****************************** Contra Entry Deletion is Over ****************************************************************************
		sQuery = "Delete from Acc_T_GLTransSummaryBreakup Where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_GLTransactions where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_PartyTransactions where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_PybleAdjustmentDetails where PaidByTransactionNo = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
				
		sQuery = "Delete from Acc_T_RcvblAdjustmentDetails where RecdByTransactionNo = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
				
		sQuery = "Delete from Acc_T_AdvancePayments where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
				
		sQuery = "Delete from Acc_T_VoucherDetails where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
				
		sQuery = "Delete from Acc_T_VoucherHeader where TransactionNumber = "&iAccTrNo&" "
		Response.write  sQuery&"<BR>"
		con.execute sQuery
'*********************************** Accounted Table Deletion is Over ***********************************
		sQuery = "Select PayablesNumber From Acc_T_CreatedPybleAdjDet  "&_
				 "Where AdjustType Is Null And CreatedTransNo = "&iTransNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Delete From Acc_T_CreatedPybleAdjDet Where PayablesNumber = "&rs("PayablesNumber")&"  "&_
					 "And CreatedTransNo = "&iTransNo
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Payables Over **********************************************************************************
		sQuery = "Select ReceivableNumber From Acc_T_CreatedRcvbleAdjDet  "&_
				 "Where AdjustType Is Null And CreatedTransNo = "&iTransNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Delete From Acc_T_CreatedRcvbleAdjDet Where ReceivableNumber = "&rs("ReceivableNumber")&"  "&_
					 "And CreatedTransNo = "&iTransNo
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Receivables Over **********************************************************************************		
		sQuery = "Select PayablesNumber,AmountPaid From Acc_T_CreatedPybleAdjDet  "&_
				 "Where AdjustType Is NOT Null And CreatedTransNo = "&iTransNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&rs("AmountPaid")&"  "&_
					 "Where CreatedAdvanceNo = "&rs("PayablesNumber")&" "
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
'***************************** Advance Payables Over **********************************************************************************				
		sQuery = "Select ReceivableNumber,AmountReceived From Acc_T_CreatedRcvbleAdjDet  "&_
				 "Where AdjustType Is Null And CreatedTransNo = "&iTransNo
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Do While Not rs.EOF 
			sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = AdvanceAdjusted - "&rs("AmountReceived")&"  "&_
					 "Where CreatedAdvanceNo = "&rs("ReceivableNumber")&" "
			Con.Execute sQuery
			rs.MoveNext
		Loop
		rs.Close
	
'***************************** Advance Receivables Over **********************************************************************************	
		sQuery = "Delete from Acc_T_CreatedPybleAdjDet where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_CreatedRcvbleAdjDet where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_CreatedAdvances where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
		
		sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet  where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
		
			
		sQuery = "Delete from Acc_T_CreatedVoucherDetails where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR>"
		con.execute sQuery
	
		sQuery = "Delete from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iTransNo&" "
		Response.write  sQuery &"<BR><BR><BR>"
		con.execute sQuery
		
		'IF CStr(sContraChk) = "T" Then
		'	
		'End IF
		
		
	End Function
%>

<%
	Function GetPartyAccHead(sParTy,sParSubTy,sParCode,sOrgID)
		Dim rs3,sQry,iParAccHeadCode
		Set rs3 = Server.CreateObject("ADODB.Recordset")
		sQry =  "Select AccountHead From vwOrgPartyType Where PartyType = '"&sParTy&"' "&_
				"And PartySubType = "&sParSubTy&" And OUDefinitionID = '"&sOrgID&"' "
		rs3.Open sQry,con
		IF Not rs3.Eof Then
			iParAccHeadCode = rs3(0)
		End IF
		rs3.close
		GetPartyAccHead = iParAccHeadCode
	End Function
%>

<%
	Function ContraDeletion()
		Dim iRecAff
		'Response.Clear
		Response.Write "<b>*************** Inside Contra Deletion **********************</b><br><br>"
		Dim sToBookCode,iToBookNo,sFrmTrType,sToTrType,rs3,iTempAccNo
		Set rs3 = Server.CreateObject("ADODB.RECORDSET")
		
		sQuery = "Select AccountHead,VoucherAmount,CrDrIndication,Convert(Datetime,VoucherDate,103) VoucherDate,TransactionType,ContraTransactionNumber From  "&_
				 "Acc_T_VoucherHeader Where TransactionNumber = "&iAccTrNo&" "
		Response.Write "<p>"& sQuery
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		IF Not rs.EOF Then
			sFrmTrType = rs("TransactionType")
			iTempAccNo = rs("ContraTransactionNumber")
			Response.Write "<p> Contra TransNo =  " & iTempAccNo &"<BR>"
			IF CStr(rs("CrDrIndication")) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
						 
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("VoucherAmount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccountHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			End IF	
			
			Response.Write "<p> GL Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
				
			sSummChk = SummaryCheck(rs("AccountHead"),sUnit)
			IF CStr(sSummChk) = "Y" Then 'The Account head is Of Summary Type Only Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sBkCode&"' And BookNumber = "&sBkNo&" And  "&_
						 "TransCrDrIndication = '"&rs("CrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&rs("AccountHead")&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
				Response.Write "<p> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
			Else
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
						 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccountHead")&" "&_
						 "And TransCrDrIndication = '"&rs("CrDrIndication")&"' And " & _
						 "Amount - "&CDbl(rs("VoucherAmount"))&" >=0 "
						 
				Response.Write "<p> New Test:-   " & sQuery &"<br><br>"
				Con.Execute sQuery,iRecAff
				
				'Response.Write "Recond Count  " & iRecAff
				
				'IF CStr(iRecAff) = "0" Then
				'	iTempAccNo = iAccTrNo + 1
				'	sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("VoucherAmount")) & " Where  "&_
				'			 "TransactionNumber = "&iAccTrNo&" And AccountHead = "&rs("AccountHead")&" "&_
				'			 "And TransCrDrIndication = '"&rs("CrDrIndication")&"' "						
				'	Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				'	Con.Execute sQuery,iRecAff
				'End IF
				
			End IF
		End IF		
		rs.Close
		
'============================ Details Table Entry Starts Here ==================================================
		sQuery = "Select AccUnitAccountHead AccHead,Amount,TransCrDrIndication From  "&_
				 "Acc_T_VoucherDetails  Where TransactionNumber = "&iAccTrNo
				 
		Response.Write "<p>"& sQuery
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		IF Not rs.EOF Then
			sQuery = "Select Bookcode,BookNumber From VwOrgBookNames Where BookAccountHead = "&rs("AccHead")&" "&_
					 "and OUDefinitionID = '"&sUnit&"' "
					 
			'Response.Write sQuery
			rs3.Open sQuery,Con
			IF Not rs3.Eof Then
				sToBookCode = rs3(0)
				iToBookNo = rs3(1)
			End IF
			rs3.Close
			
			IF CStr(sFrmTrType) = "BAP" and CStr(sToBookCode) = "02" Then
				sToTrType = "BAR"
			ElseIF 	CStr(sFrmTrType) = "BAP" and CStr(sToBookCode) = "01" Then
				sToTrType = "CAR"
			ElseIF 	CStr(sFrmTrType) = "BAR" and CStr(sToBookCode) = "02" Then
				sToTrType = "BAP"
			ElseIF 	CStr(sFrmTrType) = "BAR" and CStr(sToBookCode) = "01" Then
				sToTrType = "CAP"
			End IF
			
			IF CStr(sFrmTrType) = "CAP" and CStr(sToBookCode) = "02" Then
				sToTrType = "BAR"
			ElseIF 	CStr(sFrmTrType) = "CAP" and CStr(sToBookCode) = "01" Then
				sToTrType = "CAR"
			ElseIF 	CStr(sFrmTrType) = "CAR" and CStr(sToBookCode) = "02" Then
				sToTrType = "BAP"
			ElseIF 	CStr(sFrmTrType) = "CAR" and CStr(sToBookCode) = "01" Then
				sToTrType = "CAP"
			End IF
				
					 
			IF CStr(rs("TransCrDrIndication")) = "C" Then
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthCrAmount = MonthCrAmount - "&CDbl(rs("Amount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			Else
				sQuery = "Update Acc_T_GLAccTransactAmt Set MonthDrAmount = MonthDrAmount - "&CDbl(rs("Amount")) & " "&_
						 "Where OUDefinitionID = '"&sUnit&"' And AccountHead = "&rs("AccHead")&" "&_
						 "And MonthYear = '"&sVouMonYr&"' "
			End IF	
			Response.Write "<p> GL Account Head Reduction IN TB:-   " & sQuery &"<br><br>"
			Con.Execute sQuery
				
			sSummChk = SummaryCheck(rs("AccHead"),sUnit)
			sSummCngTy = GetSummTransTy(sToTrType)
			IF CStr(sSummChk) = "Y"  Then 'The Account head is Of Summary Type Only Then
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "OUDefinitionID = '"&sUnit&"' And BookCode = '"&sToBookCode&"' And BookNumber = "&iToBookNo&" And  "&_
						 "TransCrDrIndication = '"&rs("TransCrDrIndication")&"' And TransactionType = '"&sSummCngTy&"' And "&_
						 "TransactionNumber = 0 And AccountHead = "&rs("AccHead")&" And  "&_
						 "Month(VoucherDate) = "&sVouMon&" And Year(VoucherDate) = "&sVouYr&" "		
				Response.Write "<p> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				Con.Execute sQuery
			Else
				'iTempAccNo = Cint(iAccTrNo) + 1
				sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
						 "TransactionNumber = "&iTempAccNo&" And AccountHead = "&rs("AccHead")&" "&_
						 "And TransCrDrIndication = '"&rs("TransCrDrIndication")&"' "				
				Response.Write "<p> Test :-   " & sQuery &"<br><br>"
				Con.Execute sQuery,iRecAff
				
				'Response.Write "Recond Count  " & iRecAff
				'iTempAccNo = 0
				'IF CStr(iRecAff) = "0" Then
				'	iTempAccNo = Cint(iAccTrNo) - 1
				'	sQuery = "Update Acc_T_GLTransactions Set Amount = Amount - "&CDbl(rs("Amount")) & " Where  "&_
				'			 "TransactionNumber = "&iTempAccNo&" And AccountHead = "&rs("AccHead")&" "&_
				'			 "And TransCrDrIndication = '"&rs("TransCrDrIndication")&"' "						
				'	Response.Write "<br> GL Account Head Reduction In Ledger:-   " & sQuery &"<br><br>"
				'	Con.Execute sQuery,iRecAff
				'End IF
				
			End IF
		End IF	
		rs.Close
		Response.Write "<br>Checking..................<br>"
		'Response.End 
			
	End Function
%>

<%
	Function GetContraNumbers(iDelTrNo)
		'Response.Clear
		Response.write "ïDelTrNo   = "& iDelTrNo &"<br><br><br>"
		
		Dim sNewOrgID,iNewAccHd,iNewVouDate,iNewVouAmt,sNewCrDrIndi,iNewCrBy,sNewCrOn,iNewTempNo
		Dim iNewContraNo,iNewCrContraNo,sRetVal
		iNewContraNo = 0
		iNewCrContraNo = 0
				
		sQuery = "Select OUDefinitionID,AccUnitAccountHead,Convert(Varchar,VoucherDate,103) "&_
				 "VoucherDate,VoucherAmount,TransCrDrIndication,CreatedBy,Convert(Varchar,CreatedOn,103) CreatedOn,ContraTransactionNumber  From  "&_
				 "CashBankView Where TransactionNumber = "&iDelTrNo
				 
		Response.write sQuery & "<br><br>"
		With rs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = Con
			.Open
		End With
		Set rs.ActiveConnection = Nothing
		Response.write rs.RecordCount
		IF Not rs.EOF Then
			Response.write "Ïnside"
			sNewOrgID = rs("OUDefinitionID")
			iNewAccHd = rs("AccUnitAccountHead")
			iNewVouDate = rs("VoucherDate")
			iNewVouAmt = rs("VoucherAmount")
			sNewCrDrIndi = rs("TransCrDrIndication")
			iNewCrBy = rs("CreatedBy")
			sNewCrOn = rs("CreatedOn")
			iNewTempNo = rs("ContraTransactionNumber")
		End IF
		rs.Close
		'Response.Write "iNewTempNo="&iNewTempNo
		'iNewTempNo = CInt(iDelTrNo) + 1
		
		'sQuery = "Select TransactionNumber,CreatedTransNo,VoucherNumber,CreatedVoucherNo  From Acc_T_VoucherHeader Where OUDefinitionID = '"&sNewOrgID&"' "&_
		'		 "And AccountHead = "&iNewAccHd&" And Convert(Varchar,VoucherDate,103) = '"&iNewVouDate&"' "&_
		'		 "And VoucherAmount = "&iNewVouAmt&" And CrDrIndication = '"&sNewCrDrIndi&"' And CreatedBy = "&iNewCrBy&" "&_
		'		 "And Convert(Varchar,CreatedOn,103) = '"&sNewCrOn&"' And TransactionNumber = "&iNewTempNo
		
		
		sQuery = "Select TransactionNumber,CreatedTransNo,VoucherNumber,CreatedVoucherNo  From Acc_T_VoucherHeader Where OUDefinitionID = '"&sNewOrgID&"' "&_
				 "And AccountHead = "&iNewAccHd&" And Convert(Varchar,VoucherDate,103) = '"&iNewVouDate&"' "&_
				 "And VoucherAmount = "&iNewVouAmt&" And CrDrIndication = '"&sNewCrDrIndi&"' And CreatedBy = "&iNewCrBy&" "&_
				 "And Convert(Varchar,CreatedOn,103) = '"&sNewCrOn&"' And TransactionNumber = "&iNewTempNo
				 
		Response.write "<br>"&squery&"<br>"
		rs.Open sQuery,Con
		IF Not rs.EOF Then
			iNewContraNo = rs("TransactionNumber")	
			iNewCrContraNo = rs("CreatedTransNo")
			sConVouNo = rs("VoucherNumber")
			sCrConVouNo = rs("CreatedVoucherNo")
		Else
			iNewContraNo = 0
			iNewCrContraNo = 0
		End IF
		rs.Close	 
		
		'Response.write 
		IF CStr(iNewContraNo) = "0" Then
			'iNewTempNo = CInt(iDelTrNo) - 1
		
			sQuery = "Select ContraTransactionNumber,CreatedTransNo,VoucherNumber,CreatedVoucherNo From Acc_T_VoucherHeader Where OUDefinitionID = '"&sNewOrgID&"' "&_
					 "And AccountHead = "&iNewAccHd&" And Convert(Varchar,VoucherDate,103) = '"&iNewVouDate&"' "&_
					 "And VoucherAmount = "&iNewVouAmt&" And CrDrIndication = '"&sNewCrDrIndi&"' And CreatedBy = "&iNewCrBy&" "&_
					 "And Convert(Varchar,CreatedOn,103) = '"&sNewCrOn&"' And TransactionNumber = "&iNewTempNo
			
			Response.write "<br>"&squery &"<br>"
			rs.Open sQuery,Con
			IF Not rs.EOF Then
				iNewContraNo = rs("ContraTransactionNumber")	
				iNewCrContraNo = rs("CreatedTransNo")
				sConVouNo = rs("VoucherNumber")
				sCrConVouNo = rs("CreatedVoucherNo")
			Else
				iNewContraNo = 0
				iNewCrContraNo = 0
			End IF
			rs.Close	 
		End IF
		
		sRetVal = iNewContraNo&"|"&iNewCrContraNo&"|"&sConVouNo&"|"&sCrConVouNo
		Response.Write sRetVal 
		'Response.End 
		GetContraNumbers = sRetVal
	End Function
%>