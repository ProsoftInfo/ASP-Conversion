<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouAmdGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	January 09 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<%

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery
dim EntryNode,HeaderNode,nodANL,ObjDom

dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp

dim sOrgId,sBookNo,sVouType,sVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt,iAccHeadNo
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,iAmdendmentNo,objfs,bActionFlag
dim sAccHeadCode,sPayTo,bAddFlag
dim sBkInsDate,sBkInsNo,sBkInsType,sBkPayableat,sBkDrawnOn
dim sExp,tempNode,iTdsAmount,iTdsPer
Dim iAdvNo,iAppBy,sRecpNo,DelRoot
Dim iPayNo,dAdjAmt,iCrAdvNo,sAdjTy,iCtr
Dim sFormVal,sSelArg,sRetVal,iTdsGroupID,sGrpId,sTDSNarr,TDSFlag,sVouEntryno,iTDSEntryAmt,iTDSEntryPer
Dim iPayRecAmt,sFormula,sTdsRndOff

sSelArg = Request("voutype")
sFormVal = Request("hFormVal")


sVouCode=Request("hVouCode")
sVouName=Request("hVouName")

bActionFlag=Request.Form("hActionFlag")
iTdsGroupID = Request.Form("SelTDSGrp")
set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set ObjDom = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")


set Root=oDOM.documentElement

sVouStatus="010101" 'Crearted For Approval

sOrgId=Root.Attributes.GetNamedItem("UnitNo").value
sVouType=Root.Attributes.GetNamedItem("CRDR").value
sVoucDate=Root.Attributes.GetNamedItem("VouDate").value
sVouNo =Root.Attributes.GetNamedItem("VoucherNo").value
sAccHeadCode=Root.Attributes.GetNamedItem("BookAcchead").value
sApprove=Root.Attributes.GetNamedItem("Approver").value
iTransNo=Root.Attributes.GetNamedItem("TransNo").value
sBookNo =Root.Attributes.GetNamedItem("BookNo").value




sRetVal = GetVouchXML(iTransNo)
ObjDom.Load server.MapPath(sRetVal)
Set DelRoot = ObjDom.documentElement


' Response.Write "sOption="&sAction
'Response.end

IF CStr(sBkInsDate) = "" Then
	'sBkInsDate = FormatDateTime(Date())
	sBkInsDate = Day(Date())&"/"&Month(Date())&"/"& year(Date())
End IF
Response.write sBkInsDate &"========="
IF CStr(sBkInsType) = "" Then
	sBkInsType = "C"
End IF

IF CStr(sBkInsNo) = "" Then
	sBkInsNo = "0"
End IF

IF trim(sOption) = "Y" then
 	sTemp = split(sBkInsNo,"-")
	iEntNo		= sTemp(0)
	iInsEntNo	= stemp(1)
	sBkInsNo	= sTemp(2)
End IF


'Set tempnode = Root.childNodes(0)

'sPayTo = tempnode.Attributes.Item(2).nodevalue

FOR EACH EntryNode IN Root.childNodes
	IF EntryNode.nodeName = "Entry" Then
		sPayTo=EntryNode.Attributes.getNamedItem("Payto").value
	End IF
Next

con.BeginTrans

sQuery = "Select CreatedVouchStatus From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouStatus = objRs(0)
End IF
objRs.Close



'--------------Insert into Histroy tables---------------------------
sQuery="select isnull(max(AmendmentNo),0)+1 from Acc_T_HistoryVoucherHeader where "&_
		"TransactionNumber="&iTransNo
		'Response.Write sQuery
objRs.open sQuery,con
	iAmdendmentNo=objRs(0)
objRs.close()



sQuery="INSERT INTO Acc_T_HistoryVoucherHeader(TransactionNumber, AmendmentNo,OUDefinitionID, "&_
	"BookCode, BookNumber, TransactionType, AccountHead, PartyType, "&_
	"PartySubType, PartyCode, VoucherNumber, CreatedVoucherNo, CreatedTransNo, ReceiptNo,"&_
	"VoucherDate, VoucherAmount, CrDrIndication, CreatedBy, CreatedOn, ApprovedBy, ApprovedOn,"&_
	"AccountedBy, AccountedOn, AuditedBy, AuditedOn, BRSTransactionNo, BRSAmendmentNo, ClearedOn,"&_
	"BankInstrumentType, BankInstrumentNo, BankInstrumentDate, DrawnOnBank, PayableAt, "&_
	"HistoryType, HistoryBy, HistoryOn, HistoryReason)"&_

	"SELECT CreatedTransNo, "&iAmdendmentNo&", OUDefinitionID, "&_
	"BookCode, BookNumber, TransactionType, AccountHead, PartyType, "&_
	"PartySubType, PartyCode, CreatedVoucherNo,CreatedVoucherNo,CreatedTransNo, ReceiptNo, "&_
	"VoucherDate, VoucherAmount, CrDrIndication, CreatedBy, CreatedOn, ApprovedBy, ApprovedOn,"&_
	"NULL, NULL, NULL, NULL, NULL, NULL, NULL,"&_
	"NULL, NULL, NULL, NULL, NULL, "&_
	"'A',"&getUserid&",getdate(),'Amendment' "&_
	"FROM Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo

con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherDetails(TransactionNumber, AmendmentNo, VoucherEntryNumber,AccountingUnit, AccUnitAccountHead, "&_
		"AccUnitPartyType, AccUnitPartySubType,AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication) "&_
		"SELECT CreatedTransNo, "&iAmdendmentNo&",VoucherEntryNumber, AccountingUnit, AccUnitAccountHead, AccUnitPartyType, AccUnitPartySubType,"&_
		"AccUnitPartyCode, VoucherNarration, Amount, TransCrDrIndication FROM Acc_T_CreatedVoucherDetails	where "&_
		" CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherCC(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
		"AccountingUnit, AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount) "&_
		"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
		"AccUnitAccountHead, AccUnitCCHead, CCRatioPercent, CCRatioAmount from Acc_T_CreatedVoucherCCDet "&_
		"where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherAH(TransactionNumber, AmendmentNo, VoucherEntryNumber, "&_
		"AccountingUnit, AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount) "&_
		"select CreatedTransNo,"&iAmdendmentNo&", VoucherEntryNumber, AccountingUnit, "&_
		"AccUnitAccountHead, AccUnitAnalyticalCode, RatioPercentage, RatioAmount from Acc_T_CretedVoucherAHDet "&_
		"where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryPybleAdjustment(PayablesNumber, PaidByTransactionNo, AmendmentNo, "&_
		"PaidOn, AmountPaid)SELECT PayablesNumber, CreatedTransNo,"&iAmdendmentNo&", PaidOn, AmountPaid FROM Acc_T_CreatedPybleAdjDet "&_
		"where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryRcvblAdjustment(ReceivableNumber, RecdByTransactionNo, AmendmentNo, "&_
		"ReceivedOn, AmountReceived)SELECT ReceivableNumber, CreatedTransNo,"&iAmdendmentNo&", ReceivedOn, AmountReceived "&_
		"FROM Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="INSERT INTO Acc_T_HistoryVoucherAdvPymt(TransactionNumber, AmendmentNo, OUDefinitionID, "&_
	"PartyType, PartySubType, PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, "&_
	"AdvanceAdjusted) SELECT CreatedTransNo,"&iAmdendmentNo&", OUDefinitionID, "&_
	"PartyType, PartySubType, PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived,"&_
	"AdvanceAdjusted FROM Acc_T_CreatedAdvances where CreatedTransNo="&iTransNo
con.execute(sQuery)


'Selecting The Orginal Voucher No in the Vouch created using Insert between two Vouch

IF CStr(sVouCode) = "01" Then
	sQuery = "Select isNull(OtherApplnTableName,'') From Acc_T_CreatedVoucherHeader  "&_
			 "Where CreatedTransNo = "&iTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sRecpNo = objRs(0)
	End IF
	objRs.Close
End IF

'----------Delete Reocords-----------------------------------------
sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedPybleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedRcvbleAdjDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedAdvances where CreatedTransNo="&iTransNo
con.execute(sQuery)

sExp = "//PayRec/Doc[@AdjType=""R""]"
Set TempNode = DelRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		iPayNo = TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		dAdjAmt = TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value

		sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iPayNo&" "

		Response.Write sQuery &"<br>"
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			iCrAdvNo = objRs(0)
		Else
			iCrAdvNo = 0
		End IF
		objRs.Close

		sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted =  "&_
				 "isNull(AdvanceAdjusted,0) - "&dAdjAmt&" WHERE CreatedAdvanceNo = "&iCrAdvNo&" "

		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery

	Next
End IF




sExp = "//PayRec/Doc[@AdjType=""P""]"
Set TempNode = DelRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		iPayNo = TempNode.Item(iCtr).Attributes.getNamedItem("PayableNo").Value
		dAdjAmt = TempNode.Item(iCtr).Attributes.getNamedItem("AmtToAdjust").Value

		sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iPayNo&" "

		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			iCrAdvNo = objRs(0)
		Else
			iCrAdvNo = 0
		End IF
		objRs.Close

		sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted =  "&_
				 "isNull(AdvanceAdjusted,0) - "&dAdjAmt&" WHERE CreatedAdvanceNo = "&iCrAdvNo&" "

		Response.Write sQuery &"<br><br>"
		Con.Execute sQuery

	Next
End IF



sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)

dCRAmt = 0
dDRAmt = 0

FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		IF EntryNode.Attributes.Item(1).nodeValue="C" THEN
			dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		ELSE
			dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		END IF
	END IF
NEXT

dCRAmt = CDbl(dCRAmt)
dDRAmt = CDbl(dDRAmt)
'Response.Write dCRAmt & ">>>"&dDRAmt &sVouType&"<BR>"
IF StrComp(sVouType,"D")=0 THEN
	dTotal=dCRAmt-dDRAmt
	IF dTotal >0 THEN
		sTransType=sVouName&"R"
	ELSE
		dTotal=Abs(dTotal)
		sVouType="C"
		sTransType=sVouName&"P"
	END IF
ELSEIF StrComp(sVouType,"C")=0 THEN
	dTotal=dDRAmt-dCRAmt

	IF dTotal >0 THEN
		sTransType=sVouName&"P"
	ELSE
		dTotal=Abs(dTotal)
		sVouType="D"
		sTransType=sVouName&"R"
	END IF
ELSE
	dTotal=0
	sTransType="GJR"
END IF

iTdsAmount = 0
Dim sTempVouNo,sTemparr,sTemp2,iTdsEntryType
sExp = "//TDS"
Set tempNode = Root.selectNodes(sExp)
IF tempNode.length <> 0 Then
	For iCtr = 0 To tempNode.length - 1
		iTdsAmount = iTdsAmount + CDbl(tempNode.Item(iCtr).Attributes.getNamedItem("PayRecAmount").Value)
	Next
Else
	iTdsAmount = 0
End IF

dTotal = CDbl(dTotal) - CDbl(iTdsAmount)

IF CStr(sVouStatus) = "" Then
	sVouStatus = "010101" 'Voucher Created For Approval
End IF

IF Cstr(iTdsGroupID) = "" Then
	iTdsGroupID = 0
End IF
sApprove = Request.Form("selUserID")
Response.Write "dTotal="&dTotal&"<BR>"
IF sVouType="" or sVouCode="08" THEN
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,PayToRecdFrom,TDSGroupID) values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal &_
			 ",NULL," & sApprove & ",getdate(),NULL,'"& sVouStatus & "','"&sPayTo&"',"&iTdsGroupID&")"
ELSEIF CStr(sVouCode) = "01" and Len(sRecpNo) <> 0 Then
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,OtherApplnTableName,TDSGroupID) values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,"&sAccHeadCode&",'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			 ",'"&sPayTo&"','"&sVouType&"',"&sApprove&",getdate(),NULL,'"&sVouStatus&"', '"&sRecpNo&"',"&iTdsGroupID&" )"
Elseif Cstr(sVouCode) = "02" Then
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus, "&_
			 "BankInstrumentType,BankInstrumentNo,PayableAt,BankInstrumentDate,DrawnOnBank,TDSGroupID ) values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,"&sAccHeadCode&",'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			 ",'"&sPayTo&"','"&sVouType&"',"&sApprove&",getdate(),NULL,'"&sVouStatus&"','"&sBkInsType&"','"&sBkInsNo&"', "&_
			 "'"&sBkPayableat&"',Convert(datetime,'"&sBkInsDate&"',103),'"&sBkDrawnOn&"',"&iTdsGroupID&")"
Else
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID "&_
			 ") values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,"&sAccHeadCode&",'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			 ",'"&sPayTo&"','"&sVouType&"',"&sApprove&",getdate(),NULL,'"&sVouStatus&"',"&iTdsGroupID&")"
END IF

Response.Write sQuery& "<BR>"
'Response.End
con.execute(sQuery)
Response.Clear
'Newly Added by Maheswari on 16th June 2008
'Response.Write sAction
dim sOption,iEntNo,iInsEntNo,sAction,i,iBkInsAmt
IF  sVouCode="02" THEN
	sExp ="//BankInstrumentDet"
	Set tempNode = Root.Selectnodes(sExp)
	'	Response.Write tempNode.Length
	sQuery = "Delete from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo  = "&iTransNo&" "
	Response.Write sQuery&"<BR><BR>"
	con.execute(sQuery)
	iCtr = 1
	' Response.Write "Length = "& tempNode.length&"<BR><BR>"

	if tempNode.length > 0 then
		For i = 0 to tempNode.length -1
			Response.Write "<P>*****************"&i&"*****************<BR>"
			sBkInsType	 = tempNode.Item(i).Attributes.GetNamedItem("InsType").value
			sBkInsNo	 = tempNode.Item(i).Attributes.GetNamedItem("InsNo").value
			sBkInsDate	 = tempNode.Item(i).Attributes.GetNamedItem("InsDate").value
			sBkPayableat = tempNode.Item(i).Attributes.GetNamedItem("PayAt").value
			sBkDrawnOn	 = tempNode.Item(i).Attributes.GetNamedItem("DrawnOn").value
			iBkInsAmt	 = tempNode.Item(i).Attributes.GetNamedItem("InsAmt").value
			sOption      = tempNode.Item(i).Attributes.GetNamedItem("Option").value
			sAction		 = tempNode.Item(i).Attributes.GetNamedItem("Action").value

			IF trim(sOption) = "Y" then
				sTemp		= split(sBkInsNo,"-")
				iEntNo		= sTemp(0)
				iInsEntNo	= stemp(1)
				sBkInsNo	= sTemp(2)
			End IF
			Response.Write "Action ="& sAction &"<BR>"
			Response.Write "sOption ="& sOption &"<BR>"

			sQuery = "UPDATE Acc_T_CreatedVoucherHeader SET BankInstrumentType = '"&sBkInsType&"', BankInstrumentNo = '"&sBkInsNo&"',  " & _
					 "BankInstrumentDate = convert(datetime,'"&sBkInsDate&"',103), PayableAt = '"&sBkPayableat&"', DrawnOnBank = '"&sBkDrawnOn&"' " &_
					 "WHERE CreatedTransNo = "&iTransNo

			con.execute(sQuery)

			If trim(sOption) = "Y"  then

			 	'sQuery	= "Update Acc_T_CreatedVoucherHeader set BankInstrumentEntryNo="&iEntNo&",InstrumentEntryNo="&iInsEntNo&" where CreatedTransNo="&iTransNo
				'		   Response.Write sQuery &"<BR><BR>"
				'		   con.execute(sQuery)

				IF trim(sAction) = "C" then
					sQuery = "Update Acc_R_BankInstrumentUsage set Status = 'C' where  CreatedTransNo  = "&iTransNo&" "&_
							 " and EntryNo = "&iEntNo&"  and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo&"' "
					Response.Write sQuery&"<BR><BR>"
					con.execute(sQuery)
				ElseIF trim(sAction) = "R" then
					sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = 0,Status = 'N' where  CreatedTransNo  = "&iTransNo&" "&_
							" and EntryNo = "&iEntNo&" and InstrumentEntryNo = "&iInsEntNo&" and InstrumentNo = '"&sBkInsNo&"' "
					Response.Write sQuery&"<BR><BR>"
					con.execute(sQuery)
				End IF
				Response.Write "<B>Insertiion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
				sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
						 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,BankInstrumentEntryNo,InstrumentEntryNo1,InstrumentAmount)"&_
						 " Values ("&iTransNo&","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
						 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iEntNo&","&iInsEntNo&","&iBkInsAmt&" ) "
				Response.Write sQuery &"<BR><BR>"
				con.execute(sQuery)
			Else
				sQuery = "Update Acc_R_BankInstrumentUsage set CreatedTransNo  = "&iTransNo&",Status = 'U' where InstrumentNo = '"&sBkInsNo&"' "
				Response.Write sQuery&"<BR><BR>"
				con.execute(sQuery)

				Response.Write "<B>Insertiion in Acc_T_CreatedVoucherInstrumentDet</B>"&"<BR>"
				sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,"&_
						 "BankInstrumentNo,BankInstrumentDate,PayableAt,DrawnOnBank,InstrumentAmount)"&_
						 " Values ("&iTransNo&","&iCtr&",'"&sBkInsType&"','"&sBkInsNo&"',convert(datetime,'"&sBkInsDate&"',103),"&_
						 "'"&sBkPayableat&"','"&sBkDrawnOn&"',"&iBkInsAmt&" ) "
				Response.Write sQuery &"<BR><BR>"
				con.execute(sQuery)
			End IF

				iCtr = iCtr + 1
		Next
	end if 'if tempNode.length > 0 then
Else
	sBkInsType = "C"
	sBkInsNo = "0"
	'sBkInsDate = FormatDateTime(Date())
	sBkInsDate = Date()
	sBkPayableat = ""
	sBkDrawnOn = ""
	sOption = ""
END IF



'Response.End
'Response.Clear
'-----------------------PROCESS ENTRY NODES-------------------------------
Dim CheckNode

FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
		'sGrpID = EntryNode.Attributes.getNamedItem("GroupId").value
		sGrpID = "0"

		sExp = "//Entry[@No="&sEntryno&" and @TdsAmount]"

		Set CheckNode = Root.selectnodes(sExp)
		IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
			IF CheckNode.length <> 0 Then
				iTdsAmount = EntryNode.Attributes.getNamedItem("TdsAmount").value
				iTdsPer = EntryNode.Attributes.getNamedItem("TdsPercentage").value
			Else
				iTdsAmount = 0
				iTdsPer = 0
			End IF
		Else
			iTdsAmount = 0
			iTdsPer = 0
		End IF
		'Response.Write   "<BR>****"& sAmount &"-"& iTdsAmount & "***<BR>"
		'sAmount = Cdbl(sAmount) - Cdbl(iTdsAmount)
		'Response.Write " sAmount = "& sAmount &"<BR><BR>"
		sTDSNarr = "TDS Entry"
		TDSFlag = "Y"

	'---------PROCESS THE CHILD NODES OF ENTRY NODE FOR DETAIL TABLE UPDATION----
		sVouEntryno  = sEntryno + 1
		FOR EACH HeaderNode IN EntryNode.childNodes
			IF HeaderNode.nodeName="AccHead" THEN
					sAccCode=HeaderNode.Attributes.getNamedItem("No").value
					sAccType=HeaderNode.Attributes.getNamedItem("Type").value
			END IF 'End of Check for Account head Node

			IF HeaderNode.nodeName="TDS" THEN
		'		Response.Write HeaderNode.nodeName &"<BR>"
				iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
				iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
				iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
				iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
				sFormula	 = HeaderNode. getAttribute("Formula")
				sTdsRndOff   = HeaderNode. getAttribute("TdsRndOff")

				IF CStr(sEntryType) = "D" Then
					iTdsEntryType = "C"
				Else
					iTdsEntryType = "D"
				End IF

					'Response.Write "QQQ="&iTDSEntryAmt
				sQuery =" Insert into Acc_T_CreatedVoucherDetails(CreatedTransNo,AccountingUnit,VoucherEntryNumber,"&_
						" AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
						" VoucherNarration,Amount,TransCrDrIndication,TDSonAmount,PackCode,TDSFlag,TDSPercentage,TDSRoundOff) "&_
						" values("&iTransNo&",'"&sAccUnit&"',"&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL, "&_
						" '"&sTDSNarr&"',"&iPayRecAmt&",'"&iTdsEntryType&"',"&iTDSEntryAmt&","&sGrpId&",'"&TDSFlag&"',"&iTDSEntryPer&",'"&sTdsRndOff&"')"
						Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
						con.execute(sQuery)
				sVouEntryno = sVouEntryno + 1
			END IF 'IF HeaderNode.nodeName="TDS" THEN


			IF 	HeaderNode.nodeName="Narration" THEN
					sNarration=Replace(HeaderNode.text,"'","''")
			END IF 'End of Check for Narration Node
		NEXT

		IF Cstr(iTdsAmount) = "" Then
			iTdsAmount = 0
		End IF
	'-------------END OF PROCESSING CHILD NODES OF ENTRY NODE---------------------

	'----------------------------DETAIL TABLE UPDATION-------------------------
		Response.Write   "<BR>*******"& sAmount &"-"& iTdsPer & "*******<BR>"
		IF StrComp(sAccType,"G")=0 THEN
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSRoundOff) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&",'"&sTdsRndOff&"')"
		ELSE

			sTemp=Split(sAccCode,"?")
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication, TDSOnAmount,TDSRoundOff) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
			sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&iTdsAmount&",'"&sTdsRndOff&"')"
		END IF
		Response.Write "<BR>Details="& sQuery& "<BR><BR><BR>"
		con.execute(sQuery)
	'-----------------------END OF DETAIL TABLE UPDATION----------------------

	DIM sCCGroup,sAddCode,sAddGroupCode,sAddRatio,sAddAmount,dAddTotal
	dAddTotal=0
	'--------PROCESS CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION----
		FOR EACH HeaderNode IN EntryNode.childNodes
	'----------------------PROCESS PAYABLE / RECEVIABLE NODES ----------------
			IF 	HeaderNode.nodeName="PayRec" THEN
				FOR EACH  nodANL IN HeaderNode.childNodes
					sAddCode=nodANL.Attributes.getNamedItem("No").value
					sAddAmount=nodANL.Attributes.getNamedItem("AmtToAdjust").value
					sDocType=nodANL.Attributes.getNamedItem("DocType").value
					sAdjTy = nodANL.Attributes.getNamedItem("AdjType").value
					iAdvNo = nodANL.Attributes.getNamedItem("PayableNo").value

					'dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)

					'=====================================================
					IF CStr(Trim(sAdjTy)) = "D" Then
						sQuery = "Select DrCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&sAddCode&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							sAddCode = objRs(0)
						End IF
						objRs.Close
					Elseif CStr(Trim(sAdjTy)) = "C" Then
						sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&sAddCode&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							sAddCode = objRs(0)
						End IF
						objRs.Close
					End IF

					IF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "CR" Then
						'dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
						Select Case CStr(sAdjTy)
							Case "PI"
								dAddTotal = CDbl(dAddTotal) + CDbl(sAddAmount)
							Case "C"
								dAddTotal = CDbl(dAddTotal) + CDbl(sAddAmount)
							Case "I"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
							Case "D"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
							Case "P"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
						End Select

						'Response.Write dAddTotal &"<br><br>"

					ElseIF CStr(sAccType) = "P" and CStr(Left(sAccCode,2)) = "DR" Then
						Select Case CStr(sAdjTy)
							Case "I"
								dAddTotal = CDbl(dAddTotal) + CDbl(sAddAmount)
							Case "D"
								dAddTotal = CDbl(dAddTotal) + CDbl(sAddAmount)
							Case "PI"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
							Case "C"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
							Case "R"
								dAddTotal = CDbl(dAddTotal) - CDbl(sAddAmount)
						End Select
					Else
						dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
					End IF

					IF CStr(sAdjTy) <> "P" and CStr(sAdjTy) <> "R" Then
						IF CDbl(sAddAmount) > 0 THEN
							IF sDocType="C" THEN
							'-----------------------UPDATE PAYABLE TABLE------------------------------
								sQuery = "Select PayablesNumber From Acc_T_CreatedPayables Where PayablesNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close

								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)
							ELSEIF sDocType="D" THEN
							'-----------------------UPDATE RECEIVABLE TABLE---------------------------
								sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where ReceivableNumber = "&sAddCode&" "&_
										 "and PartyCode = "&sTemp(3)&" and PartyType = '"&sTemp(0)&"' and PartySubType = "&sTemp(1)&" "
								objRs.Open sQuery,con
								IF Not objRs.EOF Then
									sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
											 "ReceivedOn,AmountReceived,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								Else
									sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
											 "PaidOn,AmountPaid,VoucherEntryNumber) Values ("&sAddCode&","&iTransNo&","&_
											 "getdate(),"&sAddAmount&","&sEntryno&")"
								End IF
								objRs.Close

								Response.Write sQuery &"<br><br>"
								con.execute(sQuery)
							END IF
						END IF
					ElseIF CStr(sAdjTy) = "P" Then
						sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							iCrAdvNo = objRs(0)
						End IF
						objRs.Close

						sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&" WHERE  "&_
								 "CreatedAdvanceNo = "&iCrAdvNo&" AND PartyType = '"&sTemp(0)&"' AND PartySubType = "&sTemp(1)&"  "&_
								 "AND PartyCode = "&sTemp(3)&" "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

						sQuery = "INSERT INTO Acc_T_CreatedPybleAdjDet (PayablesNumber, CreatedTransNo, "&_
								 "PaidOn, AmountPaid, AdjustType,VoucherEntryNumber) "&_
								 "VALUES ("&iCrAdvNo&", "&iTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&sEntryno&") "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

					ElseIF CStr(sAdjTy) = "R" Then
						sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo&" "
						objRs.Open sQuery,Con
						IF Not objRs.EOF Then
							iCrAdvNo = objRs(0)
						End IF
						objRs.Close

						sQuery = "UPDATE Acc_T_CreatedAdvances SET AdvanceAdjusted = isNull(AdvanceAdjusted,0) + "&sAddAmount&" WHERE  "&_
								 "CreatedAdvanceNo = "&iCrAdvNo&" AND PartyType = '"&sTemp(0)&"' AND PartySubType = "&sTemp(1)&"  "&_
								 "AND PartyCode = "&sTemp(3)&" "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

						sQuery = "INSERT INTO Acc_T_CreatedRcvbleAdjDet (ReceivableNumber, CreatedTransNo, "&_
								 "ReceivedOn, AmountReceived, AdjustType,VoucherEntryNumber) "&_
								 "VALUES ("&iCrAdvNo&", "&iTransNo&", Convert(datetime,getDate(),103), "&sAddAmount&", 'A',"&sEntryno&") "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery

					End IF
				NEXT
			END IF
	'-------------END OF PROCESSING PAYABLE / RECEVIABLE NODES ---------------
		NEXT
	'-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-
	'-------------PROCESS FOR ADVANCE TABLE UPDATION -------------------------
		IF sAccType="P" AND CDbl(sAmount)> CDbl(dAddTotal) THEN
			sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
			objRs.Open sQuery,Con
			If Not objRs.EOF Then
				iAdvNo = objRs(0)
			End IF
			objRs.Close

			IF sVouType="C" THEN
				sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
					"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
					" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
					""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL)"
					Response.Write sQuery
					con.execute(sQuery)
			ELSEIF sVouType="D" THEN
			sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
					"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
					" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
					""&sTemp(3)&","&sAmount&",NULL,"&CDbl(sAmount)- CDbl(dAddTotal)&",NULL)"
					Response.Write sQuery
					con.execute(sQuery)
			END IF
		END IF
	'-------------END OF PROCESSING FOR ADVANCE TABLE UPDATION ---------------
	END IF
	'-------------END OF ENTRY NODE CHECK ---------------
NEXT
'------------------END OF PROCESSING ENTRY NODES--------------------------

'Response.clear
'Response.end
con.CommitTrans
'Con.RollbackTrans
'Response.End

'-----------------------PROCESS ENTRY NODES FOR CC/ANAL NODES-------------
FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value

		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sAccUnit
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C"

		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			nodANL.appendChild(tempNode.item(0))
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			nodANL.appendChild(tempNode.item(0))
			bAddFlag=true
		end if
		if bAddFlag then
		  Dim adoConn

		   Set adoConn = Server.CreateObject("ADODB.Connection")
		   adoConn.ConnectionString = con
		   adoConn.CursorLocation = 3
		   adoConn.Open

		   sQuery = "Proc_VouCCANALUpdate"


		   Dim adoCmd
		   Set adoCmd = Server.CreateObject("ADODB.Command")
		   Set adoCmd.ActiveConnection =adoConn
		   adoCmd.CommandText = sQuery
		   adoCmd.CommandType = 4 'adCmdStoredProc
		   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)

		   Dim adoRS
		   Set adoRS = adoCmd.Execute()

		end if
	end if
NEXT

'********************* Insertion for VoucherForApproval Table ****************************
Dim iRecCount

IF CStr(sVouStatus) = "010101" Then
	sQuery = "Select ApprovedBy From Acc_T_VoucherApprovalTracking Where CreatedTransNo = "&iTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iAppBy = objRs(0)
	Else
		iAppBy = ""
	End IF
	objRs.Close

	sQuery = "Select CreatedTransNo From Acc_T_VouchersForApproval Where CreatedTransNo = "&iTransNo&" "
	objRs.Open sQuery,Con
	IF objRs.EOF Then
		iRecCount = "0"
	End IF
	objRs.Close
End IF

IF CStr(iAppBy) = "" Then
	iAppBy = Session("userid")
End IF



sQuery = "INSERT INTO Acc_T_VouchersForApproval (CreatedTransNo, ApprovalLevel, ToBeApprovedBy) "&_
		 "VALUES ("&iTransNo&", 1, "&iAppBy&") "

Response.Write sQuery
Con.Execute sQuery

	'End IF
'End IF
'objRs.Close
'Response.Write "<BR>"&"OK"&"<BR>"
'Response.End
'********************* End of Insertion for VoucherForApproval Table *********************
'------------------END OF PROCESSING ENTRY NODES--------------------------

if con.Errors.count <>0 then
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
	if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")) then
	    objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml"))
	end if 'if objfs.FileExists(Server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")) then


	select case sVouName
	Case "CA" Response.Redirect ("VouCADisplay.asp?Flag="&bActionFlag&"&TransNo="&iTransNo&"&approver="&sApprove&"&hFormVal="&sFormVal&"&voutype="&sSelArg)
	Case "BA" Response.Redirect ("VouBADisplay.asp?Flag="&bActionFlag&"&TransNo="&iTransNo&"&approver="&sApprove&"&hFormVal="&sFormVal&"&voutype="&sSelArg)
	Case "GJ" Response.Redirect ("VouGJDisplay.asp?Flag="&bActionFlag&"&TransNo="&iTransNo&"&approver="&sApprove&"&hFormVal="&sFormVal&"&voutype="&sSelArg)
	end select
end if

%>
