<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNCommAmdGenrate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	February 28,2003
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
<!--#include file="../../include/NoSeries.asp"-->
<%

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp,VouRoot,sCrDrindi
dim sOrgId,sBookNo,sVouType,iVouNo,iCreatedVouNo,sPurInvNo,sPurInvDt,sPurTranNo
dim sVouCode,sApprove,sVoucDate,sPayto
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo,iCreatedTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode,iCrSeriesNo,iCrSeriesCode
Dim sStr,TempNode,sStr2,iCtr,iCounter
dim sParType,sParSubType,sParCode,Checknode,Objfs
Dim sTdsAmt,sTdsPer,sApproval,sApprover,iOldInvNos,iAmdendmentNo,sTemparr,iAppUpdate

Set objfs = CreateObject("Scripting.FileSystemObject")
set objRs  = server.CreateObject("adodb.recordset")

iOldInvNos = Request.Form("hOldInvCode")

sTemparr = Split(iOldInvNos,",")

'Response.Write iOldInvNos

'Response.End




' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_CNComm_"&Session.SessionID&".xml")	
set VouRoot=oDOM.documentElement



sStr = "//voucher"
Set Root = VouRoot.selectNodes(sStr)

sOrgId=Root.Item(0).Attributes.Item(0).nodeValue
sBookNo =Root.Item(0).Attributes.Item(2).nodeValue
sVoucDate=Root.Item(0).Attributes.Item(4).nodeValue
sPurTranNo=Root.Item(0).Attributes.Item(6).nodeValue
sPurInvNo=Root.Item(0).Attributes.Item(7).nodeValue
sPurInvDt=Root.Item(0).Attributes.Item(8).nodeValue
iTransNo = Root.Item(0).Attributes.Item(9).nodeValue
sTransType="CNR"
sVouType="C"
sVouCode="07"

sStr2 = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr2)
'dTotal=TempNode.Item(0).Attributes.Item(2).nodeValue
For iCounter = 0 To TempNode.length - 1
	IF CStr(TempNode.Item(iCounter).Attributes.getNamedItem("No").Value) = "1" Then
		sPayto = TempNode.Item(0).Attributes.Item(1).nodeValue
	End IF
Next

dTotal = 0
dTotal = CDbl(dTotal)
sStr2 = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr2)
For iCtr = 0 To TempNode.length - 1
	IF CStr(TempNode.Item(0).Attributes.Item(3).nodeValue) = "C" Then
		dTotal = abs(dTotal + Cdbl(TempNode.Item(iCtr).Attributes.Item(2).nodeValue))
	Else
		dTotal = abs(dTotal - Cdbl(TempNode.Item(iCtr).Attributes.Item(2).nodeValue))
	End IF
	
	Response.Write dTotal &"<br><br>"
	
Next

sStr2 = "//Party"
Set TempNode = VouRoot.selectNodes(sStr2)
sParType=Tempnode.Item(0).Attributes.Item(0).nodeValue
sParSubType=Tempnode.Item(0).Attributes.Item(1).nodeValue
sParCode=Tempnode.Item(0).Attributes.Item(2).nodeValue

sStr2 = "//Entry"
Set EntryNode = VouRoot.selectNodes(sStr2)

sQuery = "Select CreatedVouchStatus,CreatedVoucherNo From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouStatus = objRs(0)
	iCreatedVouNo = objRs(1)
End IF
objRs.Close

con.BeginTrans

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

sQuery = "INSERT INTO Acc_T_HisCreatedPayables (AmendmentNo, PayablesNumber, CreatedTransNo, "&_
		 "OUDefinitionID, PartyType, PartySubType, PartyCode, VoucherDate, PartyBillNumber, "&_
		 "PartyBillDate, AmountPayable, AmountPaid, Narration) SELECT "&iAmdendmentNo&", PayablesNumber, "&_
		 "CreatedTransNo, OUDefinitionID, PartyType, PartySubType, PartyCode, "&_
		 "VoucherDate, PartyBillNumber, PartyBillDate,AmountPayable, AmountPaid, Narration "&_
		 "FROM Acc_T_CreatedPayables where CreatedTransNo="&iTransNo 
con.execute(sQuery)

'----------Delete Reocords-----------------------------------------
sQuery="delete Acc_T_CretedVoucherAHDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherCCDet where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery = "DELETE FROM Acc_T_CreatedPayables WHERE CreatedTransNo = "&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherDetails where CreatedTransNo="&iTransNo
con.execute(sQuery)

sQuery="delete Acc_T_CreatedVoucherHeader where CreatedTransNo="&iTransNo
con.execute(sQuery)

Response.Write iOldInvNos &"<br><br><br>"
Response.Write iOldInvNos

For iCounter = 0 To UBound(sTemparr)
	sQuery = "update Sal_T_AdditionalAgents set CommissionToPay=1 where AccTransactionNo="&Trim(sTemparr(iCounter))&" and AgentCode="&sParCode
	Response.Write sQuery &"<br>"
	con.execute(sQuery)		
Next
'----------Delete Reocords-----------------------------------------

iCreatedTransNo = iTransNo

sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
		 "PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
		 "CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,PAYTORECDFROM) values"&_
		 "("&iCreatedTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
		 "'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&iCreatedVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal &_
		 ",'"&sVouType&"','SC',"&getUserid&",getdate(),NULL,'"&sVouStatus&"','"&sPayto&"')"
	
con.execute(sQuery)

'************************************ End of Header Details Entry **********************************************************
sStr = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr)
IF TempNode.length <> 0 Then
	For iCtr = 0 To TempNode.length - 1
		sEntryno = TempNode.Item(iCtr).Attributes.Item(0).nodeValue
		sAmount = TempNode.Item(iCtr).Attributes.Item(2).nodeValue 
		sCrDrindi = TempNode.Item(iCtr).Attributes.Item(3).nodeValue
		sTdsAmt = TempNode.Item(iCtr).Attributes.Item(4).nodeValue 
		sTdsPer = TempNode.Item(iCtr).Attributes.Item(6).nodeValue 
		
		for each HeaderNode in TempNode.Item(iCtr).childNodes
			if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue 
				sAccType=HeaderNode.Attributes.Item(4).nodeValue 
			end if 'End of Check for Account head Node
			if 	HeaderNode.nodeName="Narration" then
				sNarration=HeaderNode.text
				sNarration = Mid(sNarration,1,Len(sNarration)-1)
			end if 'End of Check for Narration Node
		next 'End of Entry Node Loop
		
		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
		sQuery=sQuery& iCreatedTransNo&",'"&sOrgId&"'" 
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sCrDrindi&"', "&sTdsAmt&", "&sTdsPer&")"
			
		con.execute(sQuery)
		
'************************************************ End of Details Entry Over ***************************
		dim iPayableNo
		sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
	
		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
				" PartyBillDate, AmountPayable, AmountPaid)values("&iPayableNo&",NULL,'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
					
		'Response.Write sQuery& "<BR>"	
		con.execute(sQuery)
		
		'Response.Write " ============= " &"<br>"

		sStr = "//voucher"
		Set Checknode = VouRoot.selectNodes(sStr)
		'Response.Write Checknode.length
		IF Root.length <> 0 Then
			IF Root.length = 1 Then
				sPurTranNo = Root.Item(0).Attributes.Item(6).nodeValue
				sQuery="update Sal_T_AdditionalAgents set CommissionToPay=0	where AccTransactionNo="&sPurTranNo&" and AgentCode="&sParCode
				Response.Write sQuery &"<br>"
				con.execute(sQuery)
			Else
				For iCounter = 0 To Root.length - 1
					sPurTranNo = Root.Item(iCtr).Attributes.Item(6).nodeValue
					sQuery="update Sal_T_AdditionalAgents set CommissionToPay=0	where AccTransactionNo="&sPurTranNo&" and AgentCode="&sParCode
					Response.Write sQuery &"<br>"
					con.execute(sQuery)
				Next
			End IF
		End IF
	Next
	
Dim iAppby,iRecCount
IF CStr(sVouStatus) = "010101" Then
	sQuery = "Select ApprovedBy From Acc_T_VoucherApprovalTracking Where CreatedTransNo = "&iCreatedTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		iAppBy = objRs(0)
	Else
		iAppBy = ""
	End IF
	objRs.Close
	
	sQuery = "Select CreatedTransNo From Acc_T_VouchersForApproval Where CreatedTransNo = "&iCreatedTransNo&" "
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
		 "VALUES ("&iCreatedTransNo&", 1, "&iAppBy&") "
		 
Response.Write sQuery
Con.Execute sQuery
	

con.CommitTrans

'******************************** Cost and ANal Procedure Call ************************

'dim 'sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode,
Dim sExp,bAddFlag

sExp = "//Entry"
Set HeaderNode = VouRoot.selectNodes(sExp)
IF HeaderNode.length <> 0 Then
	for iCounter = 0 To HeaderNode.length - 1
		sEntryno=HeaderNode.Item(iCounter).Attributes.getNamedItem("No").value
		
		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=VouRoot.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value
				
		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iCreatedTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sorgid
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C" 
		
		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=VouRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=VouRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
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
	next 'End of Entry Node Loop
End IF


'******************************** Cost and ANal Procedure Call ************************
if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
	'con.CommitTrans
	
	
	For iCounter = 0 To Root.length-1
		Set newElem  = oDOM.createAttribute("CrTransNo")
		newElem.value = iCreatedTransNo
		Root.Item(iCounter).setAttributeNode(newElem)
	
		Set newElem  = oDOM.createAttribute("CrVoucherNo")
		newElem.value = iCreatedVouNo
		Root.Item(iCounter).setAttributeNode(newElem)
	
		Set newElem  = oDOM.createAttribute("TransNo")
		newElem.value = iTransNo
		Root.Item(iCounter).setAttributeNode(newElem)
	
		Set newElem  = oDOM.createAttribute("VoucherNo")
		newElem.value = iVouNo
		Root.Item(iCounter).setAttributeNode(newElem)
	Next
	
	oDOM.Save server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")	
	objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher AMD_CNComm_"&Session.SessionID&".xml"))
	Response.Redirect ("VouCNCommisionDisplay.asp?TransNo="&iCreatedTransNo)
end if
End IF

%>

