<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNOthCommGen.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
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
dim sParType,sParSubType,sParCode,Checknode
Dim sTdsAmt,sTdsPer,sApproval,sApprover

sApproval = Request.Form("optApprove")
sApprover = Request.Form("selUserId")

IF CStr(sApproval) = "Y" Then
	sVouStatus = "010101"
Else
	sVouStatus = "010103"
End IF



set objRs  = server.CreateObject("adodb.recordset")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")	
set VouRoot=oDOM.documentElement
sStr = "//voucher"
Set Root = VouRoot.selectNodes(sStr)

sOrgId=Root.Item(0).Attributes.Item(0).nodeValue
sBookNo =Root.Item(0).Attributes.Item(2).nodeValue
sVoucDate=Root.Item(0).Attributes.Item(4).nodeValue
sPurTranNo=Root.Item(0).Attributes.Item(6).nodeValue
sPurInvNo=Root.Item(0).Attributes.Item(7).nodeValue
sPurInvDt=Root.Item(0).Attributes.Item(8).nodeValue
sTransType="CNR"
sVouType="C"
sVouCode="07"

sStr2 = "//Entry"
Set TempNode = VouRoot.selectNodes(sStr2)
'dTotal=TempNode.Item(0).Attributes.Item(2).nodeValue
sPayto = TempNode.Item(0).Attributes.Item(1).nodeValue

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

sQuery = "select CreatedDrSeriesNo,CreatedDrSeriesCode,DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
		 "OUDefinitionID='"&sOrgId&"' and BookCode='07' and BookNumber= "&sBookNo


objRs.open sQuery,con
	iCrSeriesNo=objRs(0)
	iCrSeriesCode=objRs(1)
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
objRs.close()

con.BeginTrans

iCreatedVouNo=GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sVoucDate)
iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCreatedTransNo=objRs(0)
objRs.Close

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
			
		Response.Write "<br><br>"
		Response.Write sQuery
		con.execute(sQuery)
		
'************************************************ End of Details Entry Over ***************************
		dim iPayableNo,iCurrCode
		iCurrCode = 1
		sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close
	
		
		'sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
		'		"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
		'		" PartyBillDate, AmountPayable, AmountPaid)values("&iPayableNo&",NULL,'"&sOrgId&"',"&_
		'		"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
		'		"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
		
		'Changed For Testing
		sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
				"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
				" PartyBillDate, AmountPayable, AmountPaid)values("&iPayableNo&","&iCreatedTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
				"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0)"
		
		con.execute(sQuery)
		
		sStr = "//voucher"
		Set Checknode = VouRoot.selectNodes(sStr)
		Response.Write Checknode.length
		IF Root.length <> 0 Then
			IF Root.length = 1 Then
				sPurTranNo = Root.Item(0).Attributes.Item(6).nodeValue
				sQuery = "Select isNull(AccTransactionNo,0) From Sal_T_AdditionalAgents Where AccTransactionNo = "&sPurTranNo&" and AgentCode="&sParCode
				objRs.Open sQuery,Con
				IF Not objRs.EOF Then
					sQuery="update Sal_T_AdditionalAgents set CommissionToPay=0	where AccTransactionNo="&sPurTranNo&" and AgentCode="&sParCode
				Else
					sQuery = "INSERT INTO Sal_T_AdditionalAgents (AgentCode, CommissionType, AgentCommission, "&_
							 "CurrencyCode, AccTransactionNo, CommissionToPay) "&_
							 "VALUES ("&sParCode&", 'A', "&dTotal&", "&iCurrCode&", "&iCreatedTransNo&", '0') "
				End IF
				objRs.Close
				Response.Write "<br><br>"
				Response.Write sQuery
				con.execute(sQuery)
			Else
				For iCounter = 0 To Root.length - 1
					sPurTranNo = Root.Item(iCtr).Attributes.Item(6).nodeValue
					sQuery = "Select isNull(AccTransactionNo,0) From Sal_T_AdditionalAgents Where AccTransactionNo = "&sPurTranNo&" and AgentCode="&sParCode
					objRs.Open sQuery,Con
					IF Not objRs.EOF Then
						sQuery="update Sal_T_AdditionalAgents set CommissionToPay=0	where AccTransactionNo="&sPurTranNo&" and AgentCode="&sParCode
					Else
						sQuery = "INSERT INTO Sal_T_AdditionalAgents (AgentCode, CommissionType, AgentCommission, "&_
								 "CurrencyCode, AccTransactionNo, CommissionToPay) "&_
								 "VALUES ("&sParCode&", 'A', "&dTotal&", "&iCurrCode&", "&iCreatedTransNo&", '0') "
					End IF
					con.execute(sQuery)
				Next
			End IF
		End IF
	Next
	
IF CStr(sApproval) = "Y" Then
	sQuery = "insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
			 "Values("&iCreatedTransNo&",1,"&sApprover&")"
	con.execute(sQuery)
End IF
	

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
	Response.Redirect ("VouCNCommisionDisplay.asp?TransNo="&iCreatedTransNo)
end if
End IF

%>

