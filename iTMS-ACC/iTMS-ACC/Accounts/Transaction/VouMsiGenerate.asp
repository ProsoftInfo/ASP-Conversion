<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouMsiGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 30, 2004
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<%

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery,sExp,iMsiTransNo
dim EntryNode,HeaderNode,nodANL,newElem,tempNode,bAddFlag

dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,sVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,sPayTo,iTdsAmount,iTdsPer
dim iSeriesNo,iSeriesCode,iAdvNo,sApprover

dim sAccHeadCode,sCallFrom
Dim iAccHeadNo,iTDSEntryAmt,iTDSEntryPer,iPayRecAmt,sFormula,TDSFlag,sGrpId
Dim iTdsEntryType,sVouEntryno,sTDSNarr,iTdsGroupID

Dim sChequeNo,sChequeDate,sPayable,sDrawnOn

sVouCode=Request("hVouCode")
sVouName=Request("hVouName")
sApprove=Request("optApprove")
sApprover = Request("selUserId")
iMsiTransNo = Request.Form("hTransNo")

iTdsGroupID = Request.Form("SelTDSGrp")
sCallFrom = Request("CallFrom")
bAddFlag=false

IF CStr(sApprove) = "Y" Then
	sVouStatus="010101" 'Crearted For Approval
Else
	sVouStatus="010103" 'Crearted For Accounting 
End IF

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_"&sVouName&"_"&Session.SessionID&".xml")	

set Root=oDOM.documentElement

sOrgId=Root.Attributes.getNamedItem("UnitNo").value
sBookNo =Root.Attributes.getNamedItem("BookNo").value
sVouType=Root.Attributes.getNamedItem("CRDR").value
sVoucDate=Root.Attributes.getNamedItem("VouDate").value
sAccHeadCode=Root.Attributes.getNamedItem("BookAcchead").value


Root.Attributes.getNamedItem("Approver").value=sApprove
if UCase(sCallFrom)=UCase("Bank") then
    sChequeNo = Root.getAttribute("InstNo")
    sChequeDate = Root.getAttribute("InstDate")
    sPayable = Root.getAttribute("PayAt")
    sDrawnOn = Root.getAttribute("DrawnOn")
end if 

sPayTo=Replace(Root.childNodes(0).Attributes.getNamedItem("Payto").value,"'","''")

FOR EACH EntryNode IN Root.childNodes
	IF EntryNode.Attributes.getNamedItem("CRDR").value="C" THEN
		dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
	ELSE
		dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.getNamedItem("Amount").value)
	END IF	
NEXT	

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

IF strcomp(sVouType,"D")=0 THEN
	sQuery="select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
ELSE
	sQuery="select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
END IF		

Response.Write sQuery

objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
end if
objRs.close()

con.BeginTrans

if trim(iSeriesNo) ="" or IsNull(iSeriesNo) then    
    Response.clear
    Dim sBookName
    sBookName = GetAccBookName(sVouCode,sBookNo)
    Response.write "<h1>Number Series is not created for <font color=red>"& sBookName &"</font></h1>" 
    Response.end
end if

sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate) 

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iTransNo=objRs(0)
objRs.Close

Response.Write iTransNo

Response.Write sVouType
'Response.End 


if trim(iTdsGroupID)="" or IsNull(iTdsGroupID) then iTdsGroupID="NULL"

IF sVouType="" or sVouCode="08" THEN 
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID) values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,NULL,'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal &_
			 ",NULL," & getUserid & ",getdate(),NULL,'"& sVouStatus & "',"&iTdsGroupID&")"
ELSE
	sQuery = "insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			 "PartyType,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"&_
			 "PayToRecdFrom,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,TDSGroupID) values"&_
			 "("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			 "NULL,"&sAccHeadCode&",'"&sVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			 ",'"&sPayTo&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"',"&iTdsGroupID&")"

END IF

Response.Write "<p>"&sQuery& "<BR><BR>"	
con.execute(sQuery)



'-----------------------PROCESS ENTRY NODES-------------------------------
FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
		'sGrpId=EntryNode.Attributes.getNamedItem("GroupId").value
		
		IF CStr(sEntryType) = "D" Then
			iTdsEntryType = "C"
		Else
			iTdsEntryType = "D"
		End IF
		
		IF CStr(sVouCode) = "01" or CStr(sVouCode) = "02" or CStr(sVouCode) = "08" Then
			iTdsAmount = EntryNode.Attributes.getNamedItem("TdsAmount").value
			iTdsPer = EntryNode.Attributes.getNamedItem("TdsPercentage").value
		Else
			iTdsAmount = 0
			iTdsPer = 0
		End IF
		IF Cstr(iTdsAmount) = "" Then
			iTdsAmount = 0
			iTdsPer = 0
		End if
	
		sTDSNarr = "TDS Entry"
		TDSFlag = "Y"
		
	
		sVouEntryno = sEntryno + 1
	'---------PROCESS THE CHILD NODES OF ENTRY NODE FOR DETAIL TABLE UPDATION----
		FOR EACH HeaderNode IN EntryNode.childNodes
		'Response.Write "HeaderNode="& HeaderNode.nodeName &"<BR>"
			IF HeaderNode.nodeName="AccHead" THEN
					sAccCode=HeaderNode.Attributes.getNamedItem("No").value
					sAccType=HeaderNode.Attributes.getNamedItem("Type").value
			END IF 'End of Check for Account head Node
			
			'Added by Maheshwari on 30 th May 2008 for TDS Entries
			
			IF HeaderNode.nodeName="TDS" THEN
				
				
				iAccHeadNo	 = HeaderNode.getAttribute("AccHeadCode")
				iTDSEntryAmt = HeaderNode.getAttribute("TDSAmount")
				iTDSEntryPer = HeaderNode.getAttribute("TdsPercentage")
				iPayRecAmt	 = HeaderNode.getAttribute("PayRecAmount")
				sFormula	 = HeaderNode. getAttribute("Formula")
			'	Response.Write "iTDSEntryPer="&iTDSEntryPer &"<BR>"
				sQuery =" Insert into Acc_T_CreatedVoucherDetails(CreatedTransNo,AccountingUnit,VoucherEntryNumber,"&_
						" AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"&_
						" VoucherNarration,Amount,TransCrDrIndication,TDSonAmount,TDSFlag,TDSPercentage) "&_
						" values("&iTransNo&",'"&sAccUnit&"',"&sVouEntryno&","&iAccHeadNo&",NULL,NULL,NULL, "&_
						" '"&sTDSNarr&"',"&iPayRecAmt&",'"&iTdsEntryType&"',"&iTDSEntryAmt&",'"&TDSFlag&"',"&iTDSEntryPer&")"
						Response.Write "<BR>TDS="& sQuery & "<BR><BR>"
						con.execute(sQuery)
				sVouEntryno = sVouEntryno + 1
			END IF 'IF HeaderNode.nodeName="TDS" THEN
			IF 	HeaderNode.nodeName="Narration" THEN
					sNarration=Replace(HeaderNode.text,"'","''")
			END IF 'End of Check for Narration Node
		NEXT 
	'-------------END OF PROCESSING CHILD NODES OF ENTRY NODE---------------------	
	'----------------------------DETAIL TABLE UPDATION-------------------------
		IF StrComp(sAccType,"G")=0 THEN
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'" 
			sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&","&iTdsPer&")"
		ELSE
			sTemp=Split(sAccCode,"?")						
			sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
			sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,"
			sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication, TDSonAmount, TDSPercentage) values ("
			sQuery=sQuery& iTransNo&",'"&sAccUnit&"'" 
			sQuery=sQuery& ","&sEntryno&",NULL,'"&sTemp(0)&"',"&sTemp(1)&","&sTemp(3)&","
			sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&iTdsAmount&", "&iTdsPer&")"
		END IF
		Response.Write sQuery& "<BR><BR>"	
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
					dAddTotal=CDbl(sAddAmount)+CDbl(dAddTotal)
					IF CDbl(sAddAmount) > 0 THEN
						IF sDocType="C" THEN
						'-----------------------UPDATE PAYABLE TABLE------------------------------								
							sQuery="insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
								"PaidOn,AmountPaid) Values ("&sAddCode&","&iTransNo&","&_
								"getdate(),"&sAddAmount&")"
							con.execute(sQuery)						
							Response.Write "<p>"&sQuery	
						ELSEIF sDocType="D" THEN
						'-----------------------UPDATE RECEIVABLE TABLE---------------------------
							sQuery="insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
								"ReceivedOn,AmountReceived) Values ("&sAddCode&","&iTransNo&","&_
								"getdate(),"&sAddAmount&")"
							con.execute(sQuery)
							Response.Write "<p>"&sQuery	
						END IF
					END IF		
				NEXT
			END IF 
	'-------------END OF PROCESSING PAYABLE / RECEVIABLE NODES ---------------	
		NEXT 
		
		Response.Write "<p>sAccType = "& sAccType 
		
	'-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-
	'-------------PROCESS FOR ADVANCE TABLE UPDATION -------------------------	
		IF sAccType="P" AND CDbl(sAmount)> CDbl(dAddTotal) THEN
			IF sVouType="C" THEN
				sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
				objRs.Open sQuery,Con
				If Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
				
				sQuery="INSERT INTO Acc_T_CreatedAdvances(CreatedAdvanceNo,CreatedTransNo, OUDefinitionID, PartyType, PartySubType, "&_
					"PartyCode, ActualVoucherAmount, AdvancePaid, AdvanceReceived, AdvanceAdjusted)"&_
					" VALUES("&iAdvNo&", "& iTransNo&",'"&sAccUnit&"','"&sTemp(0)&"',"&sTemp(1)&","&_
					""&sTemp(3)&","&sAmount&","&CDbl(sAmount)- CDbl(dAddTotal)&",NULL,NULL)"
					Response.Write sQuery	
					con.execute(sQuery)						
			ELSEIF sVouType="D" THEN
				sQuery = "Select isNull(Max(CreatedAdvanceNo),0)+1 from Acc_T_CreatedAdvances "
				objRs.Open sQuery,Con
				If Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
				
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

if UCase(Trim(sCallFrom))=UCase("Bank") then
    sQuery = "Insert Into Acc_T_CreatedVoucherInstrumentDet (CreatedTransNo,InstrumentEntryNo,BankInstrumentType,BankInstrumentNo,BankInstrumentDate,InstrumentAmount,PayableAt,DrawnOnBank)"
    sQuery= sQuery &"  values("& iTransNo &",1,'Cheque',"& pack(sChequeNo) &",Convert(datetime,'"& sChequeDate &"',103),"& sAmount &","& pack(sPayable)&","& pack(sDrawnOn) &")"
    Response.Write "<p>"&sQuery
    con.execute sQuery
end if
	
'----------------------- Changing the Msi Voucher Status Details ----------------------
sQuery = "UPDATE Acc_T_MiscPymtRequestHeader SET CreatedVouchStatus = '010104', "&_
		 "ReceiptNo = "&iTransNo&" Where MiscTransNo = "&iMsiTransNo&" "
con.execute(sQuery)
Response.Write "<p>"&sQuery
'----------------------- Changing the Msi Voucher Status Details ----------------------

IF CStr(sApprove) = "Y" Then
	sQuery = "insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
			 "Values("&iTransNo&",1,"&sApprover&")"
	con.execute(sQuery)
	Response.Write "<p>"&sQuery
End IF


'con.RollbackTrans 
'Response.End

Response.Clear 
con.CommitTrans

'-----------------------PROCESS ENTRY NODES FOR CC/ANAL NODES-------------
FOR EACH EntryNode IN Root.childNodes
	IF  EntryNode.nodeName="Entry" THEN
		sEntryno=EntryNode.Attributes.getNamedItem("No").value
		sEntryType=EntryNode.Attributes.getNamedItem("CRDR").value
		sAmount=EntryNode.Attributes.getNamedItem("Amount").value
		sAccUnit=EntryNode.Attributes.getNamedItem("AccUnit").value
		
		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=Root.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value
			
		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sAccUnit
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C" 

		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			set HeaderNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(HeaderNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=Root.selectNodes(sExp)
		if tempNode.length >0 then
			set HeaderNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(HeaderNode)
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
'------------------END OF PROCESSING ENTRY NODES--------------------------

if con.Errors.count <>0 then
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else	
	
	'Set newElem  = oDOM.createAttribute("TransNo")
	'newElem.value = iTransNo
	'Root.setAttributeNode(newElem)
	
	'Set newElem  = oDOM.createAttribute("VoucherNo")
	'newElem.value = sVouNo
	'Root.setAttributeNode(newElem)
	
	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")	
	
	Response.Redirect ("VouMsiDisplay.asp?TransNo="&iTransNo)
	
end if

%>

