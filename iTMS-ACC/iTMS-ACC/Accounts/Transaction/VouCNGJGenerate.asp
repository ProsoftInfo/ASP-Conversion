
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCNGJGenerate.asp
	'Module Name				:	Accounts (Purchase Requisition)
	'Author Name				:	MAHESHWARI S.
	'Created On					:	DEC 06, 2006
	'Modified By				:	Ragavendran
	'Modified On				:	Feb 06,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'Connects To				:	VouCNOthersEntry.asp
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:	refer
	'Boolean					:
	'Object Holders				:
	'Description				: 
%>


<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!--#include File="../../include/purItemCommon.asp" -->
<!--#include File="../../include/PurChkItemSpecPack.asp" -->
<!--#include File="../../include/NoSeries.asp" -->

<%
'Declaration
Dim oDOM,objFs,rsTemp,Objrs,Root,EntNode,AccHeadNode,sql
Dim sUnitNo,sUnitName,sBkNo,sBkName,VouCrDr,VouDate,sPartyCode,sApprover,sPartyName,sBkCode
Dim sTranType,sPartySubType,sRecNo,iTransNo,sVouAmt,sAudLock,sAmdExs,sBankInsDate,sVouStatus,sToBeExp,sChqPrn,sTDSBank,sTemp
Dim ENo,ECrDr,EPayTo,EAmt,EAccUnit,EAccName,ETDSAmt,ETDSElgi,ETDSPer,EPayRecAmt,sAccParType,sAccParSubType,sAccParCode,Arr,sTotAmt
Dim AccHNo,AccHName,AccAnaly,AccCostCen,AccType,AccTranFlag,sNarr,sVouNarr,iVouEntNo
Dim iItemCode,iClassCode,iInvQty,iInvRate,iBasAmt,iDisPer,iDisAmt,iTransRate,iTransBasAmt,iTranDisAmt,iTransVouAmt
Dim iRatePer,iTDSonAmt,iTDSPerc,sInvNo,dInvDate,sCallFrom ,SRQty
							
Dim iVouNo,iSeriesCode,iSeriesNo

set oDOM = server.CreateObject("Microsoft.XMLDOM")
set objFs = server.CreateObject("Scripting.FileSystemObject")

set Objrs = server.CreateObject("ADODB.Recordset")
set rsTemp = server.CreateObject("ADODB.Recordset")
sCallFrom = Request.QueryString("hCallFrom")
if trim(sCallFrom)="SR" then SRQty = Request.QueryString("hReturnQty")

'Response.Write Session.SessionID

if objFs.FileExists(server.MapPath("../temp/transaction/Voucher Entry_CNGJ_"&Session.SessionID&".xml")) then
	oDOM.Load server.MapPath("../temp/transaction/Voucher Entry_CNGJ_"&Session.SessionID&".xml")


	set Root = oDOM.documentElement
	
	con.BeginTrans
	
		If Root.NodeName = "voucher" then
			sUnitNo = Root.getAttribute("UnitNo")
			sUnitName = Root.getAttribute("UnitName")
			sBkNo = Root.getAttribute("BookNo")  
			sBkName = Root.getAttribute("BookName") 
			VouCrDr = Root.getAttribute("CRDR")
			VouDate = Root.getAttribute("VouDate")
			sPartyCode = Root.getAttribute("PartyCode") 
			sApprover = Root.getAttribute("Approver")
			sPartyName = Root.getAttribute("PartyName") 
			sInvNo = Root.getAttribute("InvNo")
			dInvDate = Root.getAttribute("InvDate")
		End If
			Arr = split(sPartyCode,"?")
			sAccParType = Arr(0)
			sAccParSubType = Arr(1)
			sAccParCode = Arr(3) 
			'Response.Write "sAccParCode="& sAccParCode &"<BR><BR>"
			
			sBkCode = "08"
			sTranType = "GJR"
			sPartySubType = 0
			sPartyCode = 0
			sRecNo = 0 
			sVouAmt = 0
			sAudLock = 0
			sAmdExs = 0
			sVouStatus = "010101"
			sToBeExp = 0
			sChqPrn = 0
			sTDSBank = 0
			sTemp = GetUserID()
	
			sql="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
			objRs.open sql,con
				iTransNo=objRs(0)
			objRs.Close
			'Response.Write "iTransNo="& iTransNo &"<BR><BR>"

			sql="select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
					"OUDefinitionID='"&sUnitNo&"' and BookCode='"&sBkCode&"' and BookNumber= "&sBkNo&" " 
				Response.Write sql&"<BR><BR>"
				
			objRs.open sql,con
				iSeriesNo=objRs(0)
				iSeriesCode=objRs(1)
			objRs.close()
			
			iVouNo=GenSeriesNumber(sUnitNo,iSeriesNo,iSeriesCode,VouDate)
			'Response.Write "<BR><BR>iVouNo="& iVouNo &"<BR><BR>"
			'Inserting into Header Table : Acc_T_CreatedVoucherHeader
		'"&VouDate&"'
			sql = "Insert Into Acc_T_CreatedVoucherHeader(CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,PartySubType,PartyCode,CreatedVoucherNo, "&_
				  "ReceiptNo,VoucherDate,VoucherAmount,CreatedBy,CreatedOn,AmendmentExists,CreatedVouchStatus,TDSBank,PartyType,CrDrIndication) "&_
				  " Values('" &iTransNo& "','"&sUnitNo&"','"&sBkCode&"','"&sBkNo&"','"&sTranType&"','"&sPartySubType&"','"&sPartyCode&"','"&iVouNo&"','"&sRecNo&"', "&_
				  " convert(datetime,'"&VouDate&"',103),'"&sVouAmt&"','"&sTemp&"',getdate(),'"&sAmdExs&"','"&sVouStatus&"','"&sTDSBank&"',NULL,NULL) " 
				  
				  Response.Write "CreatedVouHeader = " & sql &"<Br><Br>"
				con.execute sql
			If Root.hasChildNodes Then
			iVouEntNo = 1
			FOR EACH EntNode IN Root.childNodes
				sTotAmt = sTotAmt +CDbl(EntNode.getAttribute("Amount"))
			NEXT
			if trim(sCallFrom)="SI" then
				sVouNarr = "Credit Note for GJ "& sInvNo & " - "& dInvDate
			elseif trim(sCallFrom)="SR" then
				sVouNarr = "Credit Note for GJ "& sInvNo & " - "& dInvDate &" Sales Return Qty : "& SRQty 
			elseif trim(sCallFrom)="PI" then
				sVouNarr = "Credit Note for GJ "& sInvNo & " - "& dInvDate
			end if
							
			iItemCode = 0
			iClassCode = 0
			iInvQty = 0
			iInvRate = 0
			iBasAmt = 0
			iDisPer = 0
			iDisAmt = 0
			iTransRate = 0
			iTransBasAmt = 0
			iTranDisAmt = 0
			iTransVouAmt = 0
			iRatePer = 0
			iTDSonAmt = 0
			iTDSPerc = 0 
								
			sql = "Insert Into Acc_T_CreatedVoucherDetails(CreatedTransNo,VoucherEntryNumber,AccountingUnit,AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,VoucherNarration,"&_
					  "Amount,TransCrDrIndication,ItemCode,ClassificationCode,InvoicedQuantity,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,TransactRate,TransactBasicAmount, "&_
					  "TransactDiscAmount,TransactVouchAmount,RatePer,TDSonAmount,TDSPercentage) "&_
					  "Values('" &iTransNo& "','"&iVouEntNo&"','"&sUnitNo&"','"&sAccParType&"','"&sAccParSubType&"','"&sAccParCode&"','"&sVouNarr&"','"&sTotAmt&"','"&VouCrDr&"', "&_
					  " '"&iItemCode&"','"&iClassCode&"','"&iInvQty&"','"&iInvRate&"','"&iBasAmt&"','"&iDisPer&"','"&iDisAmt&"','"&iTransRate&"','"&iTransBasAmt&"','"&iTranDisAmt&"','"&iTransVouAmt&"', "&_
				     " '"&iRatePer&"','"&iTDSonAmt&"','"&iTDSPerc&"')"
				     
					Response.Write "Details table = "& sql &"<BR><BR>"
				con.execute sql
				
				
				For each EntNode in Root.Childnodes
				
					If EntNode.NodeName = "Entry" then
						ENo			= EntNode.getAttribute("No")
						ECrDr		= EntNode.getAttribute("CRDR")
						EAmt		= EntNode.getAttribute("Amount") 
						EAccUnit	= EntNode.getAttribute("AccUnit") 
						EAccName	= EntNode.getAttribute("AccName") 
						EPayTo		= EntNode.getAttribute("PayTo") 
						EPayRecAmt	= EntNode.getAttribute("PayRecAmount") 
						ETDSAmt		= EntNode.getAttribute("TDSAmount") 
						ETDSElgi	= EntNode.getAttribute("TDSElgi") 
						ETDSPer		= EntNode.getAttribute("TDSPercentage") 
						
											
						
						If EntNode.hasChildnodes Then
							For each AccHeadNode in EntNode.childnodes
								If AccHeadNode.NodeName = "Narration" Then
									sNarr = AccHeadNode.Text
								End If
								
								If AccHeadNode.NodeName = "AccHead" Then
									AccHNo		= AccHeadNode.getAttribute("No")
									AccCostCen	= AccHeadNode.getAttribute("CostCenter")
									AccAnaly	= AccHeadNode.getAttribute("Analytical")
									AccHName	= AccHeadNode.getAttribute("Name")
									AccType		= AccHeadNode.getAttribute("Type")
									AccTranFlag = AccHeadNode.getAttribute("TransFlag")
								
								iVouEntNo  = iVouEntNo + 1
							'Response.Write "One = "& iVouEntNo &"<BR><BR>"
							'Inserting into Acc_T_CreatedVoucherDetails Table
							sql = "Insert Into Acc_T_CreatedVoucherDetails(CreatedTransNo,VoucherEntryNumber,AccountingUnit,AccUnitAccountHead,VoucherNarration,"&_
								  "Amount,TransCrDrIndication,ItemCode,ClassificationCode,InvoicedQuantity,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,TransactRate,TransactBasicAmount, "&_
								  "TransactDiscAmount,TransactVouchAmount,RatePer,TDSonAmount,TDSPercentage) "&_
								  "Values('" &iTransNo& "','"&iVouEntNo&"','"&sUnitNo&"','"&AccHNo&"','"&sVouNarr&"','"&EAmt&"','"&ECrDr&"', "&_
								  " '"&iItemCode&"','"&iClassCode&"','"&iInvQty&"','"&iInvRate&"','"&iBasAmt&"','"&iDisPer&"','"&iDisAmt&"','"&iTransRate&"','"&iTransBasAmt&"','"&iTranDisAmt&"','"&iTransVouAmt&"', "&_
							     " '"&iRatePer&"','"&iTDSonAmt&"','"&iTDSPerc&"')"
							     
								Response.Write "Details table = "& sql &"<BR><BR>"
							con.execute sql
							End If						
								
							Next
						End If	
					End If	
						
						'Response.Write "One = "& iVouEntNo &"<BR><BR>"
							
				next
			end IF
	
		'Response.End
				
	If con.Errors.count <> 0 Then
		con.RollbackTrans
		For iCounter=0 to con.Errors.count - 1
			Response.Write con.Errors(iCounter) &"<br>"
		Next
	Else
	'	con.RollbackTrans
	'	Response.End 
				
		'Response.Clear
		con.CommitTrans
	End if		
'	if objFS.fileExists(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml")) then
'		objFS.deletefile(Server.MapPath("../temp/transaction/Voucher Entry_CN_"&Session.SessionID&".xml"))
'	End if	
	'Response.Redirect("VOUNEWCNBOOKSEL.ASP")			
	Response.Redirect("CREDITVOUCHERS.ASP")
End if 'if objFs.FileExists
%>