<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLCreate_Edit_Update.asp
	'Module Name				:	ACCOUNTS (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 25,2010
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
	Dim oDOM,Objrs,Objfs,Root,TempNode,sExp,iAccHead,sQuery,sHisno,sAccdesc,sAccShort
	Dim iCount,iCtr,sHisRes,sTempHis,iRecCount,sLogid
	Dim sIut,sLeg,sCost,sAnal,sSummary,sContra,sTds,sMemo,sCashTrans
	Dim sAppcode,sAppname,sUnitid,UnitNode,iOpenBal,sCrDrind
	Dim sAnalCode,sAnalGHCode,iCounter,iFirstTime
	Dim sCCCode,sCCGhCode
	Dim sBookCode,sBookNo,sExp2,ItemNode,Objrs2,sAction
	Dim sPartTy,sSubParTy,sOpenMonYear,sCloseMonYear,sNewCloseAmt,sSelBookCd,sTempVar
	Dim sPreCDInd,sPreCloseamt,sPreOpenamt,sDiffamt,sNewCDInd,sPreOpCdInd
	
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs2 = Server.CreateObject("ADODB.RecordSet")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	
	sLogid = Session("Userid")
	sAction = Request.QueryString("Acion")
	Response.Write "Acttion = "& sAction
	if trim(sAction)="Create" then
	
		sQuery = "Select isNull(Max(AccountHead),0)+1 from Acc_M_GLAccountHead "
		Objrs.Open sQuery,con
		if not objrs.EOF then
			iAccHead = Objrs(0)
		end if 
		Objrs.Close 
	end if
	

	oDOM.Load server.MapPath("../temp/master/GLAccount_Head_"&Session.SessionID&".xml")
	Set Root = oDOM.documentElement
	
	sExp = "//OpeningMonthYear"
	Set TempNode = Root.selectNodes(sExp)
	sOpenMonYear = TempNode.item(0).text
	
	sExp = "//ClosingMonthYear"
	Set TempNode = Root.selectNodes(sExp)
	sCloseMonYear = TempNode.item(0).text
	
	
	sExp = "//GroupCode"
	Set TempNode = Root.selectNodes(sExp)
	sAnalGHCode = TempNode.item(0).text
	
	if trim(sAction)="Edit" then
		sExp = "//AccHeadNo"
		Set TempNode = Root.selectNodes(sExp)
		iAccHead = TempNode.item(0).text
	end if 'if trim(sAction)="Edit" then
	
	sExp = "//Description"
	Set TempNode = Root.selectNodes(sExp)
	sAccdesc = TempNode.item(0).text
	
	sExp = "//ShortName"
	Set TempNode = Root.selectNodes(sExp)
	sAccShort = TempNode.item(0).text
	
	
	Con.BeginTrans
	sQuery = "Select isNull(Max(HistoryNo),0) + 1 From Acc_M_HistoryGLAccountHead "
	Objrs.Open sQuery,Con
	IF Not Objrs.EOF Then
		sHisno = Objrs(0)
	End IF
	Objrs.Close
	

	if sAction = "Edit" then	
		sQuery = "Insert into Acc_M_HistoryGLAccountHead Select *,"&sHisno&" From Acc_M_GLAccountHead "&_
				 "Where AccountHead = "&iAccHead&" "
		Response.Write "<p> "& sQuery		 
		Con.Execute sQuery
	
		sQuery = "Update Acc_R_OrgGLAccountHead Set AmendmentExists = '1' Where AccountHead = "&iAccHead&" "
		Response.Write "<p> "& sQuery
		Con.Execute sQuery
	
		sQuery = "Delete From Acc_M_GLSummaryApp Where AccountHead = "&iAccHead&" "
		Response.Write "<p> "& sQuery
		Con.Execute sQuery
	end if 'if sAction = "Edit" then
	
	iFirstTime = 1
	sExp2 = "//Units/UN"
	Set ItemNode = Root.selectNodes(sExp2)
	For iCount = 0 To ItemNode.length - 1

		sIut = ""
		sSummary= ""
		sLeg = ""
		sCost = ""
		sAnal = ""
		sContra = ""
		sTds = ""
		sMemo = ""
		sCashTrans = ""
		sSelBookCd = ""
		sSelBookCd = ""
		iOpenBal = "0"
		sCrDrind = "C"

		
		sUnitid = Trim(ItemNode.Item(iCount).Attributes.Item(0).Value)
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/IUT" 
		Set TempNode = Root.selectNodes(sExp)
		sIut = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/SummaryPosting"
		Set TempNode = Root.selectNodes(sExp)
		sSummary = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/SubLedger"
		Set TempNode = Root.selectNodes(sExp)
		sLeg = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/CostCenter"
		Set TempNode = Root.selectNodes(sExp)
		sCost = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/Analytical"
		Set TempNode = Root.selectNodes(sExp)
		sAnal = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/Contra"
		Set TempNode = Root.selectNodes(sExp)
		sContra = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/TDS"
		Set TempNode = Root.selectNodes(sExp)
		sTds = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/Memorandum"
		Set TempNode = Root.selectNodes(sExp)
		sMemo = TempNode.Item(0).Attributes.Item(0).Value
	
		sExp = "//Units/UN[@Code="&sUnitid&"]/CashTrans"
		Set TempNode = Root.selectNodes(sExp)
		sCashTrans = TempNode.Item(0).Attributes.Item(0).Value
		
		sExp = "//Units/UN[@Code="&sUnitid&"]/SummaryPostBook"
		Set TempNode = Root.selectNodes(sExp)
		sSelBookCd = TempNode.Item(0).Attributes.Item(0).Value
		sSelBookCd = Replace(sSelBookCd,",",":")
		
		sExp = "//Units/UN[@Code="&sUnitid&"]"
		Set TempNode = Root.selectNodes(sExp)
		iOpenBal = Trim(TempNode.Item(0).Attributes.Item(2).Value)
		sCrDrind = Trim(TempNode.Item(0).Attributes.Item(3).Value)
		
		
		IF CStr(sSummary) = "" Then
			sSummary = "1"
		End IF
		
		IF CStr(sIut) = "" Then
			sIut = "1"
		End IF
		
		IF CStr(sCost) = "" Then
			sCost = "0"
		End IF
		
		IF CStr(sAnal) = "" Then
			sAnal = "0"
		End IF
		
		IF CStr(sContra) = "" Then
			sContra = "0"
		End IF
		
		IF CStr(sMemo) = "" Then
			sMemo = "0"
		End IF
		
		IF CStr(sTds) = "" Then
			sTds = "0"
		End IF
		
		IF CStr(sLeg) = "" Then
			sLeg = "0"
		End IF
		
		
		'Response.Write "sCashTrans = "& sCashTrans 
		
		IF CStr(sCashTrans) = "" Then
			sCashTrans = "W"
		End IF
		
		if iFirstTime = 1 then
		
			iFirstTime = 2
			if trim(sAction)="Create" then
				sQuery =	"Insert into ACC_M_GLAccountHead (AccountHead,AccountHeadCode,AccountDescription,"&_
							" InterUnitTransact,CostCenterExists,AnalyticalHeadExists,EligibleForContras, "&_
							" MemorandumAccount,AllowTransactions,SubLedger,EligibleForTDS,SummaryPosting)"&_
							" Values("&iAccHead&",'"&sAccShort&"','"&sAccdesc&"','"&sIut&"', '"&sCost&"',"&_
							"'"&sAnal&"', '"&sContra&"','"&sMemo&"', '"&sCashTrans&"','"&sLeg&"',"&_
							"'"&sTds&"', '"&sSummary&"')"
			else
				sQuery = "UPDATE Acc_M_GLAccountHead SET AccountHeadCode = '"&sAccShort&"', AccountDescription = '"&sAccdesc&"', "&_
						 "InterUnitTransact = '"&sIut&"', CostCenterExists = '"&sCost&"', AnalyticalHeadExists = '"&sAnal&"', "&_
						 "EligibleForContras = '"&sContra&"', MemorandumAccount = '"&sMemo&"', AllowTransactions = '"&sCashTrans&"',  "&_
						 "SubLedger = '"&sLeg&"', EligibleForTDS = '"&sTds&"', SummaryPosting = '"&sSummary&"' "&_
						 "WHERE AccountHead = "&iAccHead&" "
			end if
			
			Response.Write "<p> "& sQuery
			Con.Execute sQuery
			
		end if 'if iFirstTime = 1 then
		
		IF sSelBookCd <> "" Then
			sTempVar = Split(sSelBookCd,":")
			For iCtr = 0 To UBound(sTempVar)
				sQuery = "INSERT INTO Acc_M_GLSummaryApp (AccountHead, OUDefinitionID, BookCode) "&_
						 "VALUES ("&iAccHead&", '"&sUnitid&"', '"&sTempVar(iCtr)&"') "
						 
				Response.Write sQuery &"<br>"
				
				Con.Execute sQuery
			Next
		End IF		
		
		
		
		
		sQuery = "Insert into Acc_R_HistoryGLAccApplications Select *,"&sHisno&" "&_
				 "From Acc_R_GLAccApplications Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&sUnitid&"' "
		Response.Write "<p> "& sQuery
		Con.Execute sQuery
		
		sQuery = "Select Count(1) From Acc_R_OrgGLAccountHead Where OUDefinitionID = '"&sUnitid&"' and  AccountHead = "&iAccHead&" "
		Response.Write "<p> "& sQuery
		Objrs.Open sQuery,con
		IF Not Objrs.EOF Then
			iRecCount = Objrs(0)
		End IF
		Objrs.Close
		
		sQuery = "Select AccountsGroupCode From Acc_R_OrgGLAccountHead Where AccountHead = "&iAccHead&" "
		Response.Write "<p> "& sQuery
		Objrs.Open sQuery,con
		IF Not Objrs.EOF Then
			sAnalGHCode = Objrs(0)
		End IF
		Objrs.Close
		
		Response.Write "sAnalGHCode = "& sAnalGHCode 
		
		IF CStr(iRecCount) <> "0" Then
			sQuery = "UPDATE Acc_R_OrgGLAccountHead SET InterUnitTransact = '"&sIut&"', CostCenterExists = '"&sCost&"',  "&_
					 "AnalyticalHeadExists = '"&sAnal&"', EligibleForContras = '"&sContra&"', MemorandumAccount = '"&sMemo&"',  "&_
					 "AllowTransactions = '"&sCashTrans&"', SubLedger = '"&sLeg&"', EligibleForTDS = '"&sTds&"', SummaryPosting = '"&sSummary&"', "&_
					 "OpeningBalance = "&iOpenBal&", OpeningCDIndication = '"&sCrDrind&"', AmendmentExists = '0' "&_
					 "WHERE OUDefinitionID = '"&sUnitid&"' AND AccountHead = "&iAccHead&"  "
		Else
			sQuery = "INSERT INTO Acc_R_OrgGLAccountHead (OUDefinitionID, AccountsGroupCode, AccountHead, LeafNode, "&_
					 "TransferClosing, InterUnitTransact, CostCenterExists, AnalyticalHeadExists, "&_
					 "EligibleForContras, MemorandumAccount, AllowTransactions, SubLedger, "&_
					 "EligibleForTDS, SummaryPosting, OpeningBalance, OpeningCDIndication, "&_
					 "CreatedBy, CreatedOn, ApprovedBy, ApprovedOn) "&_
					 "VALUES ('"&sUnitid&"', '"&sAnalGHCode&"', "&iAccHead&", '0', '0', '0', '"&sCost&"', '"&sAnal&"', "&_
					 "'"&sContra&"', '"&sMemo&"', '"&sCashTrans&"', '"&sLeg&"', '"&sTds&"', "&_
					 "'"&sSummary&"', "&iOpenBal&", '"&sCrDrind&"', "&sLogid&", getDate(), "&sLogid&", getDate()) "
		End IF
	
		
		Response.Write "<p> "& sQuery
							 
		Con.Execute sQuery
		
		iRecCount = 0
		
		'=============== Insertion of values in Opening and Closing Amount Values =============
		
		sQuery = "Select Count(1) From Acc_T_GLAccOpeningAmt Where AccountHead = "&iAccHead&" "&_
				 "and OUDefinitionID = '"&sUnitid&"' and OpeningMonthYear = '"&sOpenMonYear&"'" 
		Objrs.Open sQuery,con
		IF Not Objrs.EOF Then
			iRecCount = Objrs(0)
		End IF
		Objrs.Close
		Response.Write "iRecCount = "& iRecCount
		
		IF CStr(iRecCount) = "0" Then
			sQuery = "INSERT INTO Acc_T_GLAccOpeningAmt (OUDefinitionID, AccountHead, OpeningMonthYear, "&_
					 "OpeningAmount, OpeningCDIndication, ClosingMonthYear, "&_
					 "ClosingAmount, ClosingCDIndication) "&_
					 "VALUES ('"&sUnitid&"', "&iAccHead&", '"&sOpenMonYear&"', "&iOpenBal&", "&_
					 "'"&sCrDrind&"', '"&sCloseMonYear&"', "&iOpenBal&", '"&sCrDrind&"') "
			Response.Write "<p> "& sQuery
					 
			Con.Execute sQuery
		Else
			sQuery = "Select OpeningAmount,ClosingAmount,ClosingCDIndication,OpeningCDIndication From "&_
					 "Acc_T_GLAccOpeningAmt Where OUDefinitionID = '"&sUnitid&"' and "&_
					 "AccountHead = "&iAccHead&" " &_
					 "and OpeningMonthYear = '"&sOpenMonYear&"'" 
			Response.Write "<p> "& sQuery
			
		 
			Objrs.Open sQuery,Con
			IF Not Objrs.EOF Then
				sPreOpenamt = Objrs(0)
				sPreCloseamt = Objrs(1)
				sPreCDInd = Objrs(2)
				sPreOpCdInd = Objrs(3)
			End IF
			Objrs.Close
			
			sPreOpenamt = CDbl(sPreOpenamt)
			iOpenBal = CDbl(iOpenBal)
			
			IF Cdbl(sPreCloseamt) = 0 Then
			'/// IF The Closing amount is 0 Then Insert the Value for both Opening and Closong 
			'/// with the same value and CRDR Indication.
			
				sNewCloseAmt = iOpenBal
				sNewCDInd = sCrDrind
			Else	
				'/// IF The Opening CDIndication is C Then Reduce Val from Openingamt - ChangedAmt
				'/// IF D Then Reduce the value from Changedamt - Openingamy.
				IF CStr(sPreOpCdInd) = "C" Then
					sDiffamt = CDbl(sPreOpenamt - iOpenBal)
				Else
					sDiffamt = CDbl(iOpenBal - sPreOpenamt)
				End IF
			
				'/////// IF Closing CD Indication is D Then Closeamt + Diff amt +/-
				'////// IF Closing CD Indication is C Then -Closeamt + Diff amt +/- 
				IF CStr(sPreCDInd) = "D" Then
					sNewCloseAmt = CDbl(sPreCloseamt) + sDiffamt
					IF CDbl(sNewCloseAmt) > 0 Then
						sNewCDInd = "D"
					Else
						sNewCDInd = "C"
					End IF
				Else
					sPreCloseamt = CDbl(sPreCloseamt)
					sNewCloseAmt = -sPreCloseamt + sDiffamt
					IF CDbl(sNewCloseAmt) > 0 Then
						sNewCDInd = "D"
					Else
						sNewCDInd = "C"
					End IF
				End IF
			End IF
			sQuery = "UPDATE Acc_T_GLAccOpeningAmt SET OpeningAmount = "&iOpenBal&", OpeningCDIndication = '"&sCrDrind&"', "&_
					 "ClosingAmount = "&abs(sNewCloseAmt)&", ClosingCDIndication = '"&sNewCDInd&"' "&_
					 "WHERE OUDefinitionID = '"&sUnitid&"' AND AccountHead = "&iAccHead&" and OpeningMonthYear = '"&sOpenMonYear&"'" 
					 
			Response.Write "<p> "& sQuery
			
			Con.Execute sQuery
			
		End IF
		
		
			 
		'=============== Insertion of values in Opening and Closing Amount Values Ends Here ====
		
		if sAction="Edit" then
			sQuery = "Delete From Acc_R_OrgPartyType Where AccountHead = "&iAccHead&" and OUDEFINITIONID = '"&sUnitid&"' "
		
			Response.Write "<p> "& sQuery
		
			Con.Execute sQuery
		end if 'if sAction="Edit" then
		
		sExp = "//Units/UN[@Code="&sUnitid&"]/ParType"
		Response.Write "<p>"& sExp &"<p>"
		Set TempNode = Root.selectNodes(sExp)
		Response.Write "TempNode = "& TempNode.length
		IF TempNode.length <> 0 Then
			For iCounter = 0 To TempNode.length - 1
				sPartTy = TempNode.Item(iCounter).Attributes.Item(1).value
				sSubParTy = TempNode.Item(iCounter).Attributes.Item(2).value
				
					sQuery = "INSERT INTO Acc_R_OrgPartyType (PartyType, PartySubType, OUDefinitionID, AccountHead) "&_
							 "VALUES ('"&sPartTy&"', "&sSubParTy&", '"&sUnitid&"', "&iAccHead&") "
							 
					Response.Write "<p> "& sQuery
					Con.Execute sQuery
			Next
		End IF
		Response.Write ".............End"			
	Next
	
	'For Application ========================================================================
	
	sExp = "//Applications/APP"
	Set TempNode = Root.selectNodes(sExp)
	sTempHis = CInt(sHisno)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sAppcode = TempNode.Item(iCtr).Attributes.Item(0).Value
			sExp = "//Units/UN"
			Set UnitNode = Root.selectNodes(sExp)
			IF UnitNode.length <> 0 Then
				For iCount = 0 To UnitNode.length - 1
					sUnitid = Trim(UnitNode.Item(iCount).Attributes.Item(0).Value)
					iOpenBal = Trim(UnitNode.Item(iCount).Attributes.Item(2).Value)
					sCrDrind = Trim(UnitNode.Item(iCount).Attributes.Item(3).Value)
					
					sQuery = "Insert into Acc_R_HistoryOrgGLAcctHead Select *,"&sTempHis&",'A',"&sLogid&",getDate(),'"&sHisRes&"' "&_
							 "From Acc_R_OrgGLAccountHead Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&sUnitid&"' "
							 
					Response.Write "<p> "& sQuery
							 
					Next
			End IF
		Next
	End IF
	
	
	if sAction="Edit" then
			sQuery = "DELETE FROM Acc_R_GLAccApplications "&_
					 "WHERE AccountHead = "&iAccHead&" "
			Response.Write "<p> "& sQuery
			Con.Execute sQuery
	end if 'if sAction="Edit" then
					
	sExp = "//Applications/APP"
	Set TempNode = Root.selectNodes(sExp)
	sTempHis = CInt(sHisno)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sAppcode = TempNode.Item(iCtr).Attributes.Item(0).Value
			sExp = "//Units/UN"
			Set UnitNode = Root.selectNodes(sExp)
			IF UnitNode.length <> 0 Then
				For iCount = 0 To UnitNode.length - 1
					sUnitid = Trim(UnitNode.Item(iCount).Attributes.Item(0).Value)
					iOpenBal = Trim(UnitNode.Item(iCount).Attributes.Item(2).Value)
					sCrDrind = Trim(UnitNode.Item(iCount).Attributes.Item(3).Value)
		
					
					sQuery = "INSERT INTO Acc_R_GLAccApplications (OUDefinitionID, AccountHead, AvailableInAppln) "&_
							 "VALUES ('"&sUnitid&"', "&iAccHead&", "&sAppcode&") "
							 
					Con.Execute sQuery
					
					sTempHis = cint(sTempHis + 1)
				
				Next
			End IF
		Next
	End IF
	
	'For Application Over ========================================================================
	
	'For Analytical ===============================================================================
	sExp = "//Analytical/AN"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sUnitid = TempNode.Item(iCtr).Attributes.Item(0).Value
			sQuery = "Insert into Acc_R_HistoryOrgGLAH Select "&sHisno&",* "&_
					 "From Acc_R_OrgGLAnalytical Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&sUnitid&"' "
							 
			Con.Execute sQuery
		Next
	End IF

	if sAction="Edit" then		
		sQuery = "DELETE FROM Acc_R_OrgGLAnalytical "&_
			     "Where AccountHead = "&iAccHead&" "
				     
		Response.Write "<p> "& sQuery						 
		Con.Execute sQuery
	end if 'if sAction="Edit" then
	
	sExp = "//Analytical/AN"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sUnitid = TempNode.Item(iCtr).Attributes.Item(0).Value
			sAnalCode = TempNode.Item(iCtr).Attributes.Item(1).Value
			
			Response.Write "sAnalCode = "& sAnalCode 
			sAnalGHCode = TempNode.Item(iCtr).Attributes.Item(2).Value
			sQuery = "INSERT INTO Acc_R_OrgGLAnalytical(OUDefinitionID, AccountHead, AnalyticalCode,  "&_
					 "AHGroupCode, AllocationRatio) "&_
					 "VALUES ('"&sUnitid&"', "&iAccHead&", "&sAnalCode&", '"&sAnalGHCode&"', 0) "
							 
			Response.Write "<p> "& sQuery						 
			
			Con.Execute sQuery
		Next
	End IF
		
	
	'For Analytical Over ===========================================================================================
	
	'For Cost Center **************************************************************************************************
	
	
	sExp = "//CostCenter/CC"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		
		For iCtr = 0 To TempNode.length - 1
			sQuery = "Insert into Acc_R_HistoryOrgGLCC Select "&sHisno&",* "&_
					 "From Acc_R_OrgGLCostCentre Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&sUnitid&"' "
							 
			Response.Write "<p> "& sQuery						 
			
			Con.Execute sQuery
		Next
	End IF
		
	
	if sAction="Edit" then
		sQuery = "DELETE FROM Acc_R_OrgGLCostCentre "&_
			     "Where AccountHead = "&iAccHead&" "
		
		Response.Write "<p> "& sQuery						 
		Con.Execute sQuery
	end if '	if sAction="Edit" then
	
'	Response.Write " ********************* "
	
	sExp = "//CostCenter/CC"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sUnitid = Trim(TempNode.Item(iCtr).Attributes.Item(0).Value)
			sCCGhCode = Trim(TempNode.Item(iCtr).Attributes.Item(1).Value)
			sCCCode = Trim(TempNode.Item(iCtr).Attributes.Item(2).Value)
			sQuery = "INSERT INTO Acc_R_OrgGLCostCentre (OUDefinitionID, AccountHead,   "&_
					 "CostCenterHead, AllocationRatio) "&_
					 "VALUES ('"&sUnitid&"', "&iAccHead&", "&sCCCode&", 0) "
						
			Response.Write "<p> "& sQuery						 
			
			Con.Execute sQuery
			
		Next
	End IF
		
	
	'For Cost Center **********************************************************************************************************
	
	'For Frequent Books ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
	
	sExp = "//Books/BK/OrgBook"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		For iCtr = 0 To TempNode.length - 1
			sUnitid = TempNode.Item(iCtr).Attributes.Item(0).Value
			sQuery = "Insert into Acc_R_HistoryGLAccFreqUsed Select *,"&sHisno&" "&_
					 "From Acc_R_GLAccFrequentlyUsed Where AccountHead = "&iAccHead&" and OUDefinitionID = '"&sUnitid&"' "
					 
			Response.Write "<p> "& sQuery						 
							 
			Con.Execute sQuery
		Next
	End IF
	
	if sAction="Edit" then
			sQuery = "DELETE FROM Acc_R_GLAccFrequentlyUsed "&_
					 "Where AccountHead = "&iAccHead&" "
					 
			Response.Write "<p> "& sQuery
			Con.Execute sQuery
	end if'if sAction="Edit" then
	sExp = "//Books/BK"
	Set UnitNode = Root.selectNodes(sExp)
	
	IF UnitNode.length <> 0 Then	
		For iCount = 0 To UnitNode.length - 1
			sBookCode = UnitNode.Item(iCount).Attributes.Item(0).Value
			sExp = "//Books/BK[@Code="""&sBookCode&"""]/OrgBook"
			Set TempNode = Root.selectNodes(sExp)
			IF TempNode.length <> 0 Then
				For iCtr = 0 To TempNode.length - 1
					sUnitid = TempNode.Item(iCtr).Attributes.Item(0).Value
					sBookNo = TempNode.Item(iCtr).Attributes.Item(1).Value
					
					sQuery = "INSERT INTO Acc_R_GLAccFrequentlyUsed (OUDefinitionID, BookCode,   "&_
							 "BookNumber, AccountHead) "&_
							 "VALUES ('"&sUnitid&"', '"&sBookCode&"', "&sBookNo&", "&iAccHead&") "
							 
					Response.Write "<p> "& sQuery
					Con.Execute sQuery
				Next
			End IF
		Next
	End IF
		
	'For Frequent Books Over ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'	Con.RollBackTrans
'	Response.End 

	Response.Clear 	
	Con.CommitTrans

	Response.Redirect "GLACCHEADGRID.ASP"
%>

