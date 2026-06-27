<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>


<%
	'Program Name				:	GetCBGVoucherxml.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	Maheshwari S.
	'Created On					:	Oct 18, 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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



	'Set rs1 = Server.CreateObject("ADODB.Recordset")

	'iVNo = Request("Para")
	'Response.Write iVNo

	'With rs1
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source =  "Select BookCode From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iVNo & " "
	'	.ActiveConnection = con
	'	.Open
	'End With
	'If not rs1.EOF then
	'	sBkCode = rs1(0)
	'End If
	'rs1.Close
	'Response.Write "BookCode = "& sBkCode & "<BR><BR>"
	'GetXml(sBkCode)

	'GetXmlForCBG 59,"S"


Function GetXmlForCBG(iVNo,sDispTy)
	Dim sTemp,iVType,rs1,sBkCode
	Dim Root,objhttp,rs2,rs3,rs4,sQry,Elem1,oDOM,NewElem,fsobj,iCount,i,EntNode,Node
	Dim sUnitNo,sUnitName,sBkNo,sBkName,sCrDr,sVDate,sBkAccNo,sTransNo,sVouNo,sCTNo,sCVNo
	Dim iENo,iEtCrDr,iPayToRecdFrom,iEAmt,iEAccUnit,sAccName,iETdsAmt,iETdsPer,Elem5
	Dim sAccHeadNo,sAccUnit,sAccPartyType,sAccPartySubType,sAccPartyCode,sPartyName,sSubTypeName
	Dim sTempNo,Elem2,Elem3,Elem4,sAccHeadName,sCostCnt,sAnlCnt,rsCount,sNarr,sPassCode
	Dim sPayNo,sPInvNo,sPInvDate,sPTransAmt,sPAmdToAdj,sPCTransNo,sPDocType,sPayableNo,sPBookCode,sPAdjType
	Dim sRecNo,sRInvNo,sRInvDate,sRTransAmt,sRAmdToAdj,sRCTransNo,sRDocType,sReceivableNo,sRBookCode,sRAdjType
	Dim sDNo,sDInvNo,sDInvDate,sDTransAmt,sDAmdToAdj,sDCTransNo,sDDocType,sDPayableNo
	Dim iDNo,iDInvNo,iDInvDate,iDTransAmt,iDAmdToAdj,iDCTransNo,iDDocType,iDPayableNo
	Dim sBInsType,sBInsNo,sBInsDate,sBPay,sBDrwOn,BankElem,sRetVal,sCode,rsAcc
	Dim iTdsGrpID,iGrpHeadID,iTdsCtr,iAcHeadCode,iTDSPer,iFormula,iTDSOnAmt,iTDSAmt,iTemp,iTds,iCounter,iVouEntNo
	Dim sETDSFlag,TDSElem,saTemp,saTemp1,sCompMode,TDSNode ,sAmount,sTransType,iNewTdsPer,SCBQuery

	Set rs1 = server.CreateObject("ADODB.RecordSet")
	Set rs2 = server.CreateObject("ADODB.RecordSet")
	set rs3 = server.CreateObject("ADODB.RecordSet")
	set rs4 = server.CreateObject("ADODB.RecordSet")
	set rsAcc = server.CreateObject("ADODB.RecordSet")
	'Response.Write iVNo
	iTdsCtr = 0
	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source =  "Select BookCode,isNull(TDSGroupID,0) TDSGroupID,TransactionType From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iVNo & " "
		.ActiveConnection = con
		.Open
	End With
	If not rs1.EOF then
		sCode = rs1(0)
		iTds = rs1("TDSGroupID")
		sTransType = rs1(2)
	End If
	rs1.Close

	'IF CStr(iTds) = "" Then
	'	iTds = 0
	'End IF

	sPassCode = sCode
	'Response.Write "pass="&sPassCode
	'creation of Voucher node
	sQry = "Select A.OUDefinitionID,D.OrgUnitShortDescription,isNull(A.BookNumber,0),isNull(A.CrDrIndication,''),Convert(Char,A.VoucherDate,103), "&_
		   " A.CreatedTransno,A.CreatedVoucherNo from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A where "&_
		   " A.CreatedTransNo = " & iVNo & " and D.OUDefinitionID = A.OUDefinitionID "
	'Response.Write sQry & "<BR><BR>"

	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQry
		.ActiveConnection = con
		.Open
	End With

	IF not rs1.Eof then
		sUnitNo = rs1(0)
		sUnitName = rs1(1)
		sBkNo = rs1(2)
		sCrDr = rs1(3)
		sVDate = rs1(4)
		sCTNo = rs1(5)
		sCVNo = rs1(6)
	End IF
	rs1.Close
	'Response.Write "sVDate =" & sVDate & "<BR><BR>"
	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source =  " Select BookName,isNull(BookAccountHead,0) from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
				   " Where V. OUDefinitionID = " & sUnitNo & " and V.BookCode = " & sPassCode & " and A.CreatedTransNo = " & iVNo & " "&_
				   " and V.BookNumber = " & sBkNo &" "
		.ActiveConnection = con
		.Open
	End With
	'Response.Write "qry1= "& sQry & "<BR><BR>"
	IF not rs1.EOF then
		sBkName = rs1(0)
		sBkAccNo = rs1(1)
	End IF
	rs1.Close

	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source ="Select TransactionNumber,VoucherNumber from Acc_T_VoucherHeader where CreatedTransNo = " & iVNo & " "
		.ActiveConnection = con
		.Open
	End With
'	Response.Write sQry & "<BR><BR>"

	IF not rs1.EOF then
		sTransNo = rs1(0)
		sVouNo = rs1(1)
	End IF
	rs1.close

	'------------------------------------

	set oDOM=Server.CreateObject("Microsoft.XMLDOM")
	set fsobj=Server.CreateObject("Scripting.FileSystemObject")
	set Root=oDOM.DocumentElement


	set Root=oDOM.CreateElement("voucher")
	Root.setAttribute "UnitNo", sUnitNo
	Root.setAttribute "UnitName", sUnitName
	Root.setAttribute "BookNo", sBkNo
	Root.setAttribute "BookName", sBkName
	Root.setAttribute "CRDR", sCrDr
	Root.setAttribute "VouDate",sVDate
	Root.setAttribute "BookAcchead", sBkAccNo
	Root.setAttribute "Approver","Y"
	Root.setAttribute "TransNo", sCTNo
	Root.setAttribute "VoucherNo", sCVNo
	Root.setAttribute "CreatedTransNo", sCTNo
	Root.setAttribute "CreatedVoucherNo",sCVNo
	oDOM.appendchild Root


	'Newly added on 24 th June 2008 to add BankInstrument details by Maheswari
	Dim iInsEntNo,BkInsNo,BkInsDate,BkInsType,sPayAt,sDrwOn,iBkInsEntNo,iInsEntryNo1,BkNode,iBkInsAmt
	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "Select InstrumentEntryNo,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),BankInstrumentType,PayableAt,DrawnOnBank,"&_
				  "BankInstrumentEntryNo,isNull(InstrumentEntryNo1,0),InstrumentAmount from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo= "&iVNo&" "
		.ActiveConnection = con
		.Open
	End With

	'Response.Write rs1.source
	'Response.End
	Do while Not rs1.EOF
		iInsEntNo		= rs1(0)
		BkInsNo			= rs1(1)
		BkInsDate		= rs1(2)
		BkInsType		= rs1(3)
		sPayAt			= rs1(4)
		sDrwOn			= rs1(5)
		iBkInsEntNo		= rs1(6)
		iInsEntryNo1	= rs1(7)
		iBkInsAmt		= rs1(8)
		IF cint(iBkInsEntNo) <> 0 then
			BkInsNo = iBkInsEntNo &"-"&iInsEntryNo1&"-"&BkInsNo
		End IF
		set BkNode=oDOM.CreateElement("BankInstrumentDet")
		BkNode.setAttribute "SlNo", iInsEntNo
		BkNode.setAttribute "InsNo",BkInsNo
		BkNode.setAttribute "InsType",BkInsType
		BkNode.setAttribute "InsDate",BkInsDate
		BkNode.setAttribute "PayAt", sPayAt
		BkNode.setAttribute "DrawnOn",sDrwOn
		BkNode.setAttribute "InsAmt",iBkInsAmt
		IF cint(iBkInsEntNo) <> 0 then
			BkNode.setAttribute "Option", "Y"
		Else
			BkNode.setAttribute "Option", ""
		End IF
		BkNode.setAttribute "Action", "0"
		BkNode.setAttribute "BankInsEntNo",iBkInsEntNo 
		
		Root.appendchild BkNode

		rs1.MoveNext
	loop
	rs1.Close

		'creation of Entry node

	'With rs1
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source = "select count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo= "&iVNo&" "
	'	.ActiveConnection = con
	'	.Open
	'End With

	'If not rs1.EOF then
	'	iCount = rs1(0)
	'End If
	'rs1.Close

	with rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select isNull(PayToRecdFrom,''),IsNull(TDSGroupID,0) from Acc_T_CreatedVoucherHeader where CreatedTransNo = "&iVNo&" "
		.ActiveConnection = con
		.Open
	End with

	If not rs1.EOF then
		iPayToRecdFrom = rs1(0)
		iTdsGrpID	   = rs1(1)
	End IF
	rs1.Close

	'For i = 1 to iCount
		'sQry = "select Distinct VoucherEntryNumber,TransCrDrIndication,Amount,AccountingUnit,TdsOnAmount,TdsPercentage,isNull(TDSFlag,'N') from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&"  "
		sQry = "select Distinct VoucherEntryNumber,TransCrDrIndication,Amount,AccountingUnit,TdsOnAmount,TdsPercentage from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" and isNull(TDSFlag,'N') = 'N' "

		'	Response.Write sQry & "<BR><BR>"
		'	 Response.End

			with rs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con
				.Open
			End With

			Do While not rs1.EOF
				iENo		= rs1(0)
				iETCrDr		= rs1(1)
				iEAmt		= rs1(2)
				iEAccUnit	= rs1(3)
				iETdsAmt	= rs1(4)
				iETdsPer	= rs1(5)
				sAmount     = cdbl(iEAmt) + cdbl(iETdsAmt)
			'End If
			'rs1.Close
			Response.Write "iETdsPer="&iETdsPer
			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = " Select D.OrgUnitShortDescription from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherDetails A "&_
						  " where D.OUDefinitionID=A.AccountingUnit and A.CreatedTransNo = "&iVNo&" and A.VoucherEntryNumber = "&iENo&" "
				.ActiveConnection = con
				.Open
			End with
			If not rs2.EOF then
				sAccName = rs2(0)
			End IF
			rs2.Close

			set Elem1 = oDOM.CreateElement("Entry")
			Elem1.setAttribute "No",iENo
			Elem1.setAttribute "CRDR",iEtCrDr
			Elem1.setAttribute "Payto",iPayToRecdFrom
			Elem1.setAttribute "Amount",iEAmt
			'Elem1.setAttribute "Amount",sAmount
			Elem1.setAttribute "AccUnit",iEAccUnit
			Elem1.setAttribute "AccName",sAccName
			Elem1.setAttribute "TdsAmount",iETdsAmt
			Elem1.setAttribute "TDSElgi","0"
			Elem1.setAttribute "TdsPercentage",iETdsPer
			Elem1.setAttribute "PayRecAmount","0"
			Elem1.setAttribute "GroupId",iTds
			Root.Appendchild Elem1


			'creation of Acc Head

		With rsAcc
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
					 " and VoucherEntryNumber = "&iENo&" "
			.ActiveConnection = con
			.Open
		End with
		Do while not rsAcc.EOF
			sAccHeadNo = rsAcc(0)


            ''blocked and Added by ragav on Jan 13,2012 we only consider CostCenter and Analytical Eligible but here wrongly take the count 
            ''begin
			 '   With rs2
			'	    .CursorLocation = 3
			'	    .CursorType = 3
			'	    .Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
			'			      " and AccUnitAccountHead = "&sAccHeadNo&" "
			'	    .ActiveConnection = con
			'	    .Open
			 '   End With
			  '  If not rs2.EOF	then
			'	    sCostCnt = rs2(0)
			 '   End If
			 '   rs2.Close

			'    With rs2
			'	    .CursorLocation = 3
			'	    .CursorType = 3
			'	    .Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
			'			      " and AccUnitAccountHead = "&sAccHeadNo&" "
			'	    .ActiveConnection = con
			'	    .Open
			 '   End With
'
'			    If not rs2.EOF	then
'				    sAnlCnt = rs2(0)
'			    End If
'			    rs2.Close

            With rs2
                .CursorLocation = 3
                .CursorType = 3
                .Source = "Select isNull(CostCenterExists,0),isNull(AnalyticalheadExists,0) from Acc_M_GLAccountHead where AccountHead = " & sAccHeadNo 
                .ActiveConnection = con
                .Open 
            End With
            if not rs2.EOF then
                sCostCnt = rs2(0) 
                sAnlCnt = rs2(1)
            end if
            rs2.Close 

			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
				.ActiveConnection = con
				.Open
			End With
			If not rs2.EOF	then
				sAccHeadName = rs2(0)
			End If
			rs2.Close

		'Response.Write "sAccHeadNo="&sAccHeadNo
		If trim(sAccHeadNo) <> "0"  then
		
			Set Elem2 = oDOM.CreateElement("AccHead")
			Elem2.setAttribute "No",sAccHeadNo
			Elem2.setAttribute "CostCenter",sCostCnt
			Elem2.setAttribute "Analytical",sAnlCnt
			Elem2.setAttribute "Name",sAccHeadName
			Elem2.setAttribute "Type","G"
			Elem2.setAttribute "TransFlag","A"
			Elem1.Appendchild Elem2
		End If
		
		''added by ragav on Jan 13,2012
		''begin
		if sCostCnt>0 then
		    SCBQuery = "Select B.AccUnitCCHead,A.CCAccountDescription,A.CCHeadCode,B.CCRatioPercent,B.CCRatioAmount "&_
		               " from Acc_M_CCAccountHead A,Acc_T_CreatedVoucherCCDet B where CreatedTransNo = "& iVNo &" "&_
		               " and A.CostCenterHead = B.AccUnitCCHead and B.VoucherEntryNumber = "&iENo&" "
		    rs2.Open SCBQuery,con
		    if not rs2.EOF then
		        Set Elem2 = oDOM.CreateElement("CostCenter")
		        Elem1.AppendChild Elem2
		        do while not rs2.EOF 
		            Set Elem3 = oDOM.CreateElement("CC")
		            Elem3.SetAttribute "No",rs2(0)
		            Elem3.SetAttribute "Name",rs2(1)
		            Elem3.SetAttribute "ShortName",rs2(2)
		            Elem3.SetAttribute "Ratio",rs2(3)
		            Elem3.SetAttribute "Amount",rs2(4)
		            Elem2.AppendChild Elem3
		            rs2.MoveNext 
		        loop
		    end if
		    rs2.Close 
		end if
		
		if sAnlCnt>0 then
		    SCBQuery = " Select AnalyticalCode,AnalyticalName,AnalyticalShortName,RatioPercentage,RatioAmount,AccAHGroupCode "&_
		               " from Acc_M_AnalyticalHeads A,Acc_T_CretedVoucherAHDet B where CreatedTransNo = "& iVNo &" "&_
		               " and AnalyticalCode=AccUnitAnalyticalCode and B.VoucherEntryNumber = "&iENo&" "
		    rs2.Open SCBQuery,con
		    if not rs2.EOF then
		        Set Elem2 = oDOM.CreateElement("Analytical")
		        Elem1.AppendChild Elem2
		        do while not rs2.EOF 
		            Set Elem3 = oDOM.CreateElement("Anal")
		            Elem3.SetAttribute "No",rs2(0)
		            Elem3.SetAttribute "Name",rs2(1)
		            Elem3.SetAttribute "ShortName",rs2(2)
		            Elem3.SetAttribute "Ratio",rs2(3)
		            Elem3.SetAttribute "Amount",rs2(4)
		            Elem3.SetAttribute "GroupCode",rs2(5)
		            Elem2.AppendChild Elem3
		            rs2.MoveNext 
		        loop
		    end if
		    rs2.Close 
		end if
		''end

		'======================Newly Added On Jan 9th 2008======================================================
		'To check whether TDS already avoilable
		IF Root.haschildnodes then
			For each Node in Root.childnodes
				If Node.NodeName = "Entry" then
					For each EntNode in Node.Childnodes
						IF EntNode.NodeName = "TDS" then
							Set TDSNode = EntNode
						End IF
						'Root.RemoveChild TDSNode
					Next
				End If
			Next
		End IF 'IF Not Root.haschildnodes then
		'Response.Write " sAccPartyCode="& sAccPartyCode
		If sAccPartyCode <> "0" then
			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select Distinct GroupHeadID,isNull(ComputeFormula,0),AcHeadCode,ComputeMode from ACC_M_TDSHeadComputation where GroupID = "& iTdsGrpID &" "
				.ActiveConnection = con
				.Open
			End With
			'Response.Write "Chk="& rs2.Source
			Do while not rs2.EOF
				iTdsCtr		= iTdsCtr + 1
				iGrpHeadID	= rs2(0)
				iFormula	= rs2(1)
				iAcHeadCode = rs2(2)
				sCompMode	= rs2(3)
				'Response.Write "sCompMode="&sCompMode
				'Response.Write "Form="&iFormula
				IF sCompMode <> "F" then
					saTemp = split(iFormula,",")
					'MsgBox "OK 1 "
					IF UBound(saTemp) <> 0 Then
						iTemp = 0
					'	For iCounter=iTemp to UBound(saTemp)
							saTemp1 = Split(trim(saTemp(0)),"-")
							iTDSPer	=   saTemp1(1)
					'	Next
					else
						IF sCompMode <> "F" then
							If iFormula <> "0"	 then
								iTemp = split(iFormula,"-")
								iTDSPer = iTemp(1)
							else
								iTDSPer = iFormula
							End IF
						Else
							iTDSPer = iFormula
						End IF 'IF sCompMode <> "F" then
					End IF
				Else
					iTDSPer = iFormula
				End IF
				 'iVouEntNo =  Cdbl(iTdsCtr)+1

					with rs3
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "select Amount,TDSOnAmount,AccountingUnit,TDSPercentage from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" and isNull(TDSFlag,'N') <> 'N' and VoucherEntryNumber = "&iVouEntNo &""
						.Source = "select Amount,TDSOnAmount,AccountingUnit,TDSPercentage from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" and isNull(TDSFlag,'N') <> 'N' and AccUnitAccountHead = "& iAcHeadCode
						.ActiveConnection = con
						.Open
					End With
					  'Response.Write rs3.Source

					If  not rs3.EOF then
						iTDSAmt   = rs3(0)
						iTDSOnAmt = rs3(1)
					    iNewTdsPer = rs3(3)
					Response.Write "iNewTdsPer="&iNewTdsPer
						Set TDSElem = oDOM.createElement("TDS")
						TDSElem.setAttribute "Ctr",iTdsCtr
						TDSElem.setAttribute "AccHeadCode",iAcHeadCode
						TDSElem.setAttribute "TDSAmount",iTDSOnAmt
						TDSElem.setAttribute "TdsPercentage",iNewTdsPer
						TDSElem.setAttribute "PayRecAmount",iTDSAmt
						TDSElem.setAttribute "Formula",iFormula
						TDSElem.setAttribute "GroupHeadId",iTdsGrpID &"-"& 	iGrpHeadID
						Elem1.Appendchild TDSElem
					End If

				rs3.Close

				rs2.MoveNext

			loop
			rs2.Close


		End IF 'If sAccPartyCode <> "0" then
		Response.Write "sAccHeadNo="&sAccHeadNo

		'=====================================================================================
		If trim(sAccHeadNo) = "0" or trim(sAccHeadNo) = "Null"  then
			with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = " Select V.SubTypeName,A.AccUnitPartyType,A.AccUnitPartySubType,isNull(A.AccUnitPartyCode,0) from vwOrgPartyType V,Acc_T_CreatedVoucherDetails A where "&_
						  " A.AccUnitPartyType=V.PartyType and A.AccUnitPartySubType=V.PartySubType "&_
						  " and A.AccountingUnit = V.OUDefinitionID and A.CreatedtransNo = "&iVNo&" and A.VoucherEntryNumber = "&iENo&" "
				'Added By UmaMaheswari S,On April 2011 
				.Source = " Select V.SubTypeName,A.AccUnitPartyType,A.AccUnitPartySubType,isNull(A.AccUnitPartyCode,0) from vwOrgParty V,Acc_T_CreatedVoucherDetails A where "&_
						  " A.AccUnitPartyType=V.PartyType and A.AccUnitPartySubType=V.PartySubType "&_
						  " and A.AccountingUnit = V.OUDefinitionID and A.CreatedtransNo = "&iVNo&" and A.VoucherEntryNumber = "&iENo&" "
						  ''''iENo
				.ActiveConnection = con
				.Open
			End with

'			Response.Clear
			 Response.Write rs2.Source


			If not rs2.EOF then
				sSubTypeName = rs2(0)
				sAccPartyType = rs2(1)
				sAccPartySubType = rs2(2)
				sAccPartyCode = rs2(3)
			End If
			rs2.Close
			sTempNo = sAccPartyType &"?"& sAccPartySubType &"?"&sAccPartyType &"-"& sSubTypeName &"?"& sAccPartyCode

			
					'Response.Write "Name= " & sTempNo &"<BR>"
			IF trim(sAccPartyCode) <> "" then 
				With rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = "Select PartyName from App_M_PartyMaster where PartyCode = "& sAccPartyCode &" "
					.ActiveConnection = con
					.Open
				End With
				If not rs2.EOF	then
					sPartyName = rs2(0)
				End If
				rs2.Close
			Else
				sPartyName = ""
			End If
			'xml construction
			Set Elem2 = oDOM.CreateElement("AccHead")
			Response.Write "BBB"
			Elem2.setAttribute "No",sTempNo
			Elem2.setAttribute "Pay","0"
			Elem2.setAttribute "Rec","0"
			Elem2.setAttribute "Name",sPartyName
			Elem2.setAttribute "Type","P"
			Elem2.setAttribute "Adv","0"
			Elem1.Appendchild Elem2

	'next
				Set Elem3 = oDOM.createElement("PayRec")
				Elem1.Appendchild Elem3

				sQry = " Select C.PayablesNumber,isNull(C.Narration,''),Convert(Varchar,C.VoucherDate,103),C.AmountPayable,P.AmountPaid,C.CreatedTransNo "&_
					   " From Acc_T_CreatedPybleAdjDet P,Acc_T_CreatedPayables C Where P.CreatedTransNo = "&iVNo&" "&_
					   " and isNull(P.VoucherEntryNumber,0) = "&Cint(iENo)-0&" "&_
					   " and P.AdjustType is NULL and C.PayablesNumber = P.PayablesNumber "

				'Response.Write "Test1="&sQry&"<BR>"
				



				with rs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con
					.Open
				End with

				Do while not rs3.EOF
					sPayNo = rs3(0)
					sPInvNo = rs3(1)
					sPInvDate = rs3(2)
					sPTransAmt = rs3(3)
					sPAmdToAdj = rs3(4)
					sPCTransNo = rs3(5)

					sQry = "select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& sPCTransNo &" "
					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQry
						.ActiveConnection = con
						.Open
					End with

					If not rs2.EOF then
						sPDocType = rs2(0)
					End If
					rs2.close

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select PayablesNumber from Acc_T_Payables where CreatedPayablesNumber = "&sPayNo&" "
						.ActiveConnection = con
						.Open
					End with

					If not rs2.EOF then
						sPayableNo = rs2(0)
					End If
					rs2.close

					If trim(sPayableNo) = "" then
						With rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "select PayablesNumber from Acc_T_Payables where CRCreatedPayable = "&sPayNo&" "
							.ActiveConnection = con
							.Open
						End with
						If not rs2.EOF then
							sPayableNo = rs2(0)
						End If
						rs2.close
					End If

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select BookCode from Acc_T_CreatedVoucherHeader where CreatedTransNo ="& sPCTransNo &" "
						.ActiveConnection = con
						.Open
					End with
					If not rs2.EOF then
						sPBookCode = rs2(0)
					End If
					rs2.close

					If sPBookCode = "04" then sPAdjType = "PI"
					If sPBookCode = "05" then sPAdjType = "I"
					If sPBookCode = "06" then sPAdjType = "D"
					If sPBookCode = "07" then sPAdjType = "C"

					set Elem4 = oDOM.createElement("Doc")
					Elem4.SetAttribute "No",sPayNo
					Elem4.SetAttribute "InvNo",sPInvNo
					Elem4.SetAttribute "InvDate",sPInvDate
					Elem4.SetAttribute "TransAmount",sPTransAmt
					Elem4.SetAttribute "AmtAdjusted","0"
					Elem4.SetAttribute "AmtToAdjust",sPAmdToAdj
					Elem4.SetAttribute "DocType",sPDocType
					Elem4.SetAttribute "AmtToAccount","0"
					Elem4.SetAttribute "PayableNo",sPayableNo
					Elem4.SetAttribute "AdjType",sPAdjType
					Elem3.Appendchild Elem4

					rs3.MoveNext
				loop
				rs3.Close

				sQry ="Select C.ReceivableNumber,C. Narration,Convert(Varchar,C.VoucherDate,103),C.AmountReceivable,P.AmountReceived,C.CreatedTransNo "&_
				      " From Acc_T_CreatedRcvbleAdjDet P,Acc_T_CreatedReceivables C Where P.CreatedTransNo = "&iVNo&" "&_
				      " and isNull(P.VoucherEntryNumber,0) = "&Cint(iENo)-0&" "&_
				      " and P.AdjustType is NULL and C.ReceivableNumber = P.ReceivableNumber "
				Response.Write "Test="&sQry
				'Response.End 
				with rs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con
					.Open
				End with

				Do while not rs3.EOF
					sRecNo = rs3(0)
					sRInvNo = rs3(1)
					sRInvDate = rs3(2)
					sRTransAmt = rs3(3)
					sRAmdToAdj = rs3(4)
					sRCTransNo = rs3(5)

					sQry = "select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& sRCTransNo &" "
					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQry
						.ActiveConnection = con
						.Open
					End with
					If not rs2.EOF then
						sRDocType = rs2(0)
					End If
					rs2.Close

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select ReceivableNumber from Acc_T_Receivables where CreatedReceivable = "& sRecNo &" "
						.ActiveConnection = con
						.Open
					End with

					If not rs2.EOF then
						sReceivableNo = rs2(0)
					End If
					rs2.Close

					If trim(sReceivableNo) = "" then
						With rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "select ReceivableNumber from Acc_T_Receivables where DRCreatedReceivable = "& sRecNo &" "
							.ActiveConnection = con
							.Open
						End with
						If not rs2.EOF then
							sReceivableNo = rs2(0)
						End If
						rs2.Close
					End If

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select BookCode from Acc_T_CreatedVoucherHeader where CreatedTransNo ="& sRCTransNo &" "
						.ActiveConnection = con
						.Open
					End with
					If not rs2.EOF then
						sRBookCode = rs2(0)
					End If
					rs2.close

					If sRBookCode = "04" then sRAdjType = "PI"
					If sRBookCode = "05" then sRAdjType = "I"
					If sRBookCode = "06" then sRAdjType = "D"
					If sRBookCode = "07" then sRAdjType = "C"

					set Elem4 = oDOM.createElement("Doc")
					Elem4.SetAttribute "No",sRecNo
					Elem4.SetAttribute "InvNo",sRInvNo
					Elem4.SetAttribute "InvDate",sRInvDate
					Elem4.SetAttribute "TransAmount",sRTransAmt
					Elem4.SetAttribute "AmtAdjusted","0"
					Elem4.SetAttribute "AmtToAdjust",sRAmdToAdj
					Elem4.SetAttribute "DocType",sRDocType
					Elem4.SetAttribute "AmtToAccount","0"
					Elem4.SetAttribute "PayableNo",sReceivableNo
					Elem4.SetAttribute "AdjType",sRAdjType
					Elem3.Appendchild Elem4

					rs3.movenext
				loop
				rs3.Close

				sQry = "SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, Convert(Varchar,V.VoucherDate,103), V.VoucherAmount, P.AmountPaid, V.CreatedTransNo "&_
					   " FROM Acc_T_CreatedVoucherHeader V INNER JOIN Acc_T_CreatedAdvances A ON V.CreatedTransNo = A.CreatedTransNo "&_
					   " INNER JOIN Acc_T_CreatedPybleAdjDet P ON A.CreatedAdvanceNo = P.PayablesNumber WHERE P.AdjustType IS NOT NULL AND "&_
					   " V.BookCode IN ('01', '02') and P.CreatedTransNo = "& iVNo &" and P.VoucherEntryNumber = "&iENo&" "

			With rs3
			  	.CursorLocation = 3
			  	.CursorType = 3
			  	.Source = sQry
			  	.ActiveConnection = con
			  	.Open
			End with

			Do while not rs3.EOF
				sDNo = rs3(0)
				sDInvNo = rs3(0)
				sDInvDate = rs3(2)
				sDTransAmt = rs3(3)
				sDAmdToAdj = rs3(4)
				sDCTransNo = rs3(5)

				with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source ="Select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& sDCTransNo &" "
					.ActiveConnection = Con
					.Open
				End with
				If not rs2.EOF then
					sDDocType = rs2(0)
				End If
				rs2.Close

				with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source ="Select AdvanceNumber from Acc_T_AdvancePayments where CreatedAdvanceNo = "& sDNo &" "
					.ActiveConnection = Con
					.Open
				End with
				If not rs2.EOF then
					sDPayableNo = rs2(0)
				End If
				rs2.Close

				set Elem4 = oDOM.createElement("Doc")
				Elem4.SetAttribute "No",sDNo
				Elem4.SetAttribute "InvNo","Advance Payments" &" "& sDInvNo
				Elem4.SetAttribute "InvDate",sDInvDate
				Elem4.SetAttribute "TransAmount",sDTransAmt
				Elem4.SetAttribute "AmtAdjusted","0"
				Elem4.SetAttribute "AmtToAdjust",sDAmdToAdj
				Elem4.SetAttribute "DocType",sDDocType
				Elem4.SetAttribute "AmtToAccount","0"
				Elem4.SetAttribute "PayableNo",sDPayableNo
				Elem4.SetAttribute "AdjType","P"
				Elem3.Appendchild Elem4
				rs3.MoveNext
			loop
			rs3.Close

			sQry = "SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, Convert(Varchar,V.VoucherDate,103), V.VoucherAmount, R.AmountReceived,V.CreatedTransNo "&_
				   " FROM Acc_T_CreatedVoucherHeader V INNER JOIN Acc_T_CreatedAdvances A ON V.CreatedTransNo = A.CreatedTransNo "&_
				   " INNER JOIN Acc_T_CreatedRcvbleAdjDet R ON A.CreatedAdvanceNo = R.ReceivableNumber WHERE R.AdjustType IS NOT NULL "&_
				   " AND V.BookCode IN ('01', '02') and R.CreatedTransNo = "& iVNo &" and R.VoucherEntryNumber = "&iENo&" "
				   'Response.Write "Rcv="&sQry
				With rs3
				  	.CursorLocation = 3
				  	.CursorType = 3
				  	.Source = sQry
				  	.ActiveConnection = con
				  	.Open
				End with

				Do while not rs3.EOF
					iDNo = rs3(0)
					iDInvNo = rs3(0)
					iDInvDate = rs3(2)
					iDTransAmt = rs3(3)
					iDAmdToAdj = rs3(4)
					iDCTransNo = rs3(5)

					with rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source ="Select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iDCTransNo &" "
						.ActiveConnection = Con
						.Open
					End with

					If not rs2.EOF then
						iDDocType = rs2(0)
					End If
					rs2.Close

					with rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source ="Select AdvanceNumber from Acc_T_AdvancePayments where CreatedAdvanceNo = "& iDNo &" "
						.ActiveConnection = Con
						.Open
					End with
					If not rs2.EOF then
						iDPayableNo = rs2(0)
					End If
					rs2.Close

					set Elem4 = oDOM.createElement("Doc")
					Elem4.SetAttribute "No",iDNo
					Elem4.SetAttribute "InvNo","Advance Receipts" &" "& iDInvNo
					Elem4.SetAttribute "InvDate",iDInvDate
					Elem4.SetAttribute "TransAmount",iDTransAmt
					Elem4.SetAttribute "AmtAdjusted","0"
					Elem4.SetAttribute "AmtToAdjust",iDAmdToAdj
					Elem4.SetAttribute "DocType",iDDocType
					Elem4.SetAttribute "AmtToAccount","0"
					Elem4.SetAttribute "PayableNo",iDPayableNo
					Elem4.SetAttribute "AdjType","R"
					Elem3.Appendchild Elem4
					rs3.MoveNext
				loop
				rs3.Close
		End IF	'If trim(sAccHeadNo) = "Null" then
				with rs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = "select VoucherNarration from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "&_
							  " and VoucherEntryNumber = "&iENo&" "
					.ActiveConnection = con
					.Open
				End with
				If not rs3.EOF then
					sNarr = rs3(0)
				End If
				rs3.close

				Set Elem5 = oDom.CreateElement("Narration")
				Elem5.text = sNarr
				Elem1.appendchild Elem5
			rsAcc.MoveNext
		loop
		rsAcc.Close

	rs1.MoveNext
	loop
	rs1.close

	If trim(sPassCode) = "02" then
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select isNull(BankInstrumentType,''),isNull(BankInstrumentNo,''),Convert(Varchar,BankInstrumentDate,103),isNull(Payableat,''),isNull(DrawnOnBank,'') "&_
					  " from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iVNo &" and BankInstrumentType is NOT NULL "
			.ActiveConnection = con
			.Open
		End with

		If not rs1.EOF then
			sBInsType = rs1(0)
			sBInsNo = rs1(1)
			sBInsDate = rs1(2)
			sBPay = rs1(3)
			sBDrwOn = rs1(4)
		End If
		rs1.close

		IF CStr(sBInsType) <> "" Then
			'Set BankElem = oDOM.CreateElement("BankInstrumentDet")
			'BankElem.SetAttribute "InsType",sBInsType
			'BankElem.SetAttribute "InsNo",sBInsNo
			'BankElem.SetAttribute "InsDate",sBInsDate
			'BankElem.SetAttribute "PayAt",sBPay
			'BankElem.SetAttribute "Drawnon",sBDrwOn
			'Elem1.appendchild BankElem
		End IF
	End IF

	Response.Clear
'	Response.Write "dfsfds"
	'Response.end
	IF CStr(sDispTy) <> "S" Then
		Response.ContentType = "text/XML"
		Response.Write oDOM.xml
	Else
		Response.ContentType = "text/HTML"
		oDOM.save Server.MapPath("../temp/transaction/"&iVNo&".xml")
		'IF Left(sTransType,2) = "CA" then
		'	oDOM.save Server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
		'ElseIF Left(sTransType,2) = "BA" then
		'	oDOM.save Server.MapPath("../temp/transaction/Voucher AMD_BA_"&Session.SessionID&".xml")
		'End IF
		'GetXmlForCBG = "../temp/transaction/VoucherCBG_xml_"&Session.SessionID&".xml"
		GetXmlForCBG = "../temp/transaction/"&iVNo&".xml"
	End IF
End Function
%>







