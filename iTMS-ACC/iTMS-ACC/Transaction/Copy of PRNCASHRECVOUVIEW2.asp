<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PRNCashRecVouView2.asp
	'Module Name				:	Accounts (Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	26 March 2004
	'Modified On				:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->
<!--#include File="../../include/PrintFunctions.asp"-->
<%
'------------------------Declaration Constants -----------------------------
Dim sPrHeader,sPrFooter,sPgMargin,sPgBreak,iPgLineNo,iRecCount,sDisplayHead
Dim aiHeadColWidth(3,7),aiUtilColWidth(3,8),sPagetitle2,sTstr,objFSO,objTxt,sFinalText
Dim objFSO2,objTxt2,iLineno,i,jText,j,iMyPos,sAddDate,jCtr,DivStr,sPurAmt,z,pFlag

sPrHeader=""
sPrFooter=""
sPgMargin="       "
sPgBreak= FormattPrint("PAGESKIP","")
iPgLineNo=72

'WIDTH SPECIFICATION FOR PAGE TITLE LINE 1 AND 2
aiHeadColWidth(0,0)=80
aiHeadColWidth(0,1)=65
aiHeadColWidth(0,2)=28
aiHeadColWidth(0,3)=20
aiHeadColWidth(0,4)=60
aiHeadColWidth(0,5)=70
aiHeadColWidth(0,6)=10
aiHeadColWidth(0,7)=5

'WIDTH SPECIFICATION FOR PAGE HEADER LINE 1
aiHeadColWidth(1,0)=50
aiHeadColWidth(1,1)=20
aiHeadColWidth(1,2)=15
aiHeadColWidth(1,3)=10
aiHeadColWidth(1,4)=30

'WIDTH SPECIFICATION FOR PAGE HEADER LINE 2
aiHeadColWidth(2,0)=6
aiHeadColWidth(2,1)=35
aiHeadColWidth(2,2)=30
aiHeadColWidth(2,3)=41
aiHeadColWidth(2,4)=5
aiHeadColWidth(2,5)=20

'WIDTH SPECIFICATION FOR PAGE HEADER LINE 3
aiHeadColWidth(3,0)=67
aiHeadColWidth(3,1)=13
aiHeadColWidth(3,2)=13
aiHeadColWidth(3,3)=10

'WIDTH SPECIFICATION FOR OPENING/CLOSING  LINE
aiUtilColWidth(0,0)=40
aiUtilColWidth(0,1)=12
aiUtilColWidth(0,2)=28
aiUtilColWidth(0,3)=30
aiUtilColWidth(0,4)=15
aiUtilColWidth(0,5)=5
aiUtilColWidth(0,6)=2
aiUtilColWidth(0,7)=1

'WIDTH SPECIFICATION FOR DETAIL LINE
aiUtilColWidth(1,0)=10
aiUtilColWidth(1,1)=15
aiUtilColWidth(1,2)=20
aiUtilColWidth(1,3)=25
aiUtilColWidth(1,4)=30
aiUtilColWidth(1,5)=35
aiUtilColWidth(1,6)=27

'Width Specification For Details Line 2
aiUtilColWidth(2,0)=65
aiUtilColWidth(2,1)=10
aiUtilColWidth(2,2)=10
aiUtilColWidth(2,3)=10

'Width Specification For Details Line 3
aiUtilColWidth(3,0)=5
aiUtilColWidth(3,1)=10
aiUtilColWidth(3,2)=10
aiUtilColWidth(3,3)=10
aiUtilColWidth(3,4)=20
aiUtilColWidth(3,5)=10
aiUtilColWidth(3,6)=10

'WIDTH SPECIFICATION FOR TOTAL LINE
aiUtilColWidth(2,0)=82
aiUtilColWidth(2,1)=13
aiUtilColWidth(2,2)=13
aiUtilColWidth(2,3)=13

'------------------------End of Declaration Constants ----------------------
%>
<%
Dim sOrgId,sOrgName,sBookCode,sBookName,sVouCherType,sTransNo,sQuery
Dim iVouNo,objRs,objRs1,sVouDate,sGetVal,iPageno,sAccHead,sParType
Dim iEntryNo,sAccUnit,sAmount,sCrDr,sGroupCode,sPartSubType,iCtr
Dim iEnNo,Entrynode,HeaderNode,sTemp,sAccName,sQuery1
Dim sParCode,sNarration,sAccHeadname,sAccUnitName
Dim sAccType,nodADD,dCRAmt,dDRAmt,sNoAddNar,AccVoucherNo
Dim sBankInsType,sBankInsNo,sBankInsDate,sBankName,sPartyName
Dim sBlankLine,iPageLen,sTemp1,sPurNoAmt,sPayableNo
Dim saTemp,sShtOrgId,sVouNo,Node,sTemp2,Objrs2,sQuery2,iNoLen

Dim sVoucherDate,iVoucherNo,iSumAmount,sAccount1
iSumAmount = 0
sPurAmt=0
sBlankline = string(80," ") & vbCrLf
iPageLen = formattPrint("PAGELEN6IN","")

Dim sAccount,sAddtional,iSno
Dim dTotal,dTempTotal
Dim iBookCode,sPayTo
'XML DOM Variables
Dim oDOM,nodHeader,Root
Dim sAddInvDet,sAddDebDet,sAddCreDet,sAddAdvDet

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
Set Objrs2 = Server.CreateObject("ADODB.RecordSet")

'----------- To Get The Values From the Selection Page ----------------
Dim  iTransNo,sAccNo,sPartyAdd1,sPartyCity,sPartyState,sArrTemp
iTransNo=Request("Value")

'oDOM.load server.MapPath("../XmlData/Voucher/"&iTransNo&".xml")
oDOM.load server.MapPath(GetVouchXML(iTransNo))
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/"&Session.SessionID&"_CashPay_View.txt"))

set Root=oDOM.documentElement

sOrgId=Root.Attributes.Item(0).nodeValue
sOrgName=Root.Attributes.Item(1).nodeValue
iBookCode=Root.Attributes.Item(2).nodeValue
sBookName=Root.Attributes.Item(3).nodeValue
sVouCherType=Root.Attributes.Item(4).nodeValue

sShtOrgId = Right(sOrgId,2)

set EntryNode= Root.childNodes(0)
set HeaderNode=EntryNode.childNodes(0)
'sPayTo = HeaderNode.Attributes.Item(3).nodeValue

IF Len(Trim(iBookCode)) = 1 Then
	iBookCode = "0"&Trim(iBookCode)
End IF

sVouDate = Root.Attributes.Item(5).nodeValue
sVoucherDate = sVouDate

set EntryNode= Root.childNodes(0)
sPayTo= Entrynode.Attributes.Item(2).nodeValue
set HeaderNode=EntryNode.childNodes(0)

ConsPgHeader1(1)

sQuery = "Select G.AccountsGroupName From Acc_M_AccountGroups G,Acc_R_ApplicableAccountHeads H, "&_
		 "Acc_R_OrgGLAccountHead O Where H.OUDefinitionID='"&sOrgId&"' and H.BookCode='"&iBookCode&"' and  "&_
		 "H.BookAccountHead = O.AccountHead and O.AccountsGroupCode = G.AccountsGroupCode "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sAccName = objRs(0)
End IF
objRs.Close

sQuery = "Select isNull(BankInstrumentType,' '),isNull(BankInstrumentNo,' '), "&_
		 "isNull(Convert(char,BankInstrumentDate,103),' '),isNull(DrawnOnBank,''),isNull(PayToRecdFrom,'') From  "&_
		 "Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iTransNo&" "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sBankInsType = objRs(0)
	sBankInsNo = objRs(1)
	sBankInsDate = objRs(2)
	sBankName = objRs(3)
	sPartyName = objRs(4)
End IF
objRs.Close

sQuery = "SELECT VOUCHERNARRATION FROM ACC_T_CREATEDVOUCHERDETAILS WHERE CREATEDTRANSNO = "&iTransNo&" "
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing
Do While Not objRs.EOF
	sNarration = sNarration&","&objRs(0)
	objRs.MoveNext
loop
objRs.Close

sNarration = Mid(sNarration,2)



for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
			sAccNo = Split(HeaderNode.attributes.item(0).nodeValue,"?")
			sAccType=HeaderNode.Attributes.Item(4).nodeValue
		End IF
Next

	IF sAccType = "P" Then
			sQuery2 = "SELECT PARTYNAME,ADDRESSLINE1,CITY,STATE FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&sAccNo(3)

			With Objrs2
				.ActiveConnection = con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery2
				.Open
			End With
			Set Objrs2.ActiveConnection = nothing

			IF Objrs2.RecordCount > 0 Then
				sPartyName = Trim(Objrs2(0))
				sPartyAdd1 = Trim(Objrs2(1))
				sPartyCity = Trim(Objrs2(2))
				sPartyState = Trim(Objrs2(3))
			End IF
		End IF

	'	Objrs2.close

	IF sAccType = "P" Then
	
	    if trim(sVouCherType)="D" then
		    sTstr = sTstr & myalign("Revd.From "&sPartyName,aiHeadColWidth(0,0),"L") & VbCrlf
		else
		    sTstr = sTstr & myalign("Paid To "&sPartyName,aiHeadColWidth(0,0),"L") & VbCrlf
		end if 
		sTstr = sTstr & string(9," ")& myalign(sPartyAdd1,64,"L") & VbCrlf
		sTstr = sTstr & string(9," ")& myalign(sPartyCity,64,"L") & VbCrlf
		sTstr = sTstr & string(9," ")& myalign(sPartyState , 64,"L") & VbCrlf
		sTstr = sTstr & string(9," ")& myalign(" ",64,"L") & VbCrlf
	Else
	    if trim(sVouCherType)="D" then
		    sTstr = sTstr & myalign("Revd.From "&sPayTo,aiHeadColWidth(0,0),"L") & VbCrlf
		else
		    sTstr = sTstr & myalign("Paid To "&sPayTo,aiHeadColWidth(0,0),"L") & VbCrlf
		end if 
		sTstr = sTstr &  myalign(" ",64,"L") & VbCrlf
		sTstr = sTstr &  myalign(" ",64,"L") & VbCrlf
		sTstr = sTstr &  myalign(" ",64,"L") & VbCrlf
	End IF

	IF Len(Trim(sBankInsType)) <> 0 Then sTemp1 = sBankInsType
	IF Len(Trim(sBankInsNo)) <> 0 Then sTemp1 = sTemp1 &" - "&sBankInsNo
	IF Len(Trim(sBankInsDate)) <> 0 Then sTemp1 = sTemp1 &" - "&sBankInsDate

	'set objTxt2 = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/"&Session.SessionID&"_CashPaySec_View.txt"))

	iPageno = 1
	iEnNo =0
	iSumAmount = 0
	z=0

for each EntryNode in Root.childNodes
z=z+1

	if Entrynode.nodeName="Entry" then
		iSno=EntryNode.Attributes.Item(0).nodeValue
		sAmount=EntryNode.Attributes.Item(3).nodeValue
		if EntryNode.Attributes.Item(1).nodeValue="C" then
			dCRAmt=dCRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			dTempTotal=dTempTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		elseif EntryNode.Attributes.Item(1).nodeValue="D" then
			dDRAmt=dDRAmt+CDbl(EntryNode.Attributes.Item(3).nodeValue)
			dTempTotal=dTempTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		end if

		if EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "C" then
			iSumAmount = iSumAmount + sAmount
		elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "C" then
			iSumAmount = iSumAmount - sAmount
		elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "D" then
			iSumAmount = iSumAmount + sAmount
		elseif EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "D" then
			iSumAmount = iSumAmount - sAmount
		end if

		sAddtional=""
	for each HeaderNode in EntryNode.childNodes

		if HeaderNode.nodeName="AccHead" then
			sAccType=HeaderNode.Attributes.Item(4).nodeValue
			if sAccType="G" then
				'sAccount1=HeaderNode.Attributes.Item(0).nodeValue
				sAccount=HeaderNode.Attributes.Item(3).nodeValue
			else
				'sAccount=HeaderNode.Attributes.Item(3).nodeValue
				sAccount=" "
			end if
		end if 'End of Check for Account head Node

		'IF HeaderNode.nodeName="Narration" Then
		'	sNoAddNar = HeaderNode.text
		'End IF
		sNoAddNar = sNarration


		IF HeaderNode.nodeName="PayRec" Then
			For Each Node in HeaderNode.childNodes
				IF Trim(Node.nodeName) = "Doc" Then
					if Node.Attributes.Item(5).nodeValue > 0 then
						IF Left(sParType,2) = "CR" Then
							Select Case Cstr(Node.Attributes.GetNamedItem("AdjType").value)
								Case "PI"
									sAddInvDet=sAddInvDet&""
									sAddInvDet=sAddInvDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddInvDet=sAddInvDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Cdbl(sPurAmt) + CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value)
								Case "D"
									sAddDebDet=sAddDebDet&"Less "
									sAddDebDet=sAddDebDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddDebDet=sAddDebDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Abs(Cdbl(sPurAmt) - CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value))
								Case "C"
									sAddCreDet=sAddCreDet&"Add "
									sAddCreDet=sAddCreDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddCreDet=sAddCreDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Cdbl(sPurAmt) + CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value)
								Case "P"
									sAddAdvDet=sAddAdvDet&"Less "
									sAddAdvDet=sAddAdvDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddAdvDet=sAddAdvDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Abs(Cdbl(sPurAmt) - CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value))
							End Select
							sPurNoAmt = sAddInvDet
							sPurNoAmt = sPurNoAmt&sAddCreDet
							sPurNoAmt = sPurNoAmt&sAddDebDet
							sPurNoAmt = sPurNoAmt&sAddAdvDet
						Else
							Select Case Cstr(Node.Attributes.GetNamedItem("AdjType").value)
								Case "I"
									sAddInvDet=sAddInvDet&""
									sAddInvDet=sAddInvDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddInvDet=sAddInvDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Cdbl(sPurAmt) + CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value)
								Case "D"
									sAddDebDet=sAddDebDet&"Add "
									sAddDebDet=sAddDebDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddDebDet=sAddDebDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Cdbl(sPurAmt) + CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value)
								Case "C"
									sAddCreDet=sAddCreDet&"Less "
									sAddCreDet=sAddCreDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddCreDet=sAddCreDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Abs(Cdbl(sPurAmt) - CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value))
								Case "R"
									sAddAdvDet=sAddAdvDet&"Less "
									sAddAdvDet=sAddAdvDet&Node.Attributes.GetNamedItem("InvNo").value&"---"
									sAddAdvDet=sAddAdvDet&FormatNumber(Node.Attributes.GetNamedItem("AmtToAdjust").value,2,,,0)&"&"
									sPurAmt = Abs(Cdbl(sPurAmt) - CDbl(Node.Attributes.GetNamedItem("AmtToAdjust").value))
							End Select
							sPurNoAmt = sAddInvDet
							sPurNoAmt = sPurNoAmt&sAddDebDet
							sPurNoAmt = sPurNoAmt&sAddCreDet
							sPurNoAmt = sPurNoAmt&sAddAdvDet
						End IF
						'sPurNoAmt = sPurNoAmt &  Node.Attributes.Item(1).nodeValue & "-"& FormatNumber(Node.Attributes.item(5).nodeValue,2,,,0) & "&"
						'sPurAmt=sPurAmt+ Node.Attributes.item(5).nodeValue
					end if
				End IF
			Next
		End IF

		'if 	HeaderNode.nodeName="Narration" then
		'	sNarration=HeaderNode.text
		'end if 'End of Check for Narration Node

		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodADD in HeaderNode.childNodes
				sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
				sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &" "
				sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
			next
			sAddtional = Mid(sAddtional,1,25)
		end if 'End of Check for Cost Center Node

		if 	HeaderNode.nodeName="Analytical" then
			for each  nodADD in HeaderNode.childNodes
				sAddtional=sAddtional&nodADD.Attributes.Item(2).nodeValue&"-"
				sAddtional=sAddtional&nodADD.Attributes.Item(3).nodeValue &" "
				sAddtional=sAddtional&nodADD.Attributes.Item(4).nodeValue&"<br>"
			next
			sAddtional = Mid(sAddtional,1,25)
		end if 'End of Check for Analytical Node

	next 'End of Entry Node Loop

	sAmount = FormatNumber(sAmount,2,,,0)
	sPurAmt = FormatNumber(sPurAmt,2,,,0)

		objtxt.write sTstr
		sTstr=""

		IF CStr(sPartyName) <> "" Then
			if z=1 then
				i=1
				'if InStr(1,len(trim(sNoAddNar))/80,".") then
				'	jCtr=len(trim(sNoAddNar))/80+1
				'else
				'	jCtr=len(trim(sNoAddNar))/80
				'end if
				if CStr(jCtr)=""then
					objTxt.write  myalign(jText,aiHeadColWidth(0,0),"L") & VbCrlf
				end if
				if len(sNoAddNar)> 80 then
					jCtr=Len(sNoAddNar)
					for j=1 to CInt(jCtr)
						if j=1 then
							jText=Mid(sNoAddNar,i,80)
						else
							jText=Mid(sNoAddNar,i)
						end if

						if len(trim(jText))=0 then exit for

						objTxt.write myalign(trim(jText),aiHeadColWidth(0,0),"L") & VbCrlf
						i=i+80
						next
				else
					objtxt.write sNoAddNar & vbCrLf
				end if
		End IF

			pFlag=true
			if sAccType="P" then
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
				objTxt.write myalign("Particulars",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			else
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
				objTxt.write myalign("A/c Head",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			end if

		end if

		if not pFlag then
			if sAccType="P" then
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
				objTxt.write myalign("Particulars",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			else
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
				objTxt.write myalign("A/c Head",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
				objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			end if
		end if
	IF sPurNoAmt <> "" Then
		if EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "C" then
		elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "C" then
			'sTstr = sTstr & myalign(sPurNoAmt,aiHeadColWidth(2,4)+aiHeadColWidth(2,2)+4,"L") & VbCrlf
			'sTstr = sTstr & myalign(sTemp2,aiHeadColWidth(2,4)+aiHeadColWidth(2,2)+4,"L")  & myalign(sAmount,aiHeadColWidth(2,3),"R") & VbCrlf
		elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "D" then
		elseif EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "D" then
			'sTstr = sTstr & myalign(sPurNoAmt,aiHeadColWidth(2,4)+aiHeadColWidth(2,2)+4,"L") & VbCrlf
			'sTstr = sTstr & myalign(sTemp2,aiHeadColWidth(2,4)+aiHeadColWidth(2,2)+4,"L")  & myalign(sAmount,aiHeadColWidth(2,3),"R") & VbCrlf
		end if
		DivStr=split(sPurNoAmt,"&",-1,1)
		for j=0 to ubound(DivStr)-1
		iLineNo=iLineNo+1
		iEnNo =iEnNo +1
		sArrTemp=Split(DivStr(j),"---")
			if j=0 then
				sTstr=sTstr & myalign(sArrTemp(0),aiHeadColWidth(2,3)+aiHeadColWidth(2,4)-1,"L")& "---" & myalign(sArrTemp(1),12,"R") & myalign(FormatNumber(sAmount),aiHeadColWidth(2,5),"R")&  VbCrlf
			else
				sTstr=sTstr &  myalign(sArrTemp(0),aiHeadColWidth(2,3)+aiHeadColWidth(2,4)-1,"L") & "---"& myalign(sArrTemp(1),12,"R")& VbCrlf
			end if
		if iLineNo>26 then
			iPageno =iPageno+ 1
			objtxt.write  sBlankLine
			objTxt.write myalign(" ",aiHeadColWidth(0,1)+5,"L")& myalign("Contd ...",aiHeadColWidth(0,2),"L") & VbCrlf
			ConsPgHeader1 (iPageno )
		end if
		next
		if sPurAmt < sAmount  then
			sTstr=sTstr& "Advance Paid :"& myalign(FormatNumber(sAmount-sPurAmt,2,,,0),aiHeadColWidth(1,3),"R")&vbCrLf
		end if
	Else
	iLineNo=iLineNo+1
	iEnNo =iEnNo +1
		If sAddtional <> "" Then
			if EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "C" then
				sTstr = sTstr & myalign(""&sAddtional,aiHeadColWidth(2,4),"L") & string(2," ") & myalign(sAccount,aiHeadColWidth(2,2),"L")  & string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(2,3),"R") & VbCrlf
			elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "C" then
				sTstr = sTstr & myalign(""&sAddtional,aiHeadColWidth(2,4),"L") & string(2," ") & myalign(sAccount,aiHeadColWidth(2,2),"L")  & string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(2,3),"R") & VbCrlf
			elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "D" then
				sTstr = sTstr & myalign(""&sAddtional,aiHeadColWidth(2,4),"L") & string(2," ") & myalign(sAccount,aiHeadColWidth(2,2),"L")  & string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(2,3),"R") & VbCrlf
			elseif EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "D" then
				sTstr = sTstr & myalign(""&sAddtional,aiHeadColWidth(2,4),"L") & string(2," ") & myalign(sAccount,aiHeadColWidth(2,2),"L")  & string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(2,3),"R") & VbCrlf
			end if
			sAddtional = ""
			'sTstr = sTstr & myalign(sAddtional,aiHeadColWidth(2,4),"L") & string(2," ") & myalign(sAccount,aiHeadColWidth(2,2),"L")  & string(2," ") & myalign(sAmount,aiHeadColWidth(2,3),"R") & VbCrlf
		Else
			If CStr(trim(sAccount))="" then
				if EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "C" then
					sTstr = sTstr & myalign(sNoAddNar,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
					i=51
					IF CDbl(Len(sNoAddNar)) > 50 Then
						For iNoLen = 51 to Len(sNoAddNar) Step 50
							sTstr = sTstr & myalign(Mid(Trim(sNoAddNar),iNoLen,50),aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") & VbCrlf
							i = CDbl(i) + 1
						Next
					End IF
				elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "C" then
					sTstr = sTstr & myalign(sNoAddNar,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
					i=51
					IF CDbl(Len(sNoAddNar)) > 50 Then
						For iNoLen = 51 to Len(sNoAddNar) Step 50
							sTstr = sTstr & myalign(Mid(Trim(sNoAddNar),iNoLen,50),aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") & VbCrlf
							i = CDbl(i) + 1
						Next
					End IF
				elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "D" then
					sTstr = sTstr & myalign(sNoAddNar,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
					i=51
					IF CDbl(Len(sNoAddNar)) > 50 Then
						For iNoLen = 51 to Len(sNoAddNar) Step 50
							sTstr = sTstr & myalign(Mid(Trim(sNoAddNar),iNoLen,50),aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") & VbCrlf
							i = CDbl(i) + 1
						Next
					End IF
				elseif EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "D" then
					sTstr = sTstr & myalign(sNoAddNar,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
					i=51
					IF CDbl(Len(sNoAddNar)) > 50 Then
						For iNoLen = 51 to Len(sNoAddNar) Step 50
							sTstr = sTstr & myalign(Mid(Trim(sNoAddNar),iNoLen,50),aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") & VbCrlf
							i = CDbl(i) + 1
						Next
					End IF
				end if
				'i=51
			else
				if EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "C" then
					sTstr = sTstr & myalign(sAccount ,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
				elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "C" then
					sTstr = sTstr & myalign(sAccount,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
				elseif EntryNode.Attributes.Item(1).nodeValue="C" and sVoucherType = "D" then
					sTstr = sTstr & myalign(sAccount,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
				elseif EntryNode.Attributes.Item(1).nodeValue="D" and sVoucherType = "D" then
					sTstr = sTstr & myalign(sAccount,aiUtilColWidth(1,1)+aiUtilColWidth(1,5),"L") &  string(2," ") & myalign(FormatNumber(sAmount),aiHeadColWidth(0,2),"R") & VbCrlf
				end if
				i=1
			end if



		End IF
	End IF


				objTxt.write sTstr
				sTstr=""



			if iLineNo>26 then
				iPageno =iPageno+ 1
				objtxt.write  sBlankLine
				objTxt.write myalign(" ",aiHeadColWidth(0,1)+5,"L")& myalign("Contd ...",aiHeadColWidth(0,2),"L") & VbCrlf
				ConsPgHeader1 (iPageno)
			end if
		end if
next
'objtxt.write iEnNo

	if CStr(iPageNo)="1" then
		for i=1 to 8-CInt(iEnNo)
			objtxt.write sBlankLine
		next
	else
		for i=1 to 14-CInt(iEnNo)
			objtxt.write sBlankLine
		next
	end if

objtxt.write space(67) & string(13,"-") & VbCrlf
objtxt.write myalign("Total       "& FormatNumber(iSumAmount),aiHeadColWidth(0,0),"R") & VbCrlf
objtxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
	objtxt.write "( "
	objtxt.write myalign(AmountWords(iSumAmount)&" )",aiUtilColWidth(2,0),"L")
objtxt.write " " & VbCrlf
objtxt.write " " & VbCrlf

'if iLineno<20 then CheckNewPage
	CheckNewPage
sFinalText = sFinalText

objTxt.write sFinalText
%>

<%
Function TotalLine()
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,1),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,0),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,2),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,0),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,7),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(1,0),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,7),"L")
	sTstr = sTstr &myalign(" ",aiUtilColWidth(0,1),"L")
	sTstr = sTstr &myalign("----------",aiUtilColWidth(0,1),"L")
	sTstr = sTstr &sBlankLine
End Function

'------------------------End OF MyAlign Function----------------------------

Function ConsPgHeader1(iPageNo)
IF CStr(sVouCherType) = "D" Then
	sPagetitle2 = "CASH RECEIPT VOUCHER"
Else
	sPagetitle2 = "CASH PAYMENT VOUCHER"
End IF
iPageno = iPageNo
iEnNo=0

sQuery = "SELECT ORGUNITDESCRIPTION,ADDRESS1,CITY,STATE,POSTCODE FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID='"&sOrgId&"'"

With Objrs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
Set Objrs.ActiveConnection = nothing

IF Not Objrs.EOF Then
	IF Cint(iPageNo) <> 1 Then
		'objTxt.write FormattPrint("PAGESKIP","")
		objTxt.write " " & VbCrlf
		objTxt.write " " & VbCrlf
		objTxt.write " " & VbCrlf
		objTxt.write " " & VbCrlf
		objTxt.write " " & VbCrlf
	End IF

	objTxt.write myalign(Trim(Objrs(0)),aiHeadColWidth(0,0),"C") & VbCrlf
	objTxt.write myalign(Trim(Objrs(1))&","&Trim(Objrs(2)),aiHeadColWidth(0,0),"C") &" - "& Trim(Objrs(3))& VbCrlf
End IF
Objrs.close

sQuery="Select H.CreatedVouchStatus,V.VoucherNumber,V.CreatedVoucherNo from Acc_T_CreatedVoucherHeader H , Acc_T_VoucherHeader v where H.CreatedTransNo=v.CreatedTransNo and " _
& "right(H.CreatedVouchStatus,2)=04  and H.CreatedTransNo="&iTransNo
'objtxt.write sQuery
With ObjRs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
Set ObjRs.ActiveConnection = nothing
if not ObjRs.EOF then
	AccVoucherNo=ObjRs(1)
	iVoucherNo=objRs(2)
else
	iVoucherNo=Root.attributes.Item(9).nodevalue
end if
ObjRs.Close

objTxt.write myalign(sPagetitle2,aiHeadColWidth(0,0),"C") & VbCrlf

objTxt.write myalign(" ",aiHeadColWidth(1,0)-6,"L")& myalign("Date.   : ",aiHeadColWidth(3,3)+3,"R") & myalign(sVoucherDate,aiHeadColWidth(3,1),"L") & VbCrlf
objTxt.write myalign(" ",aiHeadColWidth(1,0)-6,"L")& myalign("Ref. No : ",aiHeadColWidth(3,3)+3,"R") & myalign(iVoucherNo,aiHeadColWidth(3,1)+10,"L") & VbCrlf
objTxt.write myalign(" ",aiHeadColWidth(1,0)-6,"L")& myalign("Vou. No : ",aiHeadColWidth(3,3)+3,"R") & myalign(AccVoucherNo,aiHeadColWidth(3,1)+10,"L") & VbCrlf
iLineno=16

	if not iPageNo =1 then
		iLineno=13
		if sAccType="P" then
			objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			objTxt.write myalign("Particulars",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
			objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
		else
			objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
			objTxt.write myalign("AccHead",aiHeadColWidth(0,1),"L") & string(5," ") & myalign("Amount(Rs)",aiHeadColWidth(0,2),"L") & VbCrlf
			objTxt.write string(aiHeadColWidth(0,0),"-") & VbCrlf
		end if
	end if
End Function

Function CheckNewPage
		sTstr = sTstr &sBlankLine
		sTstr = sTstr & sBlankLine
		sTstr = sTstr & sBlankLine

		sTstr = sTstr & myalign("PreparedBy ",aiUtilColWidth(1,0),"L")
		sTstr = sTstr & myalign(" ",aiUtilColWidth(0,5),"L")
		sTstr = sTstr & myalign("Cashier ",aiUtilColWidth(0,1),"L")
		sTstr = sTstr & myalign("AO / FM ",aiUtilColWidth(0,1),"L")
		sTstr = sTstr & myalign(" ",aiUtilColWidth(0,6),"L")
		sTstr = sTstr & myalign("M.D/Director ",aiUtilColWidth(0,1),"L")
		sTstr = sTstr & myalign(" ",aiUtilColWidth(0,5),"L")
		sTstr = sTstr & myalign("Receiver's Name",aiUtilColWidth(0,3),"L")

		sTstr = sTstr & sBlankLine
		sTstr = sTstr & sBlankLine
		'sTstr = sTstr & FormattPrint("PAGESKIP","")
		sTstr = sTstr & myalign(" ",aiUtilColWidth(0,0),"L")
		sFinalText=sFinalText & sTstr

		iPageno=iPageno+1
		'sTstr = ConsPgHeader(GetOrganization(),sPagetitle2,iPageNo)
		dTempTotal=0
	end Function

'------------------------End Of User Define Function------------------------
if cint(iRecCount)>=0 then
	Response.Redirect("../../Components/FormattPrintNEW.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_CashPay_View.txt&exitpath=/accounts/reports/CashVouchView_san.asp&frame=_parent")
else
	Response.Clear
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<TITLE>No Records </TITLE>
</HEAD>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<TABLE BORDER=0 ALIGN=CENTER CELLSPACING=0 CELLPADDING=0 NOF=LY>
    <TR >
		<TD height="20">&nbsp;</TD>
		<TD WIDTH=549 ><P ALIGN=CENTER><B>
		<FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT">No Records Found</FONT></B></TD>
	</TR>
	<TR>
        <TD height="20">&nbsp;</TD>
        <TD WIDTH=549 ><P ALIGN=CENTER><B>
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="javascript:window.history.back(1)">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>
<%end if%>