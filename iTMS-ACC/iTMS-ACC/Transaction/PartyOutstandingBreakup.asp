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
	'Program Name				:	PartyOutstandingBreakup.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	N.Rajkumar
	'Created On					:	19th June 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	07th April 2011
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
<%
dim objRs,objRs2,objRs3,sQuery,iPageNo,saTemp,dTotalAmt,dTotalBal,dTotAmtPaid
dim sOrgId,sOrgName,sTillDate,sPartySubType,sSubTypeName,iPartyCode,sVouDate
dim sPartyName,sGetVal,sPartyType,iInvoiceNo,sInvoiceDate,dAmount,dAmtPaid,dBalAmt,iSNo,iNoOfDays
Dim sTransType,sTransName,iDueDays,iCrTransNo,sTempVal
iPageNo=1
iSNo=0
set objRs  = server.CreateObject("adodb.recordset")
set objRs2  = server.CreateObject("adodb.recordset")
set objRs3  = server.CreateObject("adodb.recordset")

'----------- To Get The Values From the Selection Page ----------------

sGetVal=Request.QueryString("Value")
'Response.Write sGetVal
'Response.End
saTemp=split(sGetVal,"|")
sOrgId= saTemp(0)
sOrgName=saTemp(1)
sTillDate=saTemp(2)
sPartySubType =saTemp(3)
iPartyCode=saTemp(4)
sPartyType=saTemp(6)

IF CStr(iPartyCode) <> "0" Then
	sPartyName = GetPartyName(iPartyCode)
Else
	sPartyName = "All Parties"
End IF
sTempVal = iPartyCode & ":" & sPartyType & ":" & sPartySubType & ":" & sPartyName 
'To Display Organizations Full Description

sQuery="Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID='" & sOrgId & "'"

with objRs
	.CursorLocation =3
	.CursorType=3
	.ActiveConnection =con
	.Source =sQuery
	.Open 
End with 
set objrs.ActiveConnection =nothing
If not objRs.EOF then
	sOrgName =objRs(0)
End if
objRs.Close 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receivables View</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/ReportsBody.css" TYPE="text/css">
<XML ID="ReceivableData"><Reminder/></XML>
<XML ID="OutData"><Reminder/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/ReportReminderCompat.js"></script>
<Script Language="Vbs">
Dim objTemp ,nCheckInvNo
set objTemp = window.dialogArguments
'alert(objTemp.xml)
Function CloseWindow()
	window.close()
End Function

Function InvoicePopUp(iRcvbleNo,dtBillDate,iAmtPad,iBaAmt,sTransType)
	'alert iBillNo&","&trim(dtBillDate)&","&iAmtPad&","&iBaAmt
	sTemp= iRcvbleNo&":"&trim(dtBillDate)&":"&iAmtPad&":"&iBaAmt&":"&sTransType
	showModalDialog "InvoicePopUp.asp?sTemp="&sTemp,"","dialogHeight:150px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No"
End Function

Function ShowDetails(iTransNo,iNoOfDays)
	Dim sValue 
	'MsgBox iNoOfDays
	IF CDbl(iNoOfDays) > 0 Then
		sValue = iTransNo&"-"&iNoOfDays
		'sRetVal = open("ReceivablesPenalty.asp?TransNo="&sValue,"A","height=190,width=200,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0,resizable=No" ) 
		showModalDialog "ReceivablesPenalty.asp?TransNo="&sValue,"A","dialogHeight:300px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No"
	End IF
End Function

Function GetXML(sValue,sPartyType)
	Dim objHttp,Root
	if Trim(sPartyType)="CR" then
	    sFileName = "XMLGetOutstandingPayables.asp"
	else
	    sFileName = "XMLGetOutstandingReceivables.asp"
	end if
	set objHttp = CreateObject("MSXML2.XMLHTTP")
	objHttp.open "GET",sFileName&"?Value="&sValue,False
	objHttp.send
	If objHttp.responseXML.xml <> "" Then
		ReceivableData.loadXML objHttp.responseXML.xml
		set Root = ReceivableData.documentElement
	End IF
	DisplayData()
End Function

Function ClearTable()
	Dim i
	alert(document.all.RecTab.rows.length )
	For i = 1 To document.all.RecTab.rows.length 
		document.all.RecTab.deleteRow(i)
	Next
End Function

Function DisplayData()
	Dim Root,Node,DetNode
	Dim tRow,Cell,SlNo,nInvNo,sPassVal,nCheckInvNo
	set Root = ReceivableData.documentElement
	
	set objTempRoot = objTemp.documentElement
	
	If Root.hasChildNodes Then
		For Each Node in Root.childNodes
			If Node.nodeName="Party" Then
				
				nPartyCode = Node.getAttribute("CODE")
				sPartType = Node.getAttribute("TYPE")
				sPartySubType = Node.getAttribute("SUBTYPE")
				sPartyName = Node.getAttribute("NAME")
				
				For Each DetNode in Node.childNodes 
					
					sPassVal = ""
					
					SlNo = DetNode.getAttribute("SNO")
					nInvNo = DetNode.getAttribute("INVOICENO")
					sDocType = DetNode.getAttribute("DOCTYPE")
					sInvDate = DetNode.getAttribute("DATE")
					sAccOnDate = DetNode.getAttribute("ACCOUNTEDON")
					nInvAmt = DetNode.getAttribute("AMOUNT")
					nAmtPaidTillDate = DetNode.getAttribute("AMOUNTPAIDTILLDATE")
					nBalanceRec = DetNode.getAttribute("BALANCE")
					
					'sExp = "//Party[@CODE='"& trim(document.formname.hPartycode.value)&"']/DETAILS[@INVOICENO='"& trim(nInvNo)&"']"
					sExp = "//Party[@CODE='"& trim(document.formname.hPartycode.value)&"']/DETAILS"
					set TempNode = objTempRoot.selectNodes(sExp)
	
					'alert(TempNode.length)
	
					If TempNode.length > 0 Then
						For iCount = 0 To TempNode.length - 1
							nCheckInvNo = TempNode.item(icount).Attributes.getNamedItem("INVOICENO").value
						Next	
					End IF
					'alert(nCheckInvNo)
					
					'sPassVal =nPartyCode & ":" & sPartType & ":" & nInvNo&":"&sDocType&":"&sInvDate&":"&sAccOnDate&":"&nInvAmt&":"&nAmtPaidTillDate&":"&nBalanceRec&":"&sPartyName
					
					set tRow =document.all.RecTab.InsertRow(document.all.RecTab.rows.length)
					set Cell = trow.InsertCell()
					Cell.innerHTML= SlNo
					'Cell.width = 3
					Cell.className="ExcelSerial"
					Cell.align="center"
				
					set Cell=trow.insertCell()
					If nCheckInvNo <> "" Then				'4,5
						nInvNoArr = split(nCheckInvNo,",")	
						For iCtr = LBound(nInvNoArr) To UBound(nInvNoArr)
							nCheckInvNo = ""
							nCheckInvNo = nInvNoArr(iCtr) '4
							
							If Trim(nCheckInvNo) = trim(nInvNo) Then
								set oText = document.createElement("<input type=""CheckBox"" name=""ChkBox"& SlNo &""" value="""&SlNo&":"&nInvNo&""" class=""Formelem"" CHECKED>")
								Exit For 
							Else
								set oText = document.createElement("<input type=""CheckBox"" name=""ChkBox"& SlNo &""" value="""&SlNo&":"&nInvNo&""" class=""Formelem"" >")
							End IF
						Next
					Else
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkBox"& SlNo &""" value="""&SlNo&":"&nInvNo&""" class=""Formelem"" >")
					End If
					Cell.appendChild(oText)
					'Cell.width = 3
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""DocType"" value="""&sDocType&"""  style=""text-align: Center""  Size=""6""class=""Formelem"" READONLY>")
					Cell.appendChild(oText)
					'Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""InvNo"" value="""&nInvNo&"""  style=""text-align: Center""  Size=""5""class=""Formelem"" READONLY>")
					Cell.appendChild(oText)
					'Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""Date"" value="""&sInvDate&"""  style=""text-align: Center"" size=""11"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""AccOn"" value="""&sAccOnDate&"""  style=""text-align: Center"" size=""11""  class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""InvAmt"" value="""&nInvAmt&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""AmtPaidTill"" value="""&nAmtPaidTillDate&"""  style=""text-align: Right"" size=""15""  class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""Balance"" value="""&nBalanceRec&"""  style=""text-align: Right"" size=""15""  class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""NoOfDaysOut"" value="""&DetNode.getAttribute("NOOFDAYSOUT")&"""  style=""text-align: Center"" size=""7""  class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""NoOfDaysOver"" value="""&DetNode.getAttribute("NOOFDAYSOVER")&"""  style=""text-align: Center"" size=""7""  class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					document.formname.hCnt.value = SlNo
				Next 
			End IF	'If Node.nodeName="Party" Then
			If Node.NodeName = "TOTAL" Then
				
				set tRow =document.all.RecTab.InsertRow(document.all.RecTab.rows.length)
				
				set Cell=trow.insertCell()
				Cell.innerHTML= "Total Outstanding Transaction Amount"
				Cell.className="ExcelHeaderCell"
				Cell.align="center"
				Cell.colspan="6"
				cell.align = "Right"
				
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""InvAmt"" value="""&Node.getAttribute("AMOUNT")&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"
				
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""AmtPaidTill"" value="""&Node.getAttribute("PAID")&"""  style=""text-align: Right"" size=""15""  class=""FormelemRead"" READONLY>")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"
				
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""Balance"" value="""&Node.getAttribute("RECEIVABLE")&"""  style=""text-align: Right"" size=""15""  class=""FormelemRead"" READONLY>")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"
				
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""NoOfDaysOut"" value=""""""  style=""text-align: Center"" size=""7""  class=""FormelemRead"" READONLY>")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"
				
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""NoOfDaysOver"" value=""""""  style=""text-align: Center"" size=""7""  class=""FormelemRead"" READONLY>")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"
				
			End IF
		Next	
	End If	'If Root.hasChildNodes Then
	
End Function

Set oRoot = OutData.documentElement

Function AddToReminder()
	Dim nNoOfRec,iCtr,sCheck,sTemp,blnCheck
	
	set Root = ReceivableData.documentElement
	sTemp = Split(document.formname.hTempVal.value,":")
	set oRoot = OutData.documentElement
	set objTempRoot = objTemp.documentElement
	
	sFlag =True
	sExp = "//Party[@CODE='"& trim(sTemp(0))&"']"
	set TempNode = objTempRoot.selectNodes(sExp)
	'alert(TempNode.length)
	If TempNode.length > 0 Then
		set PartyNode = TempNode.item(0)
		sFlag =False
	End IF
	
	sExp = "//Party[@CODE='"& trim(sTemp(0))&"']/DETAILS"
	set TempNode = objTempRoot.selectNodes(sExp)
	If TempNode.length > 0 Then
		TempNode.removeall
	End IF
	
	If sFlag Then
		
		set PartyNode = OutData.createElement("Party")
		oRoot.appendchild PartyNode

		PartyNode.setAttribute("CODE"),sTemp(0)
		PartyNode.setAttribute("TYPE"),sTemp(1)
		PartyNode.setAttribute("SUBTYPE"),sTemp(2)
		PartyNode.setAttribute("NAME"),sTemp(3)
	
	End IF
	nNoOfRec = document.formname.hCnt.value 
	
	blnCheck = "N"
	
	If Root.haschildNodes Then
		For Each Node in Root.childNodes 
			If Node.NodeName = "Party" Then
					
					nPartyCode = Node.getAttribute("CODE")
					
				For Each DetNode in Node.childNodes 
					
					nSno   = DetNode.getAttribute("SNO")
					nInvNo = DetNode.getAttribute("INVOICENO")
					
					For iCtr = 1 to nNoOfRec
						set sCheck = Eval("document.formname.ChkBox"&iCtr)
						If sCheck.checked Then
							sVal = Split(trim(sCheck.value),":")
							If nSno = trim(sVal(0)) and nInvNo = trim(sVal(1)) Then
							sSelectedNode = sSelectedNode &","& sCheck.value
							    DetNode.setAttribute "SELECTION","Y"
							
								blnCheck = "Y"
								
								'sExp = "//Party[@CODE='"& trim(nPartyCode)&"']/DETAILS[@INVOICENO='"& trim(nInvNo)&"']"
								'set TempNode = objTempRoot.selectNodes(sExp)
	
								'alert(TempNode.length)
								
								'If TempNode.length > 0 Then
								'	TempNode.removeall
								'End IF
								
								'sDocType = DetNode.getAttribute("DOCTYPE")
								nInvoiceNo = nInvoiceNo & "," & DetNode.getAttribute("INVOICENO")
								nInvAmount = cdbl(nInvAmount) + cdbl(DetNode.getAttribute("AMOUNT"))
								nAmtPaid = cdbl(nAmtPaid)+ cdbl(DetNode.getAttribute("AMOUNTPAIDTILLDATE"))
								nBalAmount = cdbl(nBalAmount) + cdbl(DetNode.getAttribute("BALANCE"))
								
							End IF
						End IF
					Next 
					
					
				Next 
			End IF
		Next 
	End IF 
	if Trim(sSelectedNode)<>"" then
	    sSelectedNode = Mid(sSelectedNode,2)
	end if
	If blnCheck = "Y" Then
		set NeElem = OutData.CreateElement("DETAILS")
		'NeElem.setAttribute("SNO"),nSno
		'NeElem.setAttribute("DOCTYPE"),DetNode.getAttribute("DOCTYPE")
		NeElem.setAttribute("INVOICENO"),mid(nInvoiceNo,2)
		NeElem.setAttribute("AMOUNT"),nInvAmount
		NeElem.setAttribute("AMOUNTPAIDTILLDATE"),nAmtPaid
		NeElem.setAttribute("BALANCE"),nBalAmount
		NeElem.setAttribute("Selected"),sSelectedNode
		PartyNode.appendChild NeElem
	End IF
	window_onunload()
End Function


Function window_onunload()
	set window.returnvalue = OutData.documentElement
	window.close 
End Function

</Script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="GetXML('<%=sGetVal%>','<%=sPartyType%>')">

<form method="POST" name="formname" action="">	
	<Input type="Hidden" name="hCnt" value="">
	<Input type="Hidden" name="hGetVal" value="<%=sGetVal%>">
	<Input type="Hidden" name="hTempVal" value="<%=sTempVal%>">
	<Input type="Hidden" name="hPartycode" value="<%=iPartyCode%>">
	
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Outstanding  Receivables  as On <%=sTillDate%> From <%=sPartyName%>
		
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">

								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                <!--<div class="frmBody" id="frm2" style="width: 755; height:395;">-->
                                <div class="frmBody" id="frm2" style="width: 695; height:310;">
                                	<TABLE BORDER="0" CELLSPACING=1 CELLPADDING=0 WIDTH=100% class="ExcelTable" ID="RecTab">
										<tr>
											<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15"><P>S.No.</TD>
											<td class="ExcelHeaderCell" align="center" Rowspan="2">
												<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="">
												</a>
											</td>
											<TD class="ExcelHeaderCell" rowspan="2" align="center" width="15"><P>Doc Type.</TD>
											<TD COLSPAN=4 class="ExcelHeaderCell" align="center" ><P>Party Invoice </TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Amount paid <br>till date</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Balance Receivable</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80"><P>No of <br>days<br>Outstanding</TD>
											<TD class="ExcelHeaderCell" align="center" rowspan="2" width="80"><P>No of days<br>Over Due</TD>
										</tr>

										<tr>
											<TD class="ExcelHeaderCell" align="center" width="140"><P>No</TD>
											<TD class="ExcelHeaderCell" align="center" width="80"><P>Date</TD>
											<TD class="ExcelHeaderCell" align="center" width="80"><P>Accounted On</TD>
											<TD class="ExcelHeaderCell" align="center" width="100"><P>Amount</TD>
										</tr>

									</TABLE>
                                </div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                
                                                <p align="center"> 
													<input type="button" value="Add To Reminder" OnClick="AddToReminder()" class="ActionButtonX" tabindex="3"  id=button1 name=button1>
													<input type="button" value="Close" OnClick="CloseWindow()" class="ActionButton" tabindex="3"  id=button2 name=button2>
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
	Function GetDueDays(iCrTrNo,iOutStdDays,iAccTrNo,sBillDate)
		Dim sQuery,sObjDueRs,iOthAppNo,iPayTerms,iPayNoDays,iPayCount,iTotalDueDay
		Dim iDurPer,iDurDays,sPayTillDate,sObjPayRs,iPayRecNo,iParPayAmt,iPayInvAmt
		Dim iPayAmtToCome 
		Set sObjDueRs = Server.CreateObject("ADODB.RecordSet")
		Set sObjPayRs = Server.CreateObject("ADODB.RecordSet")
		
		sQuery = "Select isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTrNo
		'Response.Write sQuery
		sObjDueRs.Open sQuery,Con
		IF Not sObjDueRs.EOF Then
			iOthAppNo = sObjDueRs(0)
		End IF
		sObjDueRs.Close
		
		IF CStr(iOthAppNo) = "0" Then
			GetDueDays = 0
		Else
			sQuery = "Select InvPymtTerms From Sal_T_InvoiceHeader Where SaleTransactionNo = " & iOthAppNo
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayTerms = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			'iPayTerms = 2
			
			sQuery = "Select Count(1) From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
			
			sObjDueRs.Open sQuery,Con
			IF Not sObjDueRs.EOF Then
				iPayCount = sObjDueRs(0)
			End IF
			sObjDueRs.Close
			
			
			
			IF iPayCount = 1 Then
				sQuery = "Select DueDay From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayNoDays = sObjDueRs(0)
				End IF
				sObjDueRs.Close
				
				'iPayNoDays = 300
				
				'Response.Write iAccTrNo &"   " & iPayNoDays &"<br>"
				iTotalDueDay = CDbl(iOutStdDays) - CDbl(iPayNoDays)
			Else
				iPayAmtToCome = 0
				sQuery = "Select ReceivableNumber,AmountReceivable From Acc_T_Receivables Where TransactionNumber = "&iAccTrNo
				sObjDueRs.Open sQuery,Con
				IF Not sObjDueRs.EOF Then
					iPayRecNo = sObjDueRs(0)
					iPayInvAmt = sObjDueRs(1)
				End IF
				sObjDueRs.Close
				
				sQuery = "Select DueDay,DuePercent From APP_M_PaymentTermsDetails Where PaymentTermsNo = "&iPayTerms
				sObjDueRs.Open sQuery,Con
				Do While Not sObjDueRs.EOF
					iDueDays = sObjDueRs(0)
					iDurPer = sObjDueRs(1)
					
					'Response.Write iPayInvAmt &"  " & iDurPer &"<br>"
					
					iPayAmtToCome = Cdbl((CDbl(iPayInvAmt) * CDbl(iDurPer))/100) + CDbl(iPayAmtToCome)
					
					
					sQuery = "Select Top 1 Convert(Varchar,DateAdd(day,"&iDueDays&",Convert(Datetime,'"&sBillDate&"',103)),103) From Acc_T_RcvblAdjustmentDetails  "
					'Response.Write sQuery
					sObjPayRs.Open sQuery,Con
					
					IF Not sObjPayRs.EOF Then
						sPayTillDate = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					sQuery = "Select Sum(AmountReceived) From Acc_T_RcvblAdjustmentDetails Where ReceivableNumber = "&iPayRecNo&" and "&_
							 "Convert(Datetime,ReceivedOn,103) <=  Convert(Datetime,'"&sPayTillDate&"',103)  "
							
					'Response.Write sQuery &"<br>"
					sObjPayRs.Open sQuery,Con 
					IF sObjPayRs.EOF Then
						iParPayAmt = sObjPayRs(0)
					End IF
					sObjPayRs.Close
					
					'iParPayAmt = 119000
					'Response.Write iParPayAmt &"  " & iPayAmtToCome &"<br>"
					
					IF CDbl(iParPayAmt) >= CDbl(iPayAmtToCome) Then
						iTotalDueDay = 0
					Else
						iTotalDueDay = CDbl(iOutStdDays) - CDbl(iDueDays)
					End IF
					
					'Response.Write iTotalDueDay &"<br>"
					
					
							 
					sObjDueRs.MoveNext
				Loop
				sObjDueRs.Close
			End IF
		End IF
		
		IF iTotalDueDay <0 Then
			iTotalDueDay = 0
		End IF
		
		IF CStr(iTotalDueDay) = "" Then
			iTotalDueDay = 0
		End IF
		
		GetDueDays = iTotalDueDay
	End Function
%>
