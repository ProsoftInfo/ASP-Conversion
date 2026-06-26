<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PartyOutstanding.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 07, 2010
	'Modified By                :   
	'Modified On                :   
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sUnitID,iCnt,sSql,sTillDate
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo
	Dim iTotalPages,iTotalRecords,iPrevPage,iNextPage
	Dim sSentBy,sSentToVoucher,iSno
	Dim sQuery,sPartyCode,sPartySubType
	Dim sPartyName,sGetVal,sPartyType,dAmtLtThirty,dAmtGtThirty,dAmtGtSixty
	Dim iPartyCode,sOrgnPartyCode,dAmtGtNinety,dTotalBalance,dTotPgBal
	Dim dAmtGtNinetyTot,dAmtGtSixtyTot,dAmtGtThirtyTot,dAmtLtThirtyTot

	Dim Objrs,Objrs1
	
	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")
	
	sUnitID = Session("organizationcode")
	Response.Write "<font color=red>"
	
	Const iPageSize=16
	iPageNo = trim(Request("hPage"))
	if trim(iPageNo) = "" then iPageNo = 1	
	
'	response.write "<font color=red>"& Request.QueryString
	
	iCurrentPage=CInt(Request.Form("hPageSelection"))	
	
	sPartyType = Request("PartyType")
	sPartyCode = Request("hPartyCode")
	sPartySubType = Request("hPartySubType")
	'Response.Write "sPartyType = "& sPartyType 
	
	sTillDate = Trim(Request("hTillDate"))
	If sTillDate = "" Then
		sTillDate = FormatDate(Date())
	End IF
	If sPartyCode = "" Then sPartyCode = "0"
'	Response.Write "<p><font color=red>Data="&sPartyCode & "====="& sPartySubType
		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><PartyType/></XML>
<xml id="PartyData"><Party/></xml>
<XML id="OutStandingData"><Root/></XML>
<XML id="GenReminder"><Root/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<script language="javascript" src="../scripts/VoucherEntryCore.js"></script>
<script language="javascript" src="../scripts/BankVoucher.js"></script>
<script language="javascript" src="../scripts/ReportReminderCompat.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<Script Language=vbscript>

Function SelParty()
	Dim sUnitID
	Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
	sUnitID = document.formname.hUnitNo.value
	
	document.formname.hPartyCode.value=""
    PartyName.innerHTML = ""
    
    Set nodParty = PartyData.documentElement
    
'	set OutValue = showModalDialog("../Reports/PartySelectPopup.asp?orgId="+sUnitID&"&hSelectMode=R",PartyData,"dialogHeight:550px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	'alert(OutValue.xml)
 '   sQuery = OutValue.getAttribute("PassQuery")
 '   if OutValue.getAttribute("Action")="CLOSE" then exit function
 '           	
 '   while OutValue.getAttribute("Action")<>"Done"
 '       set OutValue = showModalDialog("../Reports/PartySelectPopup.asp?"&sQuery,PartyData,"dialogHeight:550px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
 '       sQuery = OutValue.getAttribute("PassQuery")
 '       if OutValue.getAttribute("Action")="CLOSE" then exit function
 '   wend
 '   if OutValue.hasChildNodes() then
  '      For each ndChild in OutValue.childNodes
 '           sPartyCode = sPartyCode & "," & ndChild.getAttribute("ParCode") 
 '           sPartyName = sPartyName & "," & ndChild.getAttribute("ParName")
 '           sPartySubType = sPartySubType & "," & ndChild.getAttribute("ParSubType")  
 '       Next
 '   end if 'if OutValue.hasChildNodes() then
 
        sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sUnitID,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if

        if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sParTy = sParTy &","& ndEntry.getAttribute("RetField3")
		            sPartySubType = sPartySubType &","& ndEntry.getAttribute("RetField4")
		            sPartyCode = sPartyCode &","& ndEntry.getAttribute("RetField1")
		            sPartyName = sPartyName &","& ndEntry.getAttribute("RetField0")
		        end if
            next
        end if
 
 'alert(OutValue.xml)
     
	If trim(sPartyName)<>"" then
        sPartyCode = mid(sPartyCode,2)
        sPartyName = mid(sPartyName,2)
        sPartySubType = Mid(sPartySubType,2)
    end if
       
    document.formname.hPartyCode.value=sPartyCode
    document.formname.hPartySubType.value = sPartySubType
    PartyName.innerHTML = sPartyName 
End Function

Function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.submit()
End Function

Function Validate()
	Dim sStatus
	document.formname.hTillDate.value = document.formname.ctlTillDate.GetDate 	
	document.formname.submit 
End Function

Function CreateRem()
	document.formname.action = "PartyOutstanding.asp"
	document.formname.submit 
End Function

Function ShowCCDetails(obj)
	Dim sTillDate,sOrgId,sOrgName,sTemp,nPartyCode
	
	set Root = OutStandingData.documentElement
	
	sOrgId =document.formname.hUnitNo.Value
	sOrgName=document.formname.hUnitName.Value
	sTillDate= document.formname.hTillDate.value
	
	'nPartyCode = Split(obj.Name,"|")(1)
	'alert(Root.xml)
	
	'sExp = "//Party[@CODE='"& trim(nPartyCode)&"']/DETAILS"
	'set TempNode = Root.selectNodes(sExp)
	'alert(TempNode.length)			
	'If TempNode.length > 0 Then
	'	For iCtr = 0 to TempNode.length - 1
	'		nGetSelectedInvNo = TempNode.item(iCtr).Attributes.getNamedItem("INVOICENO").value
	'	Next
	'End IF
	'alert(nGetSelectedInvNo)
	
	sTemp="Value="&sOrgId&"|"&sOrgName&"|"&sTillDate&"|"&obj.name
	'showModalDialog "PartyOutstandingBreakup.asp?"&sTemp,"A","dialogHeight:640px;dialogWidth:850px;center:Yes;help:No;resizable:No;status:No"
	'window.open "PartyOutstandingBreakup.asp?"&sTemp,"A",""
	
	Set OutDataValue = showModalDialog("PartyOutstandingBreakup.asp?"&sTemp,OutStandingData,"dialogHeight:450px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No")
	
	'alert(OutDataValue.xml)
	set Root = OutStandingData.documentElement
	
	set PartyNode = OutStandingData.createElement("Party")
	Root.appendchild PartyNode

	If OutDataValue.haschildNodes Then
		For Each Node in OutDataValue.childNodes 
			If Node.NodeName = "Party" Then
					
				PartyNode.setAttribute("CODE"),Node.getAttribute("CODE")
				PartyNode.setAttribute("TYPE"),Node.getAttribute("TYPE")
				PartyNode.setAttribute("SUBTYPE"),Node.getAttribute("SUBTYPE")
				PartyNode.setAttribute("NAME"),Node.getAttribute("NAME")
				
				For Each DetNode in Node.childNodes 
					    set NeElem = OutStandingData.CreateElement("DETAILS")
					    NeElem.setAttribute("INVOICENO"),DetNode.getAttribute("INVOICENO")
					    NeElem.setAttribute("AMOUNT"),DetNode.getAttribute("AMOUNT")
					    NeElem.setAttribute("AMOUNTPAIDTILLDATE"),DetNode.getAttribute("AMOUNTPAIDTILLDATE")
					    NeElem.setAttribute("BALANCE"),DetNode.getAttribute("BALANCE")
					    document.formname.hSelNode.value = DetNode.getAttribute("Selected")
					    PartyNode.appendChild NeElem
				Next 
			End IF
		Next 
	End IF 
	
'	alert(Root.xml)
	
	DisplayData()
	
End Function

Function DisplayData()
	Dim Root,sTemp
	Dim SlNo,nTDSAmt,nAdvance,nCNode,nFreight,nAmtPaid
	
	set Root = OutStandingData.documentElement
	
	ClearTable()
	
	If Root.hasChildNodes Then
		For Each Node in Root.childNodes
			If Node.nodeName="Party" Then
				nPartyCode = Node.getAttribute("CODE")
				sPartType = Node.getAttribute("TYPE")
				sPartySubType = Node.getAttribute("SUBTYPE")
				sPartyName = Node.getAttribute("NAME")
				
				For Each DetNode in Node.childNodes 
					
					sPassVal = ""
					SlNo = SlNo + 1
					'SlNo = DetNode.getAttribute("SNO")
					nInvNo = DetNode.getAttribute("INVOICENO")
					'sDocType = DetNode.getAttribute("DOCTYPE")
					'sInvDate = DetNode.getAttribute("DATE")
					'sAccOnDate = DetNode.getAttribute("ACCOUNTEDON")
					nInvAmt = DetNode.getAttribute("AMOUNT")
					nAmtPaidTillDate = DetNode.getAttribute("AMOUNTPAIDTILLDATE")
					nBalanceRec = DetNode.getAttribute("BALANCE")
					
					nInvAmt = FormatNumber(nInvAmt,2,,,0)
					nAmtPaidTillDate = FormatNumber(nAmtPaidTillDate,2,,,0)
					nBalanceRec = FormatNumber(nBalanceRec,2,,,0)
					
					sTemp = nPartyCode & ":" & sPartType & ":" &sPartySubType & ":"& nInvNo
						
					set tRow =document.all.RecTab.InsertRow(document.all.RecTab.rows.length)
					set Cell = trow.InsertCell()
					Cell.innerHTML= SlNo
					Cell.className="ExcelSerial"
					Cell.align="center"
				
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""CheckBox"" name=""ChkBox"& SlNo &""" value="""& Trim(sTemp) &""" class=""Formelem"" >")
					Cell.appendChild(oText)
					'Cell.width = 3
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""sPartyName"" value="""&sPartyName&"""  style=""text-align: Left"" class=""Formelem"" READONLY>")
					Cell.appendChild(oText)
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""InvAmt"" value="""&nBalanceRec&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nTDSAmt"" value="""&nTDSAmt&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nAdvance"" value="""&nAdvance&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nCNode"" value="""&nCNode&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nFreight"" value="""&nFreight&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nAmtPaid"" value="""&nAmtPaidTillDate&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					set oText = document.createElement("<input type=""text"" name=""nAmtOut"" value="""&nBalanceRec&"""  style=""text-align: Right""  size=""15"" class=""FormelemRead"" READONLY>")
					Cell.appendChild(oText)
					Cell.width = 7
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					document.formname.hCount.value = SlNo
					
				Next
				
			End IF	'If Node.nodeName="Party" Then
		Next
	End IF
End Function
Function ClearTable()
	Dim i
	For i = 2 To document.all.RecTab.rows.length - 1
		document.all.RecTab.deleteRow(2)
	Next
End Function

Function CheckSumbit(sCallFrom)
	Dim Root,nNoOfRecords,iCtr,sCheck,sPassValue,blnCheck
	
	set Root = OutStandingData.documentElement
	
	nNoOfRecords = document.formname.hCount.value 
	If nNoOfRecords = "" Then
		alert("Select Party Outstanding Break up")
		Exit Function
	End IF
	blnCheck = False
	
	For iCtr = 1 To nNoOfRecords
		set sCheck = Eval("document.formname.ChkBox"&iCtr)
		If sCheck.checked Then
			nPartyCode = nPartyCode & "," & split(sCheck.value,":")(0)
			sPartType = sPartType & "," & split(sCheck.value,":")(1)
			sPartySubType = sPartySubType & "," & split(sCheck.value,":")(2)
			nInvNo = nInvNo & "," & split(sCheck.value,":")(3)
			blnCheck = True
		End IF
	Next
	
	If Not blnCheck Then
		For iCtr = 1 To nNoOfRecords
			set sCheck =  Eval("document.formname.ChkBox"&iCtr)
			
			nPartyCode = nPartyCode & "," & split(sCheck.value,":")(0)
			sPartType = sPartType & "," & split(sCheck.value,":")(1)
			sPartySubType = sPartySubType & "," & split(sCheck.value,":")(2)
			nInvNo = nInvNo & "," & split(sCheck.value,":")(3)
		Next
	End IF
	
	If nPartyCode <> "" Then nPartyCode = Mid(nPartyCode,2)
	If sPartType <> "" Then sPartType = Mid(sPartType,2)
	If sPartySubType <> "" Then sPartySubType = Mid(sPartySubType,2)
	If nInvNo <> "" Then nInvNo = Mid(nInvNo,2)
	
	
	Dim sSendBy,sCourCompName,sCourTransID,sCourComAddress
	
	If document.formname.radSendBy(0).checked Then
		sSendBy = document.formname.radSendBy(0).value 
	Elseif document.formname.radSendBy(1).checked Then
		sSendBy = document.formname.radSendBy(1).value 
	End If
	
	sCourCompName = Trim(document.formname.txtCouComName.value)
	sCourTransID = Trim(document.formname.txtCouTransID.value)
	sCourComAddress = Trim(document.formname.txtCouComAddress.value)
	

	sPassValue = nPartyCode & ":" & sPartType & ":" &sPartySubType & ":"& nInvNo
		
	    if Trim(sPartType)="CR" then
	        sFileName = "XMLGenReminderPayables.asp"
	    else
	        sFileName = "XMLGenReminder.asp"
	    end if
	    set objHttp = CreateObject("MSXML2.XMLHTTP")
		objHttp.open "GET",sFileName&"?sData="&sPassValue,False
		objHttp.send
		If objHttp.responseXML.xml <> "" Then
			GenReminder.LoadXML objHttp.responseXML.xml
		End IF
		set sRoot = GenReminder.documentElement
		
		If sRoot.hasChildNodes Then
			sRoot.Attributes.getNamedItem("SENDBY").value	 = sSendBy
			sRoot.Attributes.getNamedItem("NAME").value	 = sCourCompName 
			sRoot.Attributes.getNamedItem("ID").value	 = sCourTransID 
			sRoot.Attributes.getNamedItem("ADDRESS").value	 = sCourComAddress 
		End IF
		
		sSelectNode  =  document.formname.hSelNode.value 
		if sRoot.hasChildNodes() then
		    For each ndParty in sRoot.childNodes
		        For each ndDet in ndParty.childNodes
	                if ndDet.nodeName="DETAILS" then
	                    iSNo = ndDet.getAttribute("SNO")
	                    sInvNo = ndDet.getAttribute("INVOICENO")
	                    arrValue = Split(sSelectNode,",")
	                    For iCnt = 0 to UBound(arrValue)
	                        sVal = Split(arrValue(iCnt),":")
	                        if sVal(0)=iSNo and sInvNo = sVal(1) then
	                            ndDet.setAttribute "SELECT","Y"
	                            exit for
	                        end if
	                    Next
	                end if 'if ndParty.nodeName="DETAILS" then
	            Next
		    Next
		end if
		
		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "POST","GenReminderInsert.asp",False
		objHttp.send GenReminder.XMLDocument
		
		If objHttp.responseText <> "" Then
		   arrTemp = Split(objhttp.responseText,"@")
		    if arrTemp(0)<>"" then
			    alert(arrTemp(0))
			else
			    Set OutDataValue = showModalDialog("PartyOutstandingPrevReminder.asp?"&arrTemp(1)&"&PassValue="&sPassValue,OutStandingData,"dialogHeight:450px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No")
			    	if Trim(sPartType)="CR" then
			            window.location.href = "PAYMENTREMINDERS.ASP"
			        else
			            window.location.href = "OverdueReminders.asp" 
			        end if 
		    end if
		End IF
End Function

</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
	<form method="POST" name="formname" action="" >
	<input type="hidden" name="PartyType" value="<%=sPartyType%>">
	<input type=hidden name="hUnitNo" value="<%=sUnitID%>">
	<input type=hidden name="hUnitName" value="<%=session("orgShortName")%>">
	<input type=hidden name="hTillDate" value="<%=sTillDate%>">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<Input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
	<Input type="hidden" name="hPartySubType" value="<%=sPartySubType%>">
	<Input type="Hidden" name="hCount" value="">
	<input type="Hidden" name="hSelNode" value="">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle">
				<p align="center">Party Outstanding
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')"  itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="right" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%">
<div id="idUnprocessed" style="display: none">
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<tr>
	<td class="FieldCellsub">Party</td>
	<td class="FieldcellSub"> 
		<span id="PartyName" class="Dataonly"></span>
		<a href="#"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" onclick="SelParty()"></a>
	</td>
</tr>
<tr>
	<td class="FieldCellsub">Till Date</td>
	<td class="FieldCellsub">
		<%Response.Write InsertDatePicker("ctlTillDate")%>
	</td>
</tr>

<tr>
<td class="FieldCell"></td>
<td class="FieldCell" >
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
</td>
<td class="FieldCell" >
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<!--<div class="frmBody" id="frm4" style="width: 585; height:140;">-->
<!--<div class="frmBody" id="frm4" style="height:270;">-->
	
	<!--<div class="frmBody" id="frm4" style="height:130;">-->
		<table cellspacing="1" class="ExcelTable" width="100%" >
		<tr>
			<TD align="center" class="ExcelHeaderCell" rowspan="2"><P>S.No.</TD>
			<TD class="ExcelHeaderCell" align="center" rowspan="2"><P>Party code</TD>
			<TD COLSPAN=5 class="ExcelHeaderCell" align="center"><P>Outstanding amount for no of days</TD>
		</tr>

		<tr>
			<TD class="ExcelHeaderCell" align="center"><P>&lt; 30</TD>
			<TD class="ExcelHeaderCell" align="center"><P>30 - 60</TD>
			<TD class="ExcelHeaderCell" align="center"><P>61 - 90</TD>
			<TD class="ExcelHeaderCell" align="center"><P>&gt; 90</TD>
			<TD class="ExcelHeaderCell" align="center"><P>Total</TD>
		</tr>

		<%	

		sQuery="select distinct(partycode),orgnpartycode,partyName,PartySubType from vwOrgparty where "&_
				" OUDefinitionID='" & sUnitID & "' and partytype='" & sPartyType & "'"
				
		If cint(trim(sPartyCode)) <> "0" then 
			sQuery= sQuery &" and PartyCode = "&sPartyCode&" " 		
		End IF
		If trim(sPartySubType) <> ""  then 
			sQuery= sQuery &" and PartySubType = "&sPartySubType&" " 		
		End IF
		' Response.write sQuery
		 'Response.End
			 
		with objRs
			.CursorLocation =3
			.CursorType =3
			.ActiveConnection=con
			.Source =sQuery 
			.Open 
		End with

		If not objRs.EOF Then
			set iPartyCode =objRs(0)
			set sOrgnPartyCode=objRs(1)	
			Set sPartyName =objRS(2)
			set sPartySubType =objRs(3)
		End if

		dAmtLtThirtyTot=0
		dAmtGtThirtyTot=0
		dAmtGtSixtyTot=0
		dTotPgBal=0

		While not objRs.EOF 
		dTotalBalance=0
		if Trim(sPartyType)="DR" then

			sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=0"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<30 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
		'	Response.Write "<font color=red>< 30 = "&sQuery &"<BR>" 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtLtThirty=objRs1(0)
			dAmtLtThirtyTot=cdbl(dAmtLtThirtyTot)+cdbl(dAmtLtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtLtThirty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=30"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<60 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "> 30 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtThirty=objRs1(0)
			dAmtGtThirtyTot=cdbl(dAmtGtThirtyTot)+cdbl(dAmtGtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtThirty)
			objRs1.Close 

		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=60"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<90 and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<font color=red>30  60 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtSixty=objRs1(0)
			dAmtGtSixtyTot=cdbl(dAmtGtSixtyTot)+cdbl(dAmtGtSixty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtSixty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountReceivable),0)-Isnull(Sum(AmountReceived),0) From Acc_T_Receivables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=90"&_
				" and AmountReceivable>AmountReceived"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<BR><BR>iPartyCode  > 90 = "&sQuery 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtNinety=objRs1(0)
			dAmtGtNinetyTot=cdbl(dAmtGtNinetyTot)+cdbl(dAmtGtNinety)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtNinety)
			objRs1.Close 
			dTotPgBal=cdbl(dtotPgBal)+cdbl(dTotalBalance)
	end if 'if Trim(sPartyType)="DR" then

			'sQuery = "Select isNull(Sum(P.AdvanceReceived - isNull(P.AdvanceAdjusted,0)),0) "&_
			'			 "From Acc_T_AdvancePayments P, Acc_T_VoucherHeader H  Where P.PartyType = '"&sPartyType&"' and  "&_
			'			 "P.PartySubType = "&sPartySubType&" and P.PartyCode = "&iPartyCode&" and P.OUDefinitionID = '"&sOrgId&"' and "&_
			'			 "H.TransactionNumber = P.TransactionNumber and P.AdvanceReceived is Not Null "
			
			'with objRs1
			'	.CursorLocation =3
			'	.CursorType =3
			'	.ActiveConnection =con
			'	.Source =sQuery
			'	.Open 
			'End with
			'set objRs1.ActiveConnection =nothing
			
			'IF Not objRs1.EOF Then
			'	dTotalBalance=cdbl(dTotalBalance)+cdbl(objRs1(0))
			'End IF
			'objRs1.Close
			
    if Trim(sPartyType)="CR" then
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=0"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<30 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
		'	Response.Write "<font color=red>< 30 = "&sQuery &"<BR>" 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtLtThirty=objRs1(0)
			dAmtLtThirtyTot=cdbl(dAmtLtThirtyTot)+cdbl(dAmtLtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtLtThirty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=30"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<60 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "> 30 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtThirty=objRs1(0)
			dAmtGtThirtyTot=cdbl(dAmtGtThirtyTot)+cdbl(dAmtGtThirty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtThirty)
			objRs1.Close 

		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "'and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=60"&_
				" and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))<90 and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<font color=red>30  60 = "&sQuery &"<BR>"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtSixty=objRs1(0)
			dAmtGtSixtyTot=cdbl(dAmtGtSixtyTot)+cdbl(dAmtGtSixty)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtSixty)
			objRs1.Close 
			
		sQuery="Select Isnull(Sum(AmountPayable),0)-Isnull(Sum(AmountPaid),0) From Acc_T_Payables"&_
				" Where OuDefinitionId='" & sUnitID & "' and DateDiff(day,VoucherDate,convert(datetime,'"& sTillDate &"',103))>=90"&_
				" and AmountPayable>AmountPaid"&_ 
				" and partycode=" & iPartyCode &" and partysubtype=" & sPartySubType &_
				" and PartyType='"& sPartyType & "'" 
			'Response.Write "<BR><BR>iPartyCode  > 90 = "&sQuery 
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.ActiveConnection =con
				.Source =sQuery
				.Open 
			End with
			set objRs1.ActiveConnection =nothing
			dAmtGtNinety=objRs1(0)
			dAmtGtNinetyTot=cdbl(dAmtGtNinetyTot)+cdbl(dAmtGtNinety)
			dTotalBalance=cdbl(dTotalBalance)+cdbl(dAmtGtNinety)
			objRs1.Close 
			dTotPgBal=cdbl(dtotPgBal)+cdbl(dTotalBalance)

	end if 'if Trim(sPartyType)="CR" then		
			
			
						 
		If dTotalBalance<>0 Then
		iSNo=iSNo+1
		%>
		<tr>
			<TD align="center" class="ExcelSerial"><P><%=iSNo%></TD>
			<TD class="ExcelDisplayCell" align="left">
				<A Name="<%=sPartySubType%>|<%=iPartyCode%>|<%=iPartyCode%>|<%=sPartyType%>"
				HRef="#" class="ExcelDisplayLink" onclick="ShowCCDetails(this)" ALT="View Paybales Details"><%=sPartyName%></a>
			</TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtLtThirty,2,,,0)%> </TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtThirty,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtSixty,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dAmtGtNinety,2,,,0)%></TD>
			<TD align="right" class="ExcelDisplayCell"><P><%=FormatNumber(dTotalBalance,2,,,0)%></TD>
		</tr>
		<%
		End if
		objRs.MoveNext 
		wend 
		%>
		<TD Colspan=2 align="right" class="ExcelDisplayCell">
		<P><b>Total</TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtLtThirtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtThirtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtSixtyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dAmtGtNinetyTot,2,,,0)%></TD>
		<TD align="right" class="ExcelDisplayCell">
		<P><b><%=FormatNumber(dTotPgBal,2,,,0)%></TD>

		</table>
	</Div>
	
	<table>
		<tr>
			<td align="center" class="MiddlePack" colspan="3"></td>
		</tr>
	</Table>

	<!--<div class="frmBody" id="frm4" style=" height:130;">-->
		<table cellspacing="1" class="ExcelTable" width="100%" ID="RecTab">
			<tr>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>S.No.</TD>
				<td class="ExcelHeaderCell" align="center" Rowspan="2">
					<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15" height="15" onclick="">
					</a>
				</td>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>Party</TD>
				<TD class="ExcelHeaderCell" rowspan="2" align="center"><P>Invoice <br>Amount</TD>
				<TD class="ExcelHeaderCell" colspan="4" align="center"><P>Deductions</TD>
				<TD class="ExcelHeaderCell" colspan="2" align="center"><P>Amount</TD>
			</tr>
			<tr>
				<TD class="ExcelHeaderCell" align="center"><P>TDS</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Advance</TD>
				<TD class="ExcelHeaderCell" align="center"><P>C.Note</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Freight</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Paid</TD>
				<TD class="ExcelHeaderCell" align="center"><P>Outstanding</TD>
			</tr>
		</Table>
		
	<!--</div>-->
	
	<!--<div class="frmBody" id="frm4" style=" width:570 ;height:90;">-->
		<Table>
			<tr>
				<td class="FieldCellSub">&nbsp;</td>
				</td>
			</tr>
		</Table>
	
		<Table cellspacing="0" class="BodyTable" width="100%">	
			<tr>
				<td class="FieldCellSub">To Send By</td>
				<td class="FieldCell">
					<Input type="Radio" name="radSendBy" Value="C">Courier
					<Input type="Radio" name="radSendBy" Value="E">E-Mail
				</td>
			</tr>
			<tr>
				<td class="FieldCell">Courier Company Name</td>
				<td class="FieldCell">
					<Input type="Text" name="txtCouComName" value="" class="FormElem">
				</td>
				<td class="FieldCell">Courier TransactionID</td>
				<td class="FieldCell">
					<Input type="Text" name="txtCouTransID" value="" class="FormElem">
				</td>
			</tr>
			<tr>
				<td class="FieldCell">Address</td>
				<td class="FieldCell" colspan="2">
					<Textarea type="Text" name="txtCouComAddress" value="" class="FormElem" cols="40"></Textarea>
				</td>
			</tr>
		</Table>
		
	<!--</div>

</div>-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>
<input type=hidden name="hCnt" value=<%=iSno-1%>>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="middle" class="ActionCell">
<p align="center">
	<tr>
		
		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="middle" class="ActionCell">
                        <p align="center">
                         <!--<input type="button" value="Preview Reminder" class="ActionButtonX"  id="button1" name=button1 OnClick="CheckSumbit('P')" >-->
                         <input type="button" value="Generate Reminder" class="ActionButtonX"  id="button2" name=button2 OnClick="CheckSumbit('G')" >
					</td>
				</tr>
			</table>
		</td>
		
    </tr>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>
<tr>
<td align="center" class="BottomPack" colspan="3">
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
</body>
</html>
