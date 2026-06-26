<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtLoanEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  1, 2003
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
<!-- #include File="../../include/IncludeDatePicker.asp" -->

<%
dim sOrgId,sOrgName,sAccCode,sAccName,sRequestType

'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("hUnitName")
sOrgId = session("organizationcode")
sOrgName = session("orgshortname")

sAccCode=Request.Form("hAccountCode")
sAccName=Request.Form("hAccountName")
sRequestType=Request.Form("hReqTypeS")
'Response.Write "<p>sRequestType="&sRequestType

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML id="AccHeadData">
<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT language="vbscript">
function selAccountHead(objAcc)
dim sTemp
if objAcc.selectedIndex >0 then
	if objAcc.selectedIndex >0 then
		'If selected account Head is Party type
		sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
		showPartyHead document.formname.hUnitId.value ,sTemp
	else
		showGLHead(document.formname.hUnitId.value)
	End if 'End of select Account Head Type check GL or PARTY
else
	document.formname.hAccountCode.value=""
	document.formname.hAccountName.value=""
	document.formname.txtPayTo.value=""

End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead----------------------
function showGLHead(sOrgId)
Dim arrTemp,sRetVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

     '   OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=00&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
     '   arrTemp = split(OutValue,":")
     '   while UBound(arrTemp) = 0
	 '       OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	 '       arrTemp = split(OutValue,":")
     '   wend

     '   sRetVal = OutValue
     '   if UBound(arrTemp) <= 1 then exit function
     
            sTempValWindowSize = GetWindowSizeForPopup("5")
            sArrTempValWindowSize = split(sTempValWindowSize,":")
            sProgramName = sArrTempValWindowSize(0)
            sPopupHeight = sArrTempValWindowSize(1)
            sPopupWidth = sArrTempValWindowSize(2)
    		
		    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId=00&BookNo="&iBookNo,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	        sAct = UCase(trim(OutValue.getAttribute("Action")))
	        sQuery = trim(OutValue.getAttribute("PassQuery"))
	        if ucase(trim(sAct)) <> "CLOSE" then
		        do while sAct <> "DONE"
			        set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			        sAct = UCase(trim(OutValue.getAttribute("Action")))
			        if ucase(Trim(sAct)) = "CLOSE" then exit do
			        sQuery = trim(OutValue.getAttribute("PassQuery"))
		        loop
	        end if
	        
            if OutValue.hasChildNodes() then
                for each ndEntry in OutValue.childNodes
                    if ndEntry.nodeName="Entry" then
                        sRetVal = ndEntry.getAttribute("RetField0")&":"&ndEntry.getAttribute("RetField1")&":"&ndEntry.getAttribute("RetField2")&":"&ndEntry.getAttribute("RetField3")&":"&ndEntry.getAttribute("RetField4")&":"&ndEntry.getAttribute("RetField5")&":"&ndEntry.getAttribute("RetField6")
                    end if
                next
            end if
     
     
GetGlHeadXml(sRetVal)

Set nodAccHead = AccHeadData.documentElement


'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=00&BookNo="+iBookNo+"&AccHead=0","","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
if nodAccHead.hasChildNodes then
	For Each HeaderNode In nodAccHead.childNodes
		document.formname.hAccountCode.value=HeaderNode.Attributes.Item(0).nodeValue
		document.formname.hAccountName.value=HeaderNode.Attributes.Item(3).nodeValue&"&nbsp;"
		document.formname.txtPayTo.value=HeaderNode.Attributes.Item(3).nodeValue
	next
end if 'End of GL Head Processing
End function
'---------------------End Of Function showGLHead--------------------------
function showPartyHead(sOrgId,sPartyType)

Dim arrTemp,sRetValue,sTemp,sParTy,sParSubType,sParCode,sPartyName


        sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
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


        '    OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
        '    arrTemp = split(OutValue,":")

        '    while UBound(arrTemp) = 0
	    '        OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	    '        arrTemp = split(OutValue,":")
        '    wend

        '    sRetValue = OutValue
        '    if UBound(arrTemp) <= 1 then exit function
        '    sTemp = Split(sRetValue,":")
        '    sParTy = sTemp(4)
        '    sParSubType = sTemp(3)
        '    sParCode = sTemp(1)
        '    sPartyName = sTemp(0)
        
        if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sParTy = ndEntry.getAttribute("RetField3")
		            sParSubType = ndEntry.getAttribute("RetField4")
		            sParCode = ndEntry.getAttribute("RetField1")
		            sPartyName = ndEntry.getAttribute("RetField0")
		        exit for
                end if
            next
        end if

'MsgBox sParTy&" >> " & sParSubType


'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

'if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	'For Each HeaderNode In nodAccHead.childNodes
		document.formname.hAccountCode.value=sPartyType&"?"&sParCode
		document.formname.hAccountName.value=sPartyName
		document.formname.txtPayTo.value=sPartyName

	'next
'end if 'End of Party Head Processing
End function
'---------------------End Of Function showGLHead--------------------------
Function DisplayTerms(objCounterType)
dim sDate,iInstallNo,sType

	sDate=document.formname.ctlDate.getdate
	document.formname.txtStartDate.value=sDate
	iInstallNo=document.formname.txtInstallmentno.value
	sType=objCounterType(objCounterType.selectedIndex).value

	if ValidateAmount(document.formname.txtLoanAmount.value)=false then
		document.formname.txtLoanAmount.select
		exit Function
	end if
	if 	trim(document.formname.txtInterstRate.value)=""	 then
		MsgBox "Enter Interst Rate"
		document.formname.txtInterstRate.select
		exit function
	elseif not IsNumeric(document.formname.txtInterstRate.value) then
		MsgBox "Interst Rate should be a numeric value"
		document.formname.txtInterstRate.select
		exit function
	elseif CDbl(document.formname.txtInterstRate.value)	<0 or CDbl(document.formname.txtInterstRate.value)	>100 then
		MsgBox "Interst Rate should >0 and <100"
		document.formname.txtInterstRate.select
		exit function
	end if

	if Trim(iInstallNo)="" then
		MsgBox("Enter no of Installment")
		document.formname.txtInstallmentno.focus
		objCounterType.selectedIndex=0
		exit Function
	elseif not (IsNumeric(iInstallNo)) then
		MsgBox("No of Installment should be number")
		document.formname.txtInstallmentno.select
		objCounterType.selectedIndex=0
		exit Function
	end if

	ClearTable
	if objCounterType.selectedIndex <> "0" then
		j=1
		set oRow = document.all.tblTerms.insertRow(1)
		InsertCell oRow,1,"",1,"ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"",GetInterval(sDate,sType,0),"ExcelDisplayCell","left","",0,0,0,0,""
		InsertCell oRow,2,"txtAmount"&j,"","ExcelInputCell","","",11,10,0,0,""
		InsertCell oRow,2,"txtPrincipal"&j,"","ExcelInputCell","","",11,10,0,0,""
		InsertCell oRow,2,"txtInterst"&j,"","ExcelInputCell","","",11,10,0,0,""

		for j=2 to iInstallNo
			set oRow = document.all.tblTerms.insertRow(j)
			InsertCell oRow,1,"",j,"ExcelSerial","Center","",0,0,0,0,""
			InsertCell oRow,1,"",GetInterval(sDate,sType,j-1),"ExcelDisplayCell","left","",0,0,0,0,""
			InsertCell oRow,2,"txtAmount"&j,"","ExcelInputCell","","",11,10,0,0,""
			InsertCell oRow,2,"txtPrincipal"&j,"","ExcelInputCell","","",11,10,0,0,""
			InsertCell oRow,2,"txtInterst"&j,"","ExcelInputCell","","",11,10,0,0,""
		next
	end if
end Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblTerms.rows.length - 1
		document.all.tblTerms.deleteRow(1)
	next
end Function
'--------
function GetInterval(sDate,sIntervalType,iInterval)
	dim iMonth,iYear,iDay,sDate1

	iDay=mid(sDate,1,2)
	iMonth=cint(mid(sDate,4,2))
	iYear=mid(sDate,7,4)
	sDate1=DateSerial(iYear,iMonth,iDay)

	select Case sIntervalType
		Case "M"
				sDate1=DateAdd("m",CInt(iInterval),DateSerial(iYear,iMonth,iDay))
		Case "Q"
				sDate1=DateAdd("m",CInt(iInterval)*4,DateSerial(iYear,iMonth,iDay))
		Case "H"
				sDate1=DateAdd("m",CInt(iInterval)*6,DateSerial(iYear,iMonth,iDay))
		Case "Y"
				sDate1=DateAdd("yyyy",CInt(iInterval),DateSerial(iYear,iMonth,iDay))
	end select

	iDay=Day(sDate1)
	iMonth=Month(sDate1)
	iYear=year(sDate1)

	if cint(iDay) <10 then
		iDay="0"&iDay
	end if
	if cint(iMonth) <10 then
		iMonth="0"&iMonth
	end if
GetInterval=cstr(iDay)&"/"&cstr(iMonth)&"/"&cstr(iYear)
End function

function checksubmit()
	if 	document.formname.hAccountCode.value=""	 then
		MsgBox "Select Account Head"
		document.formname.selAccType.focus
		exit function
	end if
	if 	trim(document.formname.txtReason.value)=""	 then
		MsgBox "Enter Reason"
		document.formname.txtReason.select
		exit function
	end if
	if ValidateAmount(document.formname.txtLoanAmount.value)=false then
		document.formname.txtLoanAmount.select
		exit Function
	end if
	if 	trim(document.formname.txtInterstRate.value)=""	 then
		MsgBox "Enter Interst Rate"
		document.formname.txtInterstRate.select
		exit function
	elseif not IsNumeric(document.formname.txtInterstRate.value) then
		MsgBox "Interst Rate should be a numeric value"
		document.formname.txtInterstRate.select
		exit function
	elseif CDbl(document.formname.txtInterstRate.value)	<0 or CDbl(document.formname.txtInterstRate.value)	>100 then
		MsgBox "Interst Rate should >0 and <100"
		document.formname.txtInterstRate.select
		exit function
	end if

	if 	trim(document.formname.txtInstallmentNo.value)=""	 then
		MsgBox "Enter No of Installment"
		document.formname.txtInstallmentNo.select
		exit function
	elseif not IsNumeric(document.formname.txtInstallmentNo.value) then
		MsgBox "Installment should be a numeric value"
		document.formname.txtInstallmentNo.select
		exit function
	end if

	if document.formname.selpayTerms.selectedIndex=0 then
		MsgBox "Select Payment Terms "
		document.formname.selpayTerms.focus
		exit function
	end if

	if document.formname.selUserId.selectedIndex=0 then
		MsgBox "Select Approver "
		document.formname.selUserId.focus
		exit function
	end if
	document.formname.B4.disabled = True
	document.formname.submit
End function

FUNCTION ValidateAmount(dAmount)
	if  trim(dAmount)="" then
		Msgbox("Amount Cannot be blank")
		ValidateAmount=false
		exit Function
	elseif IsNumeric(dAmount)=false then
		Msgbox("Enter Numeric values for Amount")
		ValidateAmount=false
		exit Function
	elseif CDbl(dAmount)<1 or CDbl(dAmount)>9999999999.99 then
		Msgbox("Amount should be >1 and < 9999999999.99")
		ValidateAmount=false
		exit Function
	end if
	ValidateAmount=true
END FUNCTION
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="PmtReqLoanInsert.asp">
<input type="hidden" name="hFlag" value="<%=sRequestType%>">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hAccountCode" value="">
<input type="hidden" name="hAccountName" value="">
<input type="hidden" name="txtStartDate" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%
			if sRequestType="L" then
				Response.Write "Loan Payment"
			else
				Response.Write "Hire Purchase Payment"
			end if

		%>

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<!--<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td class="TabCell" valign="bottom" align="center" width="110">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Request Selection</td>
									</tr>
								  </table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="132">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
								   <tr>
									  <td align="center">Requisition Details</td>
									</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
								</td>
                            </tr>
						</table>
					</td>
				</tr>-->
				<TR>
					<TD class=TabBody>
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
													<table cellpadding="0" cellspacing="0">
														<!--<tr>
															<td class=FieldCell width="140"> Unit</td>
															<td width="250">
                                                            <span class="DataOnly"><%=sOrgName%></span>
                                                            </td>
														</tr>-->
														<tr>
															<td class="FieldCell" width="105">Party Type</td>
															<td class="FieldCell">
															    <select size="1" name="selAccType" class="FormElem" onChange="selAccountHead(this)">
									   							 	<option value="A">Select Party Type</option>
									   							 	<!--option value="G">General Ledger</option-->
															   		 <%populatePartyType(sOrgId)%>
															        </select>
															        </td>
															</tr>
															   <tr>
															<td class="FieldCell" width="105">Pay To</td>
															<td class="FieldCell"> <input type="text" name="txtPayTo" size="45" class="FormElem"></span> </td>
															    </tr>
														<tr>
															<td class=FieldCell width="140"> Reason</td>
															<td>
                                                            <input type="text" name="txtReason" size="50" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Loan Amount</td>
															<td>
                                                            <input type="text" name="txtLoanAmount" size="15" style="text-align:right" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Interest&nbsp;Rate&nbsp;</td>
															<td>
                                                            <input type="text" name="txtInterstRate" style="text-align:right" size="5" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140">
                                                            Number of Installments</td>
															<td>
                                                            <input type="text" name="txtInstallmentNo" size="5" style="text-align:right" class="FormElem">
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140">
                                                            Starting From</td>
															<td>
															  <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Repayment
                                                              Terms</td>
															<td>
                                                            <select size="1" name="selpayTerms" class="Formelem" onChange="DisplayTerms(this)">
															 <option value="0">Select Payment Term</option>
                                                              <option value="M">Monthly</option>
                                                              <option value="Q">Quaterly</option>
                                                              <option value="H">Half yearly</option>
                                                              <option value="Y">Yearly</option>
                                                            </select>
                                                            </td>
														</tr>
														<tr>
															<td class=FieldCell width="140"> Immediate Approver</td>
															<td class=subCell>
											        <select size="1" name="selUserId" class="FormElem">
											<option value="0">Immediate Approver</option>
											<%=populateEmployee%>
											    </select>
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top">
												<DIV class=frmBody id=frm1 style="width: 378; height: 153">
                                                <table border="0" id="tblTerms"cellspacing="1" class="ExcelTable">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="26">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" width="77">Pay
                                          By</td>
                                        <td class="ExcelHeaderCell" align="center" width="85" valign="top">Amount
                                          to be Paid</td>
                                        <td class="ExcelHeaderCell" align="center" width="85" valign="top">Principal</td>
                                        <td class="ExcelHeaderCell" align="center" width="60">Interest</td>
                                            </tr>

                                                </table>
												</div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
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
																<input type="button" value="Ok" name="B4" class="ActionButton" onclick="checksubmit()">
																 <input type="reset" value="Reset" name="B1" class="ActionButton" >
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