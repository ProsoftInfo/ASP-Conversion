<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtReqChequeEntry.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 19, 2003
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
<%
dim sOrgId,sOrgName,sAccCode,sAccName,sReqType,dTransLimit,oDOM
'sOrgId=Request.Form("selUnitId")
'sOrgName=Request.Form("hUnitName")
sOrgId = session("organizationcode")
sOrgName = session("orgshortname")
sReqType=Request.Form("hReqTypeS")


Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

'Response.Write "<p>dTransLimit="&dTransLimit

%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<xml id="PayableData"><Payables ReqType="B"/></xml>
<XML id="PartyData"><Root></Root></XML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT language="vbscript">
function selAccountHead(objAcc)
dim sOrgId,sPartyCode,arrTemp,sTemp,sRetValue,sTempbVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
	for i=1 to  document.all.tblPayable.rows.length - 1
		 document.all.tblPayable.deleteRow(1)
	next

	if objAcc.selectedIndex >0 then
		'If selected account Head is Party type
		sPartyCode=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
		sOrgId=document.formname.hUnitId.value

		'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+ sOrgId&"&Party="&sPartyCode,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

	'	OutValue  = showModalDialog("PartySelection.asp?orgId="+ sOrgId&"&Party="&sPartyCode,"","dialogHeight:480px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
	'	arrTemp = split(OutValue,":")
'
'		while UBound(arrTemp) = 0
'			OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'			arrTemp = split(OutValue,":")
'		wend
'
'		IF UBound(arrTemp) <= 1 Then
'			document.formname.selAcctype.selectedIndex = 0
'			Exit Function
'		End IF
'

'		sRetValue = OutValue
'		sTemp = Split(sRetValue,":")
'		sParTy = sTemp(4)
'		sParSubType = sTemp(3)
'		sParCode = sTemp(1)
'		sPartyName = sTemp(0)

        sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId&"&Party="&sPartyCode,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
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
                    sParTy = ndEntry.getAttribute("RetField3")
		            sParSubType = ndEntry.getAttribute("RetField4")
		            sParCode = ndEntry.getAttribute("RetField1")
		            sPartyName = ndEntry.getAttribute("RetField0")
		        exit for
                end if
            next
        end if
	    

		'MsgBox UBound(sTemp)
		sTempbVal = "1"

		IF Cstr(sTempbVal) = "1" Then
			'MsgBox "Inside "
			'User Has Selected a GL Account Head
			'For Each HeaderNode In nodAccHead.childNodes
				document.formname.hAccountCode.value = sPartyCode&"?"& sParCode
				document.formname.hAccountName.value = sPartyName&"&nbsp;"
				document.formname.txtPayTo.value = sPartyName

				set objhttp = CreateObject("MSXML2.XMLHTTP")
				objhttp.Open "GET","XMLGetPayables.asp?orgId="+sOrgId+"&ParCode="+sPartyCode&"?"& sParCode, false
				objhttp.send
				'alert(objhttp.responseXML.xml)
				if objhttp.responseXML.xml <> "" then
					PayableData.loadXML objhttp.responseXML.xml
					Set Root = PayableData.documentElement

					iSno=0
					for each  nodCC in Root.childNodes
						sDocNo=nodCC.Attributes.Item(0).nodeValue
						sVouNo=nodCC.Attributes.Item(1).nodeValue
						sVouDate=nodCC.Attributes.Item(2).nodeValue

						sInvNo=nodCC.Attributes.Item(3).nodeValue
						sInvDate=nodCC.Attributes.Item(4).nodeValue
						sAmtPayable=FormatNumber(nodCC.Attributes.Item(5).nodeValue,2,,,0)
						sAmtPaid=FormatNumber(nodCC.Attributes.Item(6).nodeValue,2,,,0)

						set oRow = document.all.tblPayable.insertRow(iSno+1)
						InsertCell oRow,1,"",iSno+1,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sVouNo&"-"&sVouDate,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,1,"",sInvNo&"-"&sInvDate,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,1,"",sAmtPayable,"ExcelDisplayCell","right","",0,0,0,0,""
						InsertCell oRow,1,"",sAmtPaid,"ExcelDisplayCell","right","",0,0,0,0,""
						InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","right","",15,13,0,0,"style=""text-align:right"""
						iSno=iSno+1
					next
				end if
			'next
		else
			document.formname.selAcctype.selectedIndex=0
		end if 'End of Party Head Processing
	else
		document.formname.hAccountCode.value=""
		document.formname.hAccountName.value=""
		document.formname.txtPayTo.value=""
	End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead----------------------
function actionDone()
dim sAmount,bFlag,dTotal
Set Root = PayableData.documentElement
bFlag=false
dTotal=0
if 	document.formname.hAccountCode.value=""	 then
	MsgBox "Select Account Head"
	document.formname.selAcctype.focus
	exit function
end if

for each  nodCC in Root.childNodes
	sAmount = 0
	sDocNo=nodCC.Attributes.Item(0).nodeValue
	sAmtPayable=FormatNumber(nodCC.Attributes.Item(5).nodeValue,2,,,0)
	sAmtPaid=FormatNumber(nodCC.Attributes.Item(6).nodeValue,2,,,0)
	sAmount=eval("document.formname.txtDocAmount"&CStr(sDocNo)).value

	if trim(sAmount)<>"" then
		if IsNumeric(sAmount)=false then
			MsgBox "Enter Numeric Value"
			exit function
		elseif 	CDbl(sAmount)<0 or CDbl(sAmount)>9999999999.99 then
			MsgBox "Amount Should Be > 0 and < 9999999999.99"
			exit function
		elseif CDbl(sAmount)> (CDbl(sAmtPayable)-CDbl(sAmtPaid)) then
			MsgBox "Amount is greater than to be paid amount"
			exit function
		else
			'nodCC.Attributes.Item(6).nodeValue=sAmount
			if CDbl(sAmount)>0 then
				bFlag=true
				dTotal=CDbl(dTotal)+CDbl(sAmount)
			end if
		end if
	end if
next

if bFlag=false then
	MsgBox "Request should be created for atleast one Bill "
	exit function
end if

if document.formname.hRequestType.value="C" and CDbl(dTotal)> CDbl(document.formname.hCreditLimit.value)then
		MsgBox "Cash transcation should not exceed "& document.formname.hCreditLimit.value
	exit function
end if

if document.formname.selUserId.selectedIndex=0 then
	MsgBox "Select Approver "
	document.formname.selUserId.focus
	exit function
end if

for each  nodCC in Root.childNodes
	sDocNo=nodCC.Attributes.Item(0).nodeValue
	sAmount=eval("document.formname.txtDocAmount"&CStr(sDocNo)).value
	nodCC.Attributes.Item(6).nodeValue=sAmount
	sAmount = 0
Next

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Mod=CHQ&Name=Payment Requestion", false
	objhttp.send PayableData.XMLDocument

	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		document.formname.submit()
	end if
End function
'---------------------End Of Function actionDone--------------------------
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="PmtReqChequeInsert.asp">
<input type="hidden" name="hUnitName" value="">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hAccountCode" value="">
<input type="hidden" name="hAccountName" value="">
<input type="hidden" name="hCreditLimit" value="<%=dTransLimit%>">


<%if sReqType="A" then %>
<input type="hidden" name="hRequestType" value="B">
<%else%>
<input type="hidden" name="hRequestType" value="C">
<%end if%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="446">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Regular Payment
		<%if sReqType="A" then
			Response.Write "CHEQUE"
		else
			Response.Write "CASH"
		end if
		%>

		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top" height="419">
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
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                            <td class="FieldCellSub" width="139">Payment Type</td>
                                            <td class="FieldCell"><span class="DataOnly">
                                            <%if sReqType="A" then
												Response.Write "CHEQUE"
											else
												Response.Write "CASH"
											end if
											%>
                                            &nbsp;</span>
                                            </td>
                                                </tr>
                                                <!--<tr>
                                            <td class="FieldCellSub" width="139">Unit</td>
                                            <td class="FieldCell"><span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                                </tr>-->
                                                <tr>
                                            <td class="FieldCellSub" width="139">Party Type</td>
                                            <td class="FieldCell">
                                            <select size="1" name="selAcctype" class="FormElem" onChange="selAccountHead(this)">
									   		<option value="S">Select Account Head</option>
									  		 <%populatePartyType(sOrgId)%>
											</select>
                                            </td>
                                                </tr>
                                                <tr>
													<td class="FieldCellSub" width="105">Pay To</td>
													<td class="FieldCell"> <input type="text" name="txtPayTo" size="40" class="FormElem"> </td>
													    </tr>
                                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm1 style="width: 585; height:140;">
                                    <table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center"></td>
                                        <td class="ExcelHeaderCell" align="center">Voucher No- Date</td>
                                        <td class="ExcelHeaderCell" align="center">Bill No - Date</td>
                                        <td class="ExcelHeaderCell" align="center">Bill Amount</td>
                                        <td class="ExcelHeaderCell" align="center">Amount Paid</td>
                                        <td class="ExcelHeaderCell" align="center">Amount To Pay</td>
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
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell" width="130"> Immediate Approver </td>
                            <td>
									<select size="1" name="selUserId" class="FormElem">
												<option value="0">Immediate Approver</option>
												<%=populateEmployee%>
												    </select>

                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
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
 <input type="button" value="Ok" name="B4" class="ActionButton" onClick="actionDone()">
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