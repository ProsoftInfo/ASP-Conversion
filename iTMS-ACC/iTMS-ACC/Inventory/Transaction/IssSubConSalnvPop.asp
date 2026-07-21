<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssSubConSalInvPop.asp
	'Module Name				:	Inventory (Issue - Sales Invoice Case)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	MARCH 25,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<%
	Dim Salrs
	Dim sHeading,sOrgCode,sQuery,sCode,sValue,sValue1
	Dim sPOSMandatory,iCountPOS,sUserID
	sHeading = Request.QueryString("sHead")
	sOrgCode = Request.QueryString("OrgCode")
	
	sUserID = getuserid()

	set Salrs=Server.CreateObject("ADODB.RecordSet")
	
	sQuery = "Select isNull(MandatoryPOS,'N') from APP_M_ApplicationSetup where ApplicationCode = 3 and ReferenceCodeNo = 3"
	Salrs.Open sQuery,con
	if not Salrs.EOF then
	    sPOSMandatory = Salrs(0)
	end if
	Salrs.Close 
iCountPOS = 0

	'Response.Write sOrgCode
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root Confirm="N"/></script>
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="TaxData"></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Dim objTemp ,Root
set objTemp = window.dialogArguments
set Root =objTemp.documentElement
Root.setAttribute "Confirm","N"
'***********************************************************
Function FinalSubmit()
	Dim SalInvNode
	if trim(document.formname.cmbInvType.value)="0" then
		alert("Select Invoice Type")
		document.formname.cmbInvType.focus
		exit function
	end if

	if trim(document.formname.cmbSaletype.value)="0" then
		alert("Select Invoice Type")
		document.formname.cmbSaletype.focus
		exit function
	end if
	
	if Trim(document.formname.hPOSMandatory.value)="Y" then
	    if CDbl(document.formname.hCountPOS.value)=0 then
	        alert("Create POS")
	        exit function
	    end if
	end if
	
	set SalInvNode = objTemp.createElement("SALINV")
		SalInvNode.setAttribute "InvType",document.formname.cmbInvType(document.formname.cmbInvType.selectedIndex).value
		SalInvNode.setAttribute "SalType",document.formname.cmbSaletype(document.formname.cmbSaletype.selectedIndex).value
		SalInvNode.setAttribute "InvTypeName",document.formname.cmbInvType(document.formname.cmbInvType.selectedIndex).text
		SalInvNode.setAttribute "SalTypeName",document.formname.cmbSaletype(document.formname.cmbSaletype.selectedIndex).text
		if trim(document.formname.hPOSMandatory.value)="Y" then
			SalInvNode.setAttribute "POS", document.formname.cmbPOS(document.formname.cmbPos.selectedIndex).value
			SalInvNode.setAttribute "POSName",document.formname.cmbPOS(document.formname.cmbPos.selectedIndex).text
		else
			SalInvNode.setAttribute "POS",""
			SalInvNode.setAttribute "POSName",""
		end if
	Root.appendChild SalInvNode

	Root.setAttribute "Confirm","Y"
	Root.setAttribute "POConfirm","N"
window.close
End Function
'*************************************************************
Function init()
	Dim root,sTemp,objhttp
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	sOrgId = document.formname.selUnit.value
End Function
'******************************************
Function window_onunload()
'	alert(Root.xml)
	set window.returnvalue = objTemp.documentElement
End Function
'==================================================================
Function PopTaxType()
    sInvType =document.formname.cmbInvType(document.formname.cmbInvType.selectedIndex).value 
	set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "GET","../../Sales/Transaction/GetSaleType.asp?InvType="&sInvType,false
		objhttp.send
		
		if trim(objhttp.responseXML.xml)<>"" then
		   TaxData.loadXML(objhttp.responseXML.xml)
           	set ndTaxRoot = TaxData.documentElement
		    if ndTaxRoot.hasChildNodes() then
    		        document.formname.cmbSaletype.length = 0
    		        document.formname.cmbSaletype.length = document.formname.cmbSaletype.length + 1
		            document.formname.cmbSaletype(document.formname.cmbSaletype.length-1).text = "Select"
		            document.formname.cmbSaletype(document.formname.cmbSaletype.length-1).value = "0"
		        for each ndTaxDet in ndTaxRoot.childNodes
				
	                document.formname.cmbSaletype.length = document.formname.cmbSaletype.length + 1
		            document.formname.cmbSaletype(document.formname.cmbSaletype.length-1).text = ndTaxDet.getAttribute("Name")
		            document.formname.cmbSaletype(document.formname.cmbSaletype.length-1).value = ndTaxDet.getAttribute("Code")
		        next
		    end if
		    document.formname.cmbSaletype.selectedIndex = 0
        else
            alert(objhttp.responseText)
		end if
End Function
'-------------------------------
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad ="init()">
<form method="POST" name="formname">
<INPUT TYPE=HIDDEN NAME="hPurType" value="">
<INPUT TYPE=HIDDEN NAME="hPurTypeName" value="">
<INPUT TYPE=HIDDEN NAME="selUnit" value="<%=sOrgCode%>">
<Input type="hidden" name="hPOSMandatory" value="<%=sPOSMandatory%>">

<OBJECT id=penDet type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11" VIEWASTEXT>
<PARAM name="Command" value="HH Version">
</OBJECT>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sHeading%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>

		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td width="10px"></td>
					<TD class=TabBodywithtopline>

						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td width="10">
								</td>
								<td>
									<table border=0 cellspacing=0 cellpadding=0 width="100%">
									
										<%if sPOSMandatory="Y" then %>
										<tr>
										    <td align="Left" class="FieldCell">Point Of Sale 
											<td align="Left" class="FieldCellSub">
										    <%
												Response.Write "<font color=red>"
										        'sQuery = "Select POSID,POSTerminalNumber,POSDescription from Sal_M_PointOfSales where OrganisationCode = "& sOrgCode    
										        sQuery= "Select P.POSID,POSTerminalNumber,POSDescription from Sal_M_PointOfSales P,SAL_M_PointOfSalesAllotment A where P.POSID=A.POSID and AllottedToEmployeeID = "& sUserID &" and OrganisationCode = '"& sOrgCode &"'"
										       ' Response.Write "<p>"& sQuery &"</p>"
										        Salrs.Open sQuery,con
										        if not Salrs.EOF then
										            Response.Write "<select id=cmbPOS class=FormElem>"
										            do while not Salrs.EOF
										                Response.Write "<option value="&Salrs(0)&">"& Trim(Salrs(1)) &"-"& Trim(Salrs(2))&"</option>"
										                iCountPOS = iCountPOS  + 1
										                Salrs.MoveNext 
										            loop
										            Response.Write "</select>"
										        end if
										        Salrs.Close 
										    %>
										    <Input type="hidden" name="hCountPOS" value="<%=iCountPOS%>">
											</td>
										</tr>
										<%end if ' %>
							
										<tr>
											<td class="FieldCell">Invoice Type
											</td>
											<td class="FieldCellSub">
												<select size="1" name="cmbInvType" class="FormElem" onChange="PopTaxType()">
													<Option Value="0" selected>Select Invoice Type </Option>
													<Option Value="D">Direct Invoice</Option>
													<Option Value="A">Mill Sales</Option>
													<Option Value="T">Depot Transfer</Option>
													<Option Value="U">Inter Unit Transfer</Option>
													<Option Value="C">Transfer for Conversion</Option>
													<Option Value="CO">Consignment Transfer</Option>
													<Option Value="S">Sample Invoice</Option>
													<Option Value="I">Stock Issued for sales</Option>
													<Option Value="R">Conversion charge</Option>
													<Option Value="CB">Cash Bill</Option>
													<Option Value="NEB">Non Excise Bill</Option>
													<Option Value="EB">Excise Bill</Option>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">
 												Type of Sale
 											</td>
 											<td class="FieldCellSub">
 												<select size="1" name="cmbSaletype" class="FormElem">
												<option value="0" selected>Select Sale Type</option>
												<%
											'		sQuery = "Select InvoiceType,InvTypeShortName,InvoiceTypeName from VwSalInvTypes "
											'	  	With Salrs
											'	  		.CursorLocation = 3
											'	  		.CursorType = 3
											'	  		.Source = sQuery
											'	  		.ActiveConnection = con
											'	  		.Open
											'	  	End with
											'	  	Set Salrs.Activeconnection = nothing
											'	  	Set sCode = Salrs(0)
											'	  	Set sValue = Salrs(1)
											'	  	Set sValue1 = Salrs(2)
											'	  	Do while not Salrs.EOF
											'	  		%>
											'	  			<option value="<%Response.Write sCode%>"><%Response.Write sValue %>-<% Response.Write sValue1%></option>
											'	  		<%
											'		Salrs.MoveNext
											'		Loop
											'		Salrs.Close
											    %>
												</select>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="ActionCell">
									<input type=button name="btnProceed" value="Submit" class="ActionButtonX" onClick="FinalSubmit()">
									<!--<input type=button name="btnReset" value="Cancel" class="ActionButtonX" onClick="window.close">-->
								</td>
							</tr>


                        </table>
					</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
Function popAccountHead(sOrgID,iSelAccHead,sStockType)

Dim dcrs,iAccHeadNo,sAccHeadName,iCCExist,iAHExist,sSqlTemp

Set dcrs = Server.CreateObject("ADODB.RecordSet")

if trim(sStockType) = ""  then sStockType = "S"

sStockType = ucase(trim(sStockType))


	'sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForInventApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	if trim(sStockType) = "C" then
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccHeadForFAApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	else
		sSqlTemp = "SELECT A.AccountHead,B.AccountDescription,B.CostCenterExists,B.AnalyticalHeadExists,B.AccountHeadCode  from VwAccheadforPurchaseApp A,Acc_M_GLAccountHead B where A.OUDefinitionID='" & trim(sOrgID) & "' and A.AccountHead=B.AccountHead"
	end if
	'Response.Write "AccHead="&sSqlTemp
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSqlTemp
		.ActiveConnection = con
		.Open
	End With

	Set dcrs.ActiveConnection = Nothing

	Set iAccHeadNo = dcrs(0)
	Set sAccHeadName = dcrs(1)
	Set iCCExist = dcrs(2)
	Set iAHExist = dcrs(3)

	Do while not dcrs.EOF
		if trim(iSelAccHead) = trim(iAccHeadNo) then
			Response.Write "<option value="&trim(iAccHeadNo)&" selected>" & trim(sAccHeadName) & "</option>" & vbCr
		Else
			Response.Write "<option value="&trim(iAccHeadNo)&">" & trim(sAccHeadName) & "</option>" & vbCr
		End if
		dcrs.MoveNext
	Loop
'End if
dcrs.Close

End Function
%>