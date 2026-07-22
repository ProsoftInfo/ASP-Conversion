<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssSubConProcessDetailsPop.asp
	'Module Name				:	Inventory(Issue for Subcontract)
	'Author Name				:	Ragavendran R
	'Created On					:	Oct 10,2013
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS-SubContract Process Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemAddData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root></Root></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/issSubConProcessSelPop.js"></SCRIPT>

<%
Dim objRs,sDrgNo,sItemType,sDesc,OrderFor,sUnit,sSupp,sSql,iClassCode,iItemCode,sSupplier,bFlag
Dim sClassDesc,sItemDesc,saTemp,sDrawVerNo,sSource,sPRNo,sPRNoStr,iItemRecdAs,iSCProcess
Dim indrwgStoreNo,iItemRecdAt,sInstruct,iTempItemCode,sDNo,sCallFrom,iForPurNo,sItemMode
Dim sReturnable,sReturnItem,iEntryNo

Dim sRequest,sArrTemp
Set objRs = server.CreateObject("Adodb.recordset")
sUnit = Request.QueryString("Unit")


%>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hUnit" value="<%=sUnit%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		Additional Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                   <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td colspan="2">
                                                <div id="tblAddDet" style="display:block;">
                                                    <table width="100%">
                                                        <tr>
                                                        <td class="FieldCell">Sub Contracting Process</td>
                                                        <td class="FieldCell" colspan="4">
                                                            <select size="1" name="cmbSCProcess" class="FormElem" >
										                    <option value="0">Select</option>
											                <%
											                'function to populate SC process
											                 popSubContractProcess sItemType%>
									                    </select></td>
                                                       </tr>

                                                        <tr>
												            <td class="FieldCell">Materials to be&nbsp;received as</td>
												            <td class="FieldCell" colspan="4">
												                <input type="hidden" name="cmbMatRecdAs" value="0">
														        <a class="ExcelDisplayLink" href="#" onclick="SelectItem()">
															        <img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item" width="15" height="15">
														        </a>
														        <span id="SpnMaterialToBeReceived" class="DataOnly">&nbsp;</span>
												            </td>
												        </tr>
                                                            <tr>
												            <td class="FieldCell" valign="top">Instruction</td>
                                                        <td class="FieldCell" colspan="4">
											            <textarea rows="4" name="txtInstruct" cols="60" class="FormElem"></textarea>
                                                        </td>
                                                            </tr>
                                                            <tr>
											            <td class="FieldCellsub" valign="top">Labour Charges </td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtLabourCharge" value="" class=formelem size="10" style="text-align: right">
												            <SELECT size="1" class="Formelem" NAME = "cmbCurrency">
												            <%
													               'To populate currency
														            popSelCurrency(0)
												            %>
												            </select>
											            </td>
                                                    </tr>
                                                       <tr>
											            <td class="FieldCellsub" valign="top">Hard Waste</td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtHardWaste" value="" class=formelem size="10" style="text-align: right">%
											            </td>
                                                    </tr>
                                                       <tr>
											            <td class="FieldCellsub" valign="top">Invisible Waste</td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtInvWaste" value="" class=formelem size="10" style="text-align: right">%
											            </td>
                                                    </tr>
                                                    </table>
                                                </div>
                                            </td>
                                        </tr>


                                    </table>
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
                                            <input type="button" value="Done" name="B4" onClick="saveXML()" class="ActionButton" tabindex="3">
                                            <input type="button" value="Close" name="B6" onClick="javascript:window.close();" class="ActionButton" tabindex="3">
                                            <input type="reset" value="Reset" name="B5" class="ActionButton" tabindex="3">
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
</Html>

<%
function popSubContractProcess(ItemType)
Dim PrcID,PrcName

sSql ="Select SubConProcessID,SubConProcessName from App_M_SubContractProcess"
  '" Where ItemTypeId='" & ItemType & "'"
 'Response.Write sSql
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set objRs.ActiveConnection = nothing

Set PrcID = objRs(0)
Set PrcName = objRs(1)

If not objRs.EOF then
	Do While Not objRs.EOF
			Response.Write("<OPTION VALUE="""&trim(PrcID)&""">"&trim(PrcName)&"</OPTION>" &vbcrlf)
		objRs.MoveNext
	Loop
end if
objRs.Close

End function
%>

<%
'' to fetch the Unit No ,Unit Name
Function popUnitOrgRel(selUnit)
Dim sUnitNo,sUnitName
	sSql  = "Select OUDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OUDefinitionID)> 4"
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set objRs.ActiveConnection = nothing

	Set sUnitNo = objRs(0)
	Set sUnitName = objRs(1)

	If not objRs.EOF then
		Do While Not objRs.EOF
			if trim(sUnitNo) = trim(selUnit) then
				Response.Write("<OPTION VALUE="""&trim(sUnitNo)&""" selected>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sUnitNo)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			end if
			objRs.MoveNext
		Loop
	end if
	objRs.Close
End function
%>

