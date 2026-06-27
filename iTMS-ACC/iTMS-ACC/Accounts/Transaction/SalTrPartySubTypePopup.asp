<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BankInsDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
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
<!--#include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!--#include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim sQry,rs
	Set rs = Server.CreateObject("ADODB.Recordset")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Party Sub Type</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script Language=vbscript>
Dim objTemp,Root,newElem,Hnode
set objTemp = window.dialogArguments
'*************************************************************************************************
Function CheckSubmit()
	window_onunload()
End Function
'*************************************************************************************************
Function window_onunload()
	'alert ObjTemp.xml
	set window.returnValue = ObjTemp.documentElement
	window.close()
end Function
'*************************************************************************************************
Function Init()
	
	set objTemp = window.dialogArguments
	'alert(objTemp.xml)
	
	sExp = "//Party"
	Set TempNode = objTemp.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		DisplaySubType()
	End IF
End Function
'*************************************************************************************************
Function DisplaySubType()
	Dim Root,SlNoCR,SlNoDR,sPartySubTypeCheck,sPartyTypeCheck
	set Root = objTemp.documentElement
	
	SlNoCR = 0
	SlNoDR = 0
	
	If Root.hasChildNodes Then
		sPartySubTypeCheck = Split(Trim(Root.getAttribute("PartySubType")),":")
		sPartyTypeCheck = Split(Trim(Root.getAttribute("PartyType")),":")
		For Each Node in Root.childNodes 
			If Node.nodeName = "Party" Then
				sPartyType = Split(Node.getAttribute("SubType"),"|")(0)
				sPartySubType = Split(Node.getAttribute("SubType"),"|")(1)
				sSubTypeName = Node.Text
				sCheck = Node.getAttribute("Check")
						
				If sPartyType = "CR" Then
					SlNoCR = SlNoCR+1
					set trow = document.all.CRTable.InsertRow(document.all.CRTable.Rows.length)
					
					set Cell=trow.insertCell()
					If sCheck = "LCheck" and sPartyType = "CR" Then
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkCR"& SlNoCR &""" value="""&sPartySubType&":"& sPartyType &""" checked>")
					Elseif sPartyType = "CR" Then
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkCR"& SlNoCR &""" value="""&sPartySubType&":"& sPartyType &""" >")
					End IF
					Cell.appendChild(oText)
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
					
					set Cell=trow.insertCell()
					Cell.innerHTML = sSubTypeName
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					set Cell=trow.insertCell()
					If sCheck = "RCheck" and sPartyType = "CR" Then
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkPartySubType"" checked>")
					Else
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkPartySubType"" >")
					End IF
					Cell.appendChild(oText)
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
				Else
					SlNoDR = SlNoDR+1
					set trow = document.all.DRTable.InsertRow(document.all.DRTable.Rows.length)
					
					set Cell=trow.insertCell()
					If sCheck = "LCheck" and sPartyType = "DR" Then
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkDR"& SlNoDR &""" value="""&sPartySubType&":"& sPartyType &""" checked>")
					Else
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkDR"& SlNoDR &""" value="""&sPartySubType&":"& sPartyType &""">")
					End IF
					Cell.appendChild(oText)
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
					
					set Cell=trow.insertCell()
					Cell.innerHTML = sSubTypeName
					Cell.className="ExcelDisplayCell"
					Cell.align="Left"
					
					Set Cell=trow.insertCell()
					If sCheck = "RCheck" and sPartyType = "DR" Then
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkPartySubTypeDR"" checked>")
					Else
						set oText = document.createElement("<input type=""CheckBox"" name=""ChkPartySubTypeDR"">")
					End IF
					Cell.appendChild(oText)
					Cell.className="ExcelDisplayCell"
					Cell.align="center"
				End IF	'If sPartyType = "CR" Then
			End IF
		Next
	End IF
	
End Function

Function ClearTable()
	Dim i
	for	i = 1 to document.all.InsTab.rows.length - 1
		document.all.InsTab.deleteRow(1)
	next
End function
'*************************************************************************************************

</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "partySubtypeDialog" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="POST" name="formname" action="">
<Input Type="hidden" name="hExists" Value="">
<Input Type="hidden" name="hEditNo" Value="">

<input type="hidden" name="hCtr" value="1">
<input type="hidden" name="hInsType" value="C">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Party Sub Type Details
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
													<table cellpadding="0" cellspacing="0" width="100%">
														<Tr>
														    <td width="100%">
														        <div id="divCRDRSubType" style="display:block" >
														            <Table border="0" cellspacing="0" cellpadding="0" width="100%">
														            <Tr>
														                <Td width="1%"></td>
														                <TD width="48%"  style="height: 35px" valign="top">
														                        <Table  id="CRTable" border="0" cellspacing="1" cellpadding="0" class="ExcelTable" width="100%">
														                            <tr>
														                                <td class="ExcelHeaderCell" colspan=3 align=center>Creditor Party SubType</td>
														                            </tr>
															                        <Tr>
																                        <td class="ExcelHeaderCell" width="10"></td>
																                        <td class="ExcelHeaderCell" align=center>Sub Type</td>
																                        <td class="ExcelHeaderCell" align=center>Status</td>
															                        </tr>
														                        </Table>
														                    
														                </Td>
														                <Td  width="2%" style="height: 35px"></td>
														                <TD  width="48%" style="height: 35px" valign="Top">
																			
														                        <Table id="DRTable" border="0" cellspacing="1" cellpadding="0" class="ExcelTable" width="100%">
														                            <tr>
														                                <td class="ExcelHeaderCell" colspan=3 align=center>Debtor Party SubType</td>
														                            </tr>
															                        <Tr>
																                        <td class="ExcelHeaderCell" width="10"></td>
																                        <td class="ExcelHeaderCell" align=center>Sub Type</td>
																                        <td class="ExcelHeaderCell" align=center>Status</td>
															                        </tr>
														                        </Table>
														                    
														                </Td>
														                <td  width="1%"></td>
														                </tr>
														            </table>
														        </div>
														    </td>
														</Tr>
                                                        
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
															<p align="center">
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
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
