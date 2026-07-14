<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TDSGroupingSetup.asp
	'Module Name				:	Accounts-TDS (Master Amedment)
	'Author Name				:	Kumar K.A.
	'Created On					:	January 17 2007
	'Modified By				:	UmaMaheswari S
	'Modified On				:	May 05, 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	
	'Procedures/Functions Used	:
	'Internal Variables			:

	'Database					:	ITMS_Test
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
</head>
<script type="application/xml" data-itms-xml-island="1" ID="TempData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="GLHeadData"><Root></Root></script>
<%
Dim Objrs,OrgUnit,GroupName,GNameText,sType,sTemp,nGroupHeadID

Set Objrs = Server.CreateObject("ADODB.RecordSet")

OrgUnit = Session("organizationcode")
sTemp  = Trim(Request.QueryString("CallType"))
sType = Split(sTemp,":")(0)

'Response.Write "<p><Font color=red>Data="&sTemp

If sType = "E" Then
	'GroupName = Request.Form("SelGPName")
	GroupName =  Split(sTemp,":")(1)
Else
	GroupName = "0"
	nGroupHeadID = "0"
End IF
GNameText = Request.Form("TxtGroupName")

sql = "Select GroupID,GroupName From ACC_M_TDSGroup Where isNull(Useable,'Y') = 'Y' and OUDefinitionid = '"&OrgUnit&"' "
With Objrs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sql 
	.Open 
End With
Do While Not Objrs.EOF 
	If (Cstr(GroupName) = Cstr(Objrs(0))) Or (Cstr(GNameText)= Cstr(Objrs(1))) Then 
		GroupName = Objrs(0)
		GNameText = objrs(1)
	End IF
	Objrs.MoveNext 
Loop
Objrs.Close 
'Response.Write "<p>Data="&GroupName & " = "& GNameText
'If GroupName="A" Then GroupName = 0
%>
<script>
window.__itmsPopupCompat = { type: "tdsGroupingSetup" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="post" name="formname" action="">
	<input type="hidden" name="SelSeh" value="3">
	<input type="hidden" name="OrgID" value="<%=OrgUnit%>">
	<input type="hidden" name="GroupName" value="<%=GroupName%>">
	<input type="hidden" name="hAccHead" value="">
	<input type="hidden" name="hType" value="<%=sType%>">
	
	<input type="hidden" name="iButCount" value="1">
	<input type="hidden" name="iRowCount" value="1">
	<input type="hidden" name="hHeadName" value="">
	
	<input type="hidden" name="TxtFormula" size="40" maxlength="30" class="FormElem"></td>
	<input type="hidden" name="TxtVal" value="">
	<input type="hidden" name="iTxtCount" value="">
	
	<Input type="hidden" name="hRequest" value="<%=Trim(sTemp)%>">
	<input type="hidden" name="hDelFrom" value="S">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="middle" class="PageTitle" height="20">TDS Grouping Setup
			</td>
		</tr>

		<tr>
			<td align="middle" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
									</td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<!--<tr>
												<td class="FieldCell">Unit
												</td>
												<td class="FieldCellSub"><select size="1" name="selUnit" class="FormElem" Onchange="UnitChange()" >
														<option selected>Select</option>
                        <% 
							Dim sql
							sql = "select Orgunitdescription,OUDefinitionID from DCS_OrganizationunitDefinitions WHere Len(OUDefinitionID) > 4 Order By OUDefinitionID "
							With Objrs
								.CursorLocation = 3
								.CursorType = 3
								.Source = sql 
								.ActiveConnection = con
								.Open 
							End With
								While Not Objrs.EOF 
                        %>
                        <% If OrgUnit = Objrs(1) Then %>
                        <Option  Value="<%=Objrs(1)%>" Selected="<%=Objrs(1)%>"><%=Objrs(0)%></Option>
                        <% Else %>
                        <Option Value="<%=Objrs(1)%>"><%=Objrs(0)%></Option>
                        <%End If%>
						<%
							Objrs.MoveNext
							Wend
							Objrs.Close
						%>                        
							</select>
							</td>
							<td class="FieldCellSub">
							</td>
							<td class="FieldCellSub">Created On
							</td>
							<td class="FieldCellSub">
												</td>
											</tr>-->

											<tr>
												<td class="FieldCell">TDS Group Name
												</td>
												<td class="FieldCellSub" colspan="4">
													<%If sType = "E"and 1=2Then%>
													<select size="1" name="selGPName" class="FormElem" OnChange="GNameChange()">
														<option selected value="0">Select</option>
														<%
														IF CStr(OrgUnit) <> "" Then
															sql = "Select GroupID,GroupName From ACC_M_TDSGroup Where isNull(Useable,'Y') = 'Y' and OUDefinitionid = '"&OrgUnit&"' "
															With Objrs
																.CursorLocation = 3
																.CursorType = 3
																.ActiveConnection = con
																.Source = sql 
																.Open 
															End With
															While Not Objrs.EOF 
															%>
															<% If (Cstr(GroupName) = Cstr(Objrs(0))) Or (Cstr(GNameText)= Cstr(Objrs(1))) Then 
																	GroupName = Objrs(0)
																	GNameText = objrs(1)
															%>
																<Option Selected Value="<%=Objrs(0)%>"><%=Objrs(1)%></Option>
															<% Else %>
																<Option Value="<%=Objrs(0)%>"><%=Objrs(1)%></Option>
															<%End If%>
															<%
															Objrs.MoveNext 
															Wend
															Objrs.Close 
														End IF
														%>
														
													<!--<option value="A">Add New</Option>-->
													</select> 
													<%End IF%>
													<input name="TxtGroupName" size="30" maxlength="100" class="FormElem" Value="<%=GNameText%>" >
												</td>
											</tr>
											<tr>
												<td class="FieldCell">TDS Head
												</td>
												<td class="FieldCellSub" colspan="4"><select size="1" name="selHead" class="FormElem" Onchange="">
													<option selected value="0">Select</option>
													<%
													sql = "select GroupHeadID,GroupHeadName,ComputeMode,AcheadCode,ComputeFormula,Herarchy from ACC_M_TDSHeadComputation Where GroupID="&Cint(GroupName)
													With Objrs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = sql
														.Open 
													End With
													While Not Objrs.EOF 
													%>	
													<option Value="<%=Objrs(0)&","&Objrs(2)&","&Objrs(3)&","&Objrs(4)&","&Objrs(5)%>"><%=Objrs(1)%></option>
													<%
													Objrs.MoveNext 
													Wend
													Objrs.Close 
													%>
													
													<option Value="A">Add New</option>
													</select> 
													<input name="TxtHead" maxlength="30" class="FormElem" >
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Hierarchy 
													<input name="TxtHierachy" size="4" maxlength="3" class="FormElem" >
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Applicable A/c Head
												</td>
												<td class="FieldCellSub" colspan="4">
													<input name="TxtAccHeadName" size="40" maxlength="30" class="FormElem" >
													<IMG style="cursor: pointer" onclick="SuppName()" height=11 alt="Select Account Head" src  ="../../assets/images/iTMS%20Icons/EntryIcon.gif" width=10 align=center border=0 >
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Computation Mode
												</td>
												<td class="FieldCell" colspan="4">
													<input type="radio" value="V2" checked name="R1">
 													Flat&nbsp;&nbsp; 
													<input type="radio" value="V1" name="R1">
 													Percentage
												</td>
											</tr>

											<tr>
												<td class="FieldCell">
											
												</td>
												<td class="FieldCellSub">
													<input type="button" value=" Add " name="ButAdd" class="AddButtonx" Onclick="AddGroup()" >
													<input type="button" value=" Update " name="ButUpdate" class="AddButtonx" Onclick="UpdateGroup()" disabled>
													<input type="button" value=" Delete " name="ButDelete" class="AddButtonx" Onclick="DeleteGroup()" disabled>
												</td>
												<td class="FieldCellSub">
															
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">
												</td>
											</tr>

										</table>
									</td>
									<td align="middle">
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
									</td>
									<td>
										<div class="frmBody" id="frm2" style="WIDTH: 555px; HEIGHT: 231px">
											<table border="0" cellspacing="1" class="ExcelTable" width="550" id="tblemp">
												<tr>
													<td class="ExcelHeaderCell" align="middle" width="10" rowspan="2">S.No.
													</td>
													<td class="ExcelHeaderCell" align="middle" colspan="5">TDS Group Name
													</td>
												</tr>

												<tr>
													<td class="ExcelHeaderCell" align="middle" >
													</td>
													<td class="ExcelHeaderCell" align="middle" width="100">TDS Head
													</td>
													<td class="ExcelHeaderCell" align="middle" width="300">Account Head
													</td>
													<td class="ExcelHeaderCell" align="middle" > Detail
													</td>
												</tr>
												<%
												Dim Txtcount,HID,Formula
													txtcount = 0
													sql = "Select T.GroupHeadID,T.GroupHeadName,T.ComputeMode,T.AcHeadCode,M.AccountDescription,T.ComputeFormula From "&_
														   "ACC_M_TDSHeadComputation T,Acc_M_GLAccountHead M Where T.AcHeadCode = M.AccountHead "&_
															" and T.GroupID = '"&Cint(GroupName)&"'"
													'Response.Write sql
													With Objrs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = sql 
														.Open 
													End With
													While Not Objrs.EOF 
													HID = Objrs(0)						
													If Objrs(5) <> "" then Formula = Objrs(5) Else Formula = ""
													%>
													<tr>
														<td class="ExcelSerial" align="center"><%=Objrs(0)%> 
														</td>
														<td class="ExcelDisplayCell" align="left" >
														<a href="#" onclick="ShowVouch(<%=Objrs(0)%>); return false;" class="ExcelDisplayLink">Edit</a></td>

														<td class="ExcelDisplayCell" ><%=Objrs(1)%>
														</td>
														<td class="ExcelDisplayCell"><%=Objrs(4)%> 
														</td>
														<%If Objrs(2) = "F" Then %>
														<%If Formula <>"" Then%>
																<td class="ExcelInputCell"><input name="TxtFormula" size="14" maxlength="30" Value="" class="FormElem" style="text-align:right"></td>
															<%Else%>
																<td class="ExcelInputCell"><input name="TxtFormula" size="14" maxlength="30" Value="" class="FormElem" style="text-align:right"></td>
															<%End If%>
																<input type="hidden" name="TxtVal" value="<%=HId%>">
																<input type="hidden" name="iTxtCount" value="<%=TxtCount%>">
														<%
														Txtcount = Txtcount + 1
														Else%>			
														<td class="ExcelDisplayCell" align="center"><input type="Button" value=" Select " name="BtnCalc" class="ActionButtonX" onclick="Calc(<%=Cint(GroupName)%>,<%=Cint(HID)%>)"></td>
														<%End If%>
														
														
													</tr>
													<%
														Objrs.MoveNext 
														Wend	
													Objrs.Close 
													If Txtcount>0 Then
													%>
														<input type="hidden" name="TCount" value="1">
														<%
														Else
														%>
															<input type="hidden" name="TCount" value="0">
														<%End If%>
												</table>
											</div>
									</td>
									<td align="middle">
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
										<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="center" class="ActionCell" align="middle">
													<input type="Button" value="Save" name="BtnSubmit" class="ActionButton" onclick="Submit();">
													<input type="Button" value="Delete" name="BtnDel" class="ActionButton" onclick="TDSDel()">
													<!--input type="button" value="Done" name="B1" class="ActionButton" -->
													
												</td>
											</tr>

										</table>
									</td>
									<td align="middle">
										<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="BottomPack">
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
