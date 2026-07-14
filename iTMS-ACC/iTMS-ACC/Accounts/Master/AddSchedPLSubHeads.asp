<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedPLSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 29 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	PLSetUp.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>Add P/L Heads</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	Dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName,Objrs,sInsDate
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera
	Set objrs1 = Server.createObject("ADODB.Recordset")
	Set Objrs = Server.createObject("ADODB.Recordset")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	sInsDate = Request("InsDate")
	'Response.Write sInsDate
%>

<script type="application/xml" data-itms-xml-island="1" ID="XmlData">
<Root>
	<Details OrgID="" SchName="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="" Hierachy="" InsDate="<%=sInsDate%>">
	</Details>
</Root>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" ID="TempData"><Root/></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "plBsScheduleSubHeadsPopup", kind: "PL" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="DispName()">
<form method="POST" name="formname">
<input type=hidden name="hOrgId" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sfinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hAccHead" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">P/L Setup
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="650"  >
				<TR>
					<TD class=TabBodyWithTopLine width="786">
						<table border="0" cellpadding="0" cellspacing="0" width="723">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" width="702">
								</td>
                            </tr>
                             <tr>
								<td class="FieldCellSub" width="187">Select Schedule</td>
                                <td class="FieldCell" width="268">
									<select size="1" name="selSch" class="formElem" onchange="SelHead()">
									<option value="0">Select</option>
									<%
									sql = "SELECT PLHeadID, PLHeading, FinYear FROM dbo.Acc_M_PLSetupHeads"
									with Objrs1
										.CursorLocation = 3
										.CursorType = 3
										.ActiveConnection = con
										.Source = sql
										.Open
									End With
									Set Objrs1.ActiveConnection = Nothing
									while not Objrs1.EOF
										IF CStr(sSchedNo) = CStr(Objrs1(0)) Then
									%>
														<Option Value="<%=Objrs1(0)%>" Selected><%=Objrs1(1)%></Option>
										<% Else %>
														<Option Value="<%=Objrs1(0)%>"><%=Objrs1(1)%> </Option>
											<%
										End IF
									%>
									</option>
									<%
									Objrs1.MoveNext
									wend
									Objrs1.Close
									%>
									<Option Value="A"><%="AddNew"%> </Option>
									</select>
									</td>
                                <td class="FieldCell" width="331">
                                    <input type=text name=txtLev size=56 class="Formelem" align="Right" value="" maxlength="200">
								</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="187">Add/Modify Level</td>
										<td class="FieldCell" width="268">
											<Input type="radio" name="optLevel" value="Lev1" class="FormElem"  onclick="LevelFun()" Checked>Level1
											<Input type="radio" name="optLevel" value="Lev2" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
										<td class="FieldCell" width="331">
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="187">Select Level 1</td>
										<td class="FieldCell" width="359">
											<select size="1" name="selLevel1" class="FormElem" OnChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select PLSubID,PLSubSubID,EntryType,Computemode,isNull(PLSubHeadingName,'') as"&_
													" PLSubHeadingName,Hierachy From ACC_M_PLSetupsubheads Where "&_
													" FinYear = '"&sFinyr&"' and PLHeadID = "&sSchedNo&" and PLsubsubid = 0 Order By Hierachy "
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF
													Dim AcVal
													AcVal = 0
													If Objrs1(2)="S" Then
													sql = "select PLSubHeadValue from Acc_T_PLAcDetail where PLHeadID='"&sSchedNo&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"
													With objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sql
														.ActiveConnection = con
														.Open
													End with
													If Not Objrs.EOF= True Then	AcVal = Objrs(0) Else AcVal = 0
													Objrs.Close
													End If
													If objrs1(2)="A" Then
														sql = "select ApplicableACHeadCode from Acc_T_PLAcDetail where PLHeadID='"&sSchedNo&"' and PLSubID='"&objrs1(0)&"' and PLSubSubID='"&objrs1(1)&"'"
													With Objrs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sql
														.ActiveConnection = con
														.Open
													End with
													If Not Objrs.EOF = True Then AcVal = Objrs(0) Else AcVal = 0
													Objrs.close
													End If

											%>
											<option value="<%=Objrs1(0)&","&Objrs1(1)&","&objrs1(2)&","&objrs1(3)&","&AcVal&","&objrs1(5) %>"><%Response.Write (Objrs1("PLSubHeadingName")) %></b></option>
											<%Objrs1.MoveNext
											loop
											Objrs1.Close
											%>
											<option Value="A">Add New</option>
											</select>
										</td>
										<td class="FieldCell" width="233">
                                            <input type=text name=txtLev1 size=56 class="Formelem" align="Right" maxlength="200">
										</td>
									</tr>
								  <tr>
								<td class="FieldCellSub" width="187">Select Level 2</td>
										<td class="FieldCell" width="359">
										<select size="1" name="selLevel2" class="FormElem" Disabled onchange="setlevelfun1()">
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>

									</td>
									<input type=hidden name=hSubHeadName value="">
										<td class="FieldCell" width="233">
                                        <input type=text name=txtLev2 size=56 class="Formelem" align="Right" maxlength="200">
									</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Mode</td>
								<td class="FieldCell" width="599" colspan="2">
									<Input type=radio name=optMode value="D" class="FormElem" checked OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A1" class="FormElem" OnClick="ModeFun()">
                                    Schedule
									<Input type=radio name=optMode value="A2" class="FormElem" OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Select</td>
								<td class="FieldCell" width="359">
								<Input type="Button" name="ButAcHead" value="Select" class="ActionButton"  disabled OnClick="popAccList()" >
                                &nbsp;</td>
								<td class="FieldCell" width="233">
								<Input type="text" name="txtAcHead" size="56" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">FinYear</td>
										<td class="FieldCell" colspan="2" width="599">
											<select size="1" name="FinYear" class="FormElem">
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option>
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="187">Compute Mode</td>
								<td class="FieldCell" width="359">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>

								<td class="FieldCell">Hierarchy &nbsp;&nbsp;
								<Input type="text" name="txtHierarchy" size="5" class="formElem" value="" ></td>
								</tr>
                            <td colspan="3">

								<table border="0" cellpadding="0" cellspacing="0" width="776">
										<tr>
											<td valign="middle" class="ActionCell" width="770">
                                                <p align="center">
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "SaveXML()" >&nbsp;
												<Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick = "Del()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">
											</td>
										</tr>
									</table>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
							<tr>
								<td align="center" colspan="3" class="BottomPack" width="702">
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

