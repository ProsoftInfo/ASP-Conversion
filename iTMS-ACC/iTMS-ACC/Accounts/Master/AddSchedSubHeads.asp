<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Maheshwari S.
	'Created On					:	Dec 19 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	SchSetUp.asp
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS - Accounts - Add Schedule Sub Heads</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<% sInsDate = Request("InsDate") %>
<XML ID="XmlData">
<Root>
	<Details OrgID="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="" InsDate="<%=sInsDate%>">
	</Details>
</Root>
</XML>
<XML id="OutData"><Root/></xml>
<XML ID="TempData"><Root/></XML>
<%
	Dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName,objfs
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera,sInsDate
	Set objrs1 = Server.createObject("ADODB.Recordset")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	
%>
	

<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "scheduleSubHeadsPopup" };
</script>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" >
<input type=hidden name=hSubHeadName value=""><input type=hidden name="hOrgId" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sfinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hInsDate" value="<%=sInsDate%>">

<input type=hidden name="hAccHead" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Add Schedule Sub Heads
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="720">
				<TR>
					<TD class=TabBodyWithTopLine>
                        <table border="0" width="100%" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%">
						<table border="0" cellpadding="0" cellspacing="0">
                             <tr>
								<td class="FieldCellSub">Select Schedule</td>
                                <td class="FieldCellSub">
								
									<%
								
									sql = "SELECT ScheduleID, ScheduleHeading, FinYear FROM dbo.Acc_M_SchdSetupHeads Where ScheduleID = "&sSchedNo
									with Objrs1
										.CursorLocation = 3
										.CursorType = 3
										.ActiveConnection = con
										.Source = sql
										.Open
									End With
									Set Objrs1.ActiveConnection = Nothing 
										
									If not Objrs1.EOF Then
									%>					
											<input type=text name="selSch1" size=25 class="FormelemRead" align="Right" value="<%Response.Write(Objrs1(1))%>" readonly>
											<input type="Hidden" name="SelSch" size=25 class="Formelem" align="Right" value="<%Response.Write(Objrs1(0))%>">
									<%
									
									End IF
									Objrs1.Close
									
									%>
									
									</select>
									</td>
								  </tr>
							  <tr>
								<td class="FieldCellSub" valign="top">Add/Modify Level</td>
										<td class="FieldCell" valign="top">
											<Input type="radio" name="optLevel" value="Lev3" class="FormElem"  onclick="LevelFun()" Checked>Level1
											<Input type="radio" name="optLevel" value="Lev4" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub">Select Level 1</td>
										<td class="FieldCellSub">
											<select size="1" name="selLevel1" class="FormElem" OnChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select Distinct ScheduleSubID,ScheduleSubSubID,isNull(SubHeadingName,'') "&_
													" SubHeadingName,EntryType,computemode,Hierarchy From Vw_Acc_SchSetup Where  OrganisationCode = '"&sOrgId&"' and "&_
													" FinYear = '"&sFinyr&"' and scheduleID = "&sSchedNo&" and schedulesubsubid = 0 Order By Hierarchy "
						
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF 
											%>
											<option value="<%=Objrs1(0)&"-"&Objrs1(1)&"-"&Objrs1("EntryType")&"-"&Objrs1("computemode")%>"><%Response.Write (Left(Objrs1("SubHeadingName"),70)) %></b></option>
											<%Objrs1.MoveNext 
											loop
											Objrs1.Close
											
											%>
											<option Value="A">Add New</option>
											</select>
										</td>
									</tr>
									
									<tr>
									
									<td class="FieldCellSub">&nbsp;</td>
									<td class="FieldCellSub">
										<input type=text name=txtLev1 size=25 class="Formelem" align="Right" >
									</td>
									</tr>
									
								  <tr>
								<td class="FieldCellSub">Select Level 2</td>
										<td class="FieldCellSub">
										<select size="1" name="selLevel2" class="FormElem" disabled onchange="setlevelfun1()">
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>
									
									</td>
								</tr>
								
								<tr>
									<td class="FieldCellSub">&nbsp;</td>
									<td class="FieldCellSub">
										<input type=text name=txtLev2 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								
								<tr>
								<td class="FieldCellSub">Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optMode value="D" class="FormElem" OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A3" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A4" class="FormElem" OnClick="ModeFun()">
                                    Schedule
									<Input type=radio name=optMode value="A5" class="FormElem" checked OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub">Select A/c Head</td>
								<td class="FieldCellSub">
								<Input type="Button" name="ButAcHead" value="A/cHead" class="ActionButton"  disabled OnClick="popAccList()" >
                                &nbsp;<Input type="text" name="txtAcHead" size="25" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub">FinYear</td>
										<td class="FieldCellSub">
											<select size="1" name="FinYear" class="FormElem">
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option> 
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub">Compute Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>
								
								</tr>
		
						</table>
                            </td>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                          </tr>
                          <tr>
                            <td colspan="3" class="MiddlePack"></td>
                          </tr>
                          <tr>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> 
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "SaveXML()" >&nbsp;
                                                <Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick = "Del()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">
											</td>
										</tr>
									</table>
                            </td>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                          </tr>
                          <tr>
                            <td colspan="3" class="MiddlePack"></td>
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

