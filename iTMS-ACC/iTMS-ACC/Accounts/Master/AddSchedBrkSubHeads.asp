<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedBrkSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Kumar K A
	'Created On					:	Dec 23 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	SchBreakupSetUp.asp
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
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera
	Dim ShSubID,ShSubSubID,sInsDate
	Set objrs1 = Server.createObject("ADODB.Recordset")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	sInsDate = Request("InsDate")
	
	IF CStr(sSchedNo) = "" Then
		sSchedNo = Request("sschedno")
	End IF
	If ShSubID = "" Then ShSubID = 0
	If ShSubSubID ="" Then ShSubSubID = 0
	
	sql = "select scheduleheading from dbo.acc_m_schdsetupheads where scheduleID ='"&sSchedNo&"'"
	with objrs1
		.CursorLocation = 3
		.CursorType =3
		.ActiveConnection = con
		.Source = sql
		.Open 
	end with
	if not Objrs1.EOF then sHead = objrs1(0)
	Objrs1.Close
		
%>
<script type="application/xml" data-itms-xml-island="1" ID="XmlData">
<Root>
	<Details OrgID="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="">
	</Details>
</Root>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Root/></script>

<script type="application/xml" data-itms-xml-island="1" ID="TempData">
	<Root/>
</script>

<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "scheduleBreakupSubHeadsPopup" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" >
<input type=hidden name="sUnit" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sFinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hAccHead" value="0">
<input type=hidden name="ShSubID" value="0">
<input type=hidden name="ShSubSubID" value="0">
<input type=hidden name="hInsDate" value="<%=sInsDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Schedule SubHeads
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="685" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr> <td class="FieldCellSub" width="120">Schedule Name<space><space><space></td>
							<td class="FieldCellSub" colspan="2"><%=sHead%></td>
							</tr>
                            	<td class="FieldCellSub" width="120">Add/Modify Level</td>
										<td class="FieldCell" colspan="2">
											<Input type="Hidden" name="optLevel" value="Lev1" class="FormElem" onclick="LevelFun()">
											<Input type="radio" name="optLevel" value="Lev2" class="FormElem" Checked onload="LevelFun()" onclick="LevelFun()">Level1
											<Input type="radio" name="optLevel" value="Lev3" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="120">Select Level 1</td>
										<td class="FieldCell">
											<select size="1" name="sel1" class="FormElem" onChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select SubHeadingName,ScheduleSubID,ScheduleSubSubID From Vw_Acc_SchSetup "&_
												  "Where EntryType = 'S' and OrganisationCode = '"&sOrgId&"' and FinYear = '"&sFinyr&"' and ScheduleID = "&sSchedNo &_ 
												  "Order By SubHeadingName"
										'Response.Write sql
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF
												
											%>
											<option value="<%=Objrs1(1)&"-"&Objrs1(2)%>"><%Response.Write (Left(Trim(Objrs1("SubHeadingName")),50)) %></b></option>
											<%Objrs1.MoveNext 
											loop
											Objrs1.Close
											%>
											</select>
										</td>
										<td class="FieldCell">
                                            <input type="Hidden" name=txtLev1 size=25 class="Formelem" align="Right" disabled>
										</td>
									</tr>
								  <tr>
								<td class="FieldCellSub" width="120">Select Level 2</td>
										<td class="FieldCell">
										<select size="1" name="sel2" class="FormElem" onchange="setlevelfun1()" >
										<OPTION Value="0">Select</option>
						
									</select>
								</td>
									<input type=hidden name=hSubHeadName value="">
										<td class="FieldCell">
                                        <input type=text name=txtLev2 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Select Level 3</td>
										<td class="FieldCell">
										<select size="1" name="sel3" class="FormElem" onchange="Setlevelfun2()" disabled>
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>
									
									</td>
									<input type=hidden name=hSubHeadName3 value="">
										<td class="FieldCell">
                                        <input type=text name=txtLev3 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								
								
								
								<tr>
								<td class="FieldCellSub" width="120">Mode</td>
								<td class="FieldCell" colspan="2">
									<Input type=radio name=optMode value="D" class="FormElem" checked OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A2" class="FormElem" OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Select A/c Head</td>
								<td class="FieldCell">
								<Input type="Button" name="ButAcHead" value="A/cHead" class="ActionButton" disabled OnClick="popAccList()" >
                                &nbsp;</td>
								<td class="FieldCell">
								<Input type="text" name="txtAcHead" size="25" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">FinYear</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="FinYear" class="FormElem" disabled>
											<Option>Select</Option>
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option> 
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="120">Compute Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>
								<td class="FieldCell">Hierarchy &nbsp;&nbsp; 
								<Input type="text" name="txtHierarchy" size="5" class="formElem" value="" ></td>
								</tr>
							
                            <td colspan="3">
								
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> 
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "CheckSubmit()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">&nbsp;
                                                <Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick ="Del()">
											</td>
										</tr>
									</table>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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

