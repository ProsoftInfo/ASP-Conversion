<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	BalScheduleSetup.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Maheshwari S.
	'Created On					:	Dec 16 2006
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
<%
	Dim sOrgId,Objrs,Root,iCtr
	Dim sql,sNo,sHead,sHiera,sApp,sFinyear,iSchId,oDOM,iNo,iCnt,iHiera,iSchno
	Dim SchHeading,ApplicableFor,sFinyr,sInsDate,sCallFrom
	
	Set  objrs = Server.createObject("ADODB.Recordset")
	Set oDOM  = Server.CreateObject("Microsoft.XMLDOM")	
	
	sOrgId = Request("sUnit")
	iSchId = Request("sSchID")
	sInsDate = Request("InsDate")
	sCallFrom = Request("CallFrom")
	
	'To check whether ScheduleNo is already exist in database or not
	'sql = "select ScheduleNumber,Count(0),Hierarchy,scheduleHeading from Acc_M_SchdSetupHeads where ScheduleID <> 0 "&_
	'		"group by ScheduleNumber,Hierarchy,scheduleHeading" 
		'Response.Write sql
	sql = "SELECT ScheduleID, OrganisationCode, ScheduleNumber, ScheduleHeading, Hierarchy, ApplicableFor, FinYear " &_
			" FROM dbo.Acc_M_SchdSetupHeads WHERE (ScheduleID = "&iSchId&")"
		objRs.Open sql,Con
		iCnt = 1
		Do while Not objrs.EOF
			iNo =  objRs(0) & "," & iNo 
			iCtr = iCtr + objRs(1) 
			iSchno = Objrs(2)
			iHiera = Objrs(4) '& "," & iHiera
			SchHeading = Objrs(3)
			sFinyear = Objrs(6)
			ApplicableFor =Objrs(5)
			'Response.Write ApplicableFor 
			'iCnt=iCnt+1
		Objrs.MoveNext 
		loop
		objRs.Close
		
	'Response.Write "iCtr="& iCtr & "<BR><BR>"
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>Add/Edit Schedule</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<XML ID="TempData">
	<Root/>
</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = { type: "balanceScheduleSetup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname">
<input type=hidden name="hOrgId" value="<%=sOrgId%>">
<input type=hidden name="hNo" value="<%=sNo%>">
<input type=hidden name="hName" value="<%=sHead%>">
<input type=hidden name="hHiera" value="<%=sHiera%>">
<input type=hidden name="hHierarchy" value="<%=iHiera%>">
<input type=hidden name="hApp" value="<%=sApp%>">
<input type=hidden name="hFinYr" value="<%=sFinYr%>">
<input type=hidden name="hino" value="<%=iNo%>">
<input type=hidden name="hiCtr" value="<%=iCtr%>">
<input type=hidden name="iSchId" value="<%=iSchId%>">
<input type=hidden name="hInsDate" value="<%=sInsDate%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Add Schedule    
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                             <tr>
								<td class="FieldCellSub" width="150">Schedule No</td>
                                <td class="FieldCell" width="110" colspan="4">
									<input type=text name=txtno size=10 class="Formelem" align="Left" Value="<%=iSchno%>" OnChange = "CheckVal()">
								</td>
							  </tr>
							  <tr>
								<td class="FieldCellSub" width="150">Schedule Heading</td>
										<td class="FieldCell" colspan="4">
											<Input type=text name="txtSchdHead" size=55 class="Formelem"  align="Right" Value="<%=SchHeading%>" OnChange = "CheckVal()" maxlength="200">
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub" width="150">Schedule Hierarchy</td>
										<td class="FieldCell" colspan="4">
											<Input type=text name="txtSchdHiera" size=5 class="Formelem" value="<%=iHiera%>" OnChange = "ChkVal()">
										</td>
								</tr>
								    <tr>
								<td class="FieldCellSub" width="150">Applicable For</td>
										<td class="FieldCell" colspan="4">
										<select size="1" name="selApp" class="FormElem">
										<OPTION>Select</option>
										<%'If ApplicableFor ="L" Then	%>
										<%If sCallFrom ="BS" Then	%>
											<OPTION Value="L" selected="L">Balance Sheet</option>
											<!--<OPTION value="P">P & L</option>-->
										<% Else %>
											<!--<OPTION  value="L">Balance Sheet</option>-->
											<OPTION Value ="P" selected="P">P & L</option>
										<%End If%>
									</select>
									</td>
								</tr>
							  <tr>
								<td class="FieldCellSub" width="150">Financial Period</td>
										<td class="FieldCell" colspan="2">
											<select size="1" name="FinYear" class="FormElem">
											
											<%
												Dim sQry,rs1,sFmPer,sToPer,Arr,Arr1,sFmYr,sToYr
																								
												Set rs1 = Server.CreateObject("ADODB.Recordset")
												sQry = "Select convert(DateTime,FromPeriod,103),convert(dateTime,ToPeriod,103) "&_
														" from Ms_FinancialPeriod where Active = 'Y' "
												Response.Write  sQry &"<BR><BR>"
												rs1.Open sQry,Con		
												If Not rs1.EOF Then
													sFmPer = rs1(0)
													sToPer = rs1(1)
												End If
												Arr= Split(sFmPer,"/")
												sFmYr = Arr(2)
												Arr1= Split(sToPer,"/")
												sToYr = Arr1(2)
												sFinYr = sFmYr &":"& sToYr
												
												Response.Write "From Date="& sFmPer &"<BR><BR>"
												Response.Write "To Date="& sToPer &"<BR><BR>"
												if sFmYr&"-"&sToYr = sFinyear Then %>
													<Option Selected="<%=sFinYr%>" ><%=sFmYr%> - <%=sToYr%></Option>
												<%End If%>
											<Option Value="<%=sFinYr%>" ><%=sFmYr%> - <%=sToYr%></Option>
											</select>
										</td>
								</tr>
		
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> 
                                                <input type="button" value="Save" name="B4" class="ActionButton" onClick="CheckSubmit()">
                                                 <input type="button" value="Delete" name="B5" class="ActionButton" OnClick="Del()">
                                                <input type="button" value="Close" name="B5" class="ActionButton" OnClick="Window.close()">
                                                
									<% oDOM.save server.MapPath("../temp/transaction/Schedule"&Session.SessionID&".xml")	%>			
												
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

