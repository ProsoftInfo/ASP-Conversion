<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ContraConfigurations.asp
	'Module Name				:	ACCOUNTS (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 19,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	ContraEntry.asp
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
dim sOrgId,sOrgName,objRs,objRs1,sQuery,sAccName,iToAccCode,sToAccName,iSno,objRs2
Dim sFromAccNo,sFromAccName
Dim iBookCode,iBookNumber,iRecordsCount


'sOrgId=Request.QueryString("OrgCode")
sOrgId = session("organizationcode")
sOrgName = Session("orgshortname")
'sFromAccNo=Request.QueryString("FromAcc")
'sFromAccNo = "12"	'For Test

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

'sQuery = "Select b.AccountDescription from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
'		  "where a.OUDefinitionID='"& sOrgId &"' and b.AccountHead=a.FromAccountHead and "&_
'		  " a.FromAccountHead = "& sFromAccNo &" Group by a.FromAccountHead,b.AccountDescription"
'		  'Response.Write sQuery
'objRs.Open sQuery,con
'if not objRs.EOF then
'	sFromAccName = trim(objRs(0))
'end if
'objRs.Close

'Response.Write "<p><font color=red>Query=select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,Acc_M_GLAccountHead b where a.OUDefinitionID='010101' and a.EligibleForContras=1 and a.AccountHead=b.AccountHead and a.SubLedger=0 and A.AmendmentExists = '0' Order By b.AccountDescription "

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Contra List</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML id="AccData"><Root/></XML>
<XML ID="BookData"><Book/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = { type: "contraConfiguration" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="DisplayBook()">
<form name="formname" action="">
<input type=hidden name="hOrgCode" value="<%=sOrgId%>">
<input type=hidden name="hFromHead" value="<%=sFromAccNo%>">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class=PageTitle>
			<%'=sOrgName%>Contra Configuration
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td align="center" colspan="3" class="MiddlePack" height="8">
		</td>
	</tr>

	<tr>
		<td align="center" width="5px" class="ClearPixel">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
		</td>
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%"  >
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" >
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>

							<!--<tr>
								<td class="FieldCell" align="center" width="5">
								</td>
								<td class="FieldCellSub" valign="Top">Map To Book

									<%	If 1= 2 Then
										sQuery = " Select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,"&_
												 " Acc_M_GLAccountHead b Where a.OUDefinitionID='"& sOrgId &"' and a.EligibleForContras=1 and "&_
												 " a.AccountHead=b.AccountHead and a.SubLedger=0 and A.AmendmentExists = '0' and a.AccountHead<>"& sFromAccNo &" and a.AccountHead Not in ("&_
												 " select a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
												 " where a.OUDefinitionID='"& sorgID &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& sFromAccNo  &") "&_
												 " and a.AccountHead in (Select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead = b.AccountHead and Useable=0)"&_
												 " Order By b.AccountDescription "
												' Response.Write sQuery
										objRs.Open sQuery,con
										if not objRs.EOF then
											Response.Write "<input type=hidden name=hToAccHead value='Y'>"
											'Response.Write "Map To Book  : "
											Response.Write "<Select name=selToAccHead class=FormElem size=5 multiple>"

											do while not objRs.EOF
												Response.Write "<option value="& trim(objRs(0)) &">"&trim(objrs(2))&"</option>"
												objRs.MoveNext
											loop
											Response.Write "</Select>"
										else
											Response.Write "<input type=hidden name=hToAccHead value='N'>"
											Response.Write "Map To Book : No Books Available For Mapping"
										end if
										objRs.Close
										End IF
									%>

								</td>
							</tr>-->
							<%
							sQuery = " Select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,"&_
									 " Acc_M_GLAccountHead b Where a.OUDefinitionID='"& sOrgId &"' and a.EligibleForContras=1 and "&_
									 " a.AccountHead=b.AccountHead and a.SubLedger=0 and A.AmendmentExists = '0' and a.AccountHead<>"& sFromAccNo &" and a.AccountHead Not in ("&_
									 " select a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
									 " where a.OUDefinitionID='"& sorgID &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& sFromAccNo  &") "&_
									 " and a.AccountHead in (Select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead = b.AccountHead and Useable=0)"&_
									 " Order By b.AccountDescription "
									' Response.Write sQuery
							%>
							<!--<tr>
								<td align="center" height=10 colspan="3">
									<Table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="FieldcellSub" valign="top">Map To Book</tD>
											<td class="fieldcell" align="Left">
											<%
											If 1 = 2 Then
											objRs.Open sQuery,con
											if not objRs.EOF then
												Response.Write "<input type=hidden name=hToAccHead value='Y'>"
												'Response.Write "Map To Book  : "
												Response.Write "<Select name=selToAccHead class=FormElem size=5 multiple>"

												do while not objRs.EOF
													Response.Write "<option value="& trim(objRs(0)) &">"&trim(objrs(2))&"</option>"
													objRs.MoveNext
												loop
												Response.Write "</Select>"
											else
												Response.Write "<input type=hidden name=hToAccHead value='N'>"
												Response.Write "Map To Book : No Books Available For Mapping"
											end if
											objRs.Close
											End IF
										%>
											</td>
										</tr>
									</Table>
								</td>
							</tr>-->
							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
                               <td class="MiddlePack" colspan="4" width="100%">

								<Table border="0" cellpadding="0" cellspacing="0" width="100%" class="TableOutlineOnly">
									<tr>
										<td class="fieldcell" colspan="4">&nbsp;</td>
									</tr>
									<tr>
										<td class="FieldCellSub" valign="top" width="25%"> Select From Book</td>
										<td class="FieldCellSub" width="25%">
									         <select size="10" name="selFormHead" class="FormElem" onChange="popToHead()">
									         </select>
										</td>
										<td class="FieldCell" valign="top" width="25%"> Select To Book</td>
										<td class="FieldCellSub" width="25%">
										    <select size="10" name="selToHead" class="FormElem" Multiple>
										    </select>
										</td>
									</tr>
								</Table>
							</td>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>

							<tr>
								<td align="center" height="5px" colspan="4">
								</td>
							</tr>

							<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td align="center" class="BottomPack" colspan="4">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="ActionCell">
                                                    <input type="button" value="Save" name="B2" class="ActionButton" onclick="CheckSubmit()">&nbsp;
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton" >
											</td>
										</tr>
									</table>
								</td>
							</tr>

							<tr>
								<td align="center" height="5px" colspan="3">
								</td>
							</tr>

							<tr>
								<td align="center" width="5px">
								</td>
								<td valign="top" width="100%">

									<DIV class="frmbody" id="frm1" style="width: 585; height:150">
                                        <table id="tblMap"  border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" width="10" Rowspan="2">S.No.</td>
												<td class="ExcelHeaderCell" Rowspan="2">
													<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15px" height="15px" onClick="DelMapBook()">
												</td>
												<td class="ExcelHeaderCell" colspan="2" width="100%">Books Already Mapped</td>
                                            </tr>
                                            <tr>
												<td class="ExcelHeaderCell" width="50%">From</td>
												<td class="ExcelHeaderCell" width="50%">To</td>
                                            </tr>
<%
		iSno=0
		sQuery="select b.AccountDescription,a.ToAccountHead,a.FromAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
				"where a.OUDefinitionID='"&sorgID&"' and b.AccountHead=a.FromAccountHead "
				'and a.FromAccountHead="& sFromAccNo

		'Response.Write sQuery
		with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with

		set objRs.ActiveConnection = nothing

		set sAccName=objRs(0)
		set iToAccCode=objRs(1)
		set sFromAccNo = objrs(2)

		if not objRs.EOF then
			do while not objRs.EOF
			iRecordsCount = 0
				iSno=cint(iSno)+1

				sQuery="select AccountDescription from Acc_M_GLAccountHead "&_
						"where AccountHead="&iToAccCode
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing

				if not objRs1.EOF then
					sToAccName=objRs1(0)
				end if
				objRs1.Close

				sQuery="select AccountDescription from Acc_M_GLAccountHead "&_
						"where AccountHead="&sFromAccNo
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing

				if not objRs1.EOF then
					sFromAccName =objRs1(0)
				end if
				objRs1.Close

				sQuery = "Select BookCode,BookNumber from Acc_R_ApplicableAccountHeads where BookAccountHead = "& iToAccCode
				'Response.Write sQuery
				objRs1.Open sQuery,con
				if not objRs1.EOF then
					iBookCode = objRs1(0)
					iBookNumber = objrs1(1)
					sQuery = "Select Count(CreatedTransNo) from ACC_T_CreatedVoucherheader where BookCode = "& iBookCode  &" and BookNumber = "& iBookNumber
					'Response.Write sQuery
					objRs2.Open sQuery,con
					if not objRs2.EOF then
						iRecordsCount = objRs2(0)
					end if
					objRs2.Close
				end if
				objRs1.Close

			'	Response.Write "iRecordsCount = "& iRecordsCount

%>
                <tr>
					<td class="ExcelSerial" align="center"><%=iSno%></td>
					<td class="ExcelDisplayCell" align="center">
					<% IF iRecordsCount=0 THEN %>
						<input type="checkbox" name="chkBox<%=iSno%>" class=FormElem value="<%=iToAccCode%>">
					<% Else%>
						<input type="checkbox" name="chkBox<%=iSno%>" class=FormElem value="<%=iToAccCode%>" disabled>
					<% End IF%>
					</td>
					<td class="ExcelDisplayCell"><%=sFromAccName%></td>
					<td class="ExcelDisplayCell"><%=sToAccName%></td>
                </tr>
<%
			objRs.MoveNext
			loop
		end if
%>
	<input type="hidden" name="hRowCnt" value="<%=iSno%>">
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
							<!--<tr>
								<td align="center" width="5px" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                            <p align="center">
                                                                <input type="button" value="Close" name="B7" onclick="window.close()" class="ActionButton" >&nbsp;
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>-->
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
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
