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
	'Program Name				:	ContraListPopup.asp
	'Module Name				:	ACCOUNTS (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 10,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim sOrgId,sOrgName,objRs,objRs1,sQuery,sAccName,iToAccCode,sToAccName,iSno,objRs2
Dim sFromAccNo,sFromAccName
Dim iBookCode,iBookNumber,iRecordsCount


sOrgId=Request.QueryString("OrgCode")
sFromAccNo=Request.QueryString("FromAcc")

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")

sQuery = "Select b.AccountDescription from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
		  "where a.OUDefinitionID='"& sOrgId &"' and b.AccountHead=a.FromAccountHead and "&_
		  " a.FromAccountHead = "& sFromAccNo &" Group by a.FromAccountHead,b.AccountDescription"
		  'Response.Write sQuery
objRs.Open sQuery,con
if not objRs.EOF then
	sFromAccName = trim(objRs(0))
end if
objRs.Close

sQuery = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = '"& sOrgId &"'"
objRs.Open sQuery,con
if not objRs.EOF then
	sOrgName = trim(objRs(0))
end if
objRs.Close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Contra List</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="AccData"><Root/></script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "contraList" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form name="formname">
<input type=hidden name="hOrgCode" value="<%=sOrgId%>">
<input type=hidden name="hFromHead" value="<%=sFromAccNo%>">
<table border="0" cellspacing="0" cellpadding="0" class="popuptable">
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">
			<%=sOrgName%></P>
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
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td class="FieldCellSub">
								<p align=center><B>Contra Entries For <%=sFromAccName%></B></p>
								</td>
							</tr>

							<tr>
								<td align="center" width="5">
								</td>
								<td class="FieldCellSub" height=10>
								</td>
							</tr>

							<tr>
								<td align="center" width="5">
								</td>
								<td class="FieldCellSub" valign=top>
									<%
										sQuery = "Select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,"&_
												 "Acc_M_GLAccountHead b Where a.OUDefinitionID='"& sOrgId &"' and a.EligibleForContras=1 and "&_
												 "a.AccountHead=b.AccountHead and a.SubLedger=0 and A.AmendmentExists = '0' and a.AccountHead<>"& sFromAccNo &" and a.AccountHead Not in ("&_
												 " select a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
												"where a.OUDefinitionID='"& sorgID &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& sFromAccNo  &") "&_
												" and a.AccountHead in (Select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead = b.AccountHead and Useable=0)"&_
												"Order By b.AccountDescription "
												' Response.Write sQuery
										objRs.Open sQuery,con
										if not objRs.EOF then
											Response.Write "<input type=hidden name=hToAccHead value='Y'>"
											Response.Write "Map To Book  : "
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
									%>

								</td>
							</tr>

							<tr>
								<td align="center" height=10 colspan="3">
								</td>
							</tr>

							<tr>
								<td align="center" class="BottomPack" colspan="3">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >&nbsp;
                                                    <input type="reset" value="Reset" name="B1" class="ActionButton" >
											</td>
										</tr>
									</table>
								</td>
							</tr>

							<tr>
								<td align="center" height=10 colspan="3">
								</td>
							</tr>

							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">

									<DIV class=frmBody id=frm1 style="width: 415; height:150">
                                        <table id="tblMap"  border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">
													<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DelMapBook()">
												</td>
												<td class="ExcelHeaderCell" align="center" width="100%">Books Already Mapped</td>
                                            </tr>
<%
		iSno=0
		sQuery="select b.AccountDescription,a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
				"where a.OUDefinitionID='"&sorgID&"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead="& sFromAccNo
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
					<td class="ExcelDisplayCell"><%=sToAccName%></td>
                </tr>
<%
			objRs.MoveNext
			loop
		end if
%>
	<input type=hidden name="hRowCnt" value="<%=iSno%>">
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
							</tr>
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
