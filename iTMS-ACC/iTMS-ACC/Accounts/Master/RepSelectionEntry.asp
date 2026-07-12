<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	RepSelectionEntry.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 04,2012
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/populate.asp" -->
<%
    Dim rsTemp
    Dim sAgentName,sAgentShortName,sAgentAddress1,sAgentAddress2,sPartyCode,sQuery,sCity,sPinCode
    Dim sAgentAddress4,sAgentPhone,sAgentFax,sAgentMailID,sExternalOrInternal,sAgentType,sArrTemp
    Dim iAgentEntryID,iRepAreaCode,sOrgID
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    sPartyCode = Request("PartyCode")
    sOrgID = Session("organizationcode")
    
    Response.Write "<font colo=red>"
    
    iAgentEntryID = 0
    sQuery = "Select RepAreaCode,RepAgentEntryID from APP_R_OrgParty where PartyCode = "& sPartyCode &" and OUDefinitionID="& pack(sOrgID)
    'Response.Write sQuery 
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        iRepAreaCode = rsTemp(0)
        iAgentEntryID = rsTemp(1)
    end if
    rsTemp.Close 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self"></base>
<script type="application/xml" data-itms-xml-island="1" id="RepAreaData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="AgentData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ConPerForArea"><Root></Root></script>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<script SRC="../../scripts/rolloverout.js"></SCRIPT>
<script SRC="../../scripts/trim.js"></SCRIPT>

<script>
window.__itmsPopupCompat = { type: "representativeSelection" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="post" name="formname" action="">
<input type="hidden" name="hPartyCode" value="<%=sPartyCode%>">
<input type="hidden" name="hAreaCode" value="<%=iRepAreaCode%>">
<input type="hidden" name="hAgentEntryID" value="<%=iAgentEntryID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="middle" class=PageTitle height="20"><p align="center">Representative Selection</p>
		</td>
    </tr>
	<tr>
		<td align="middle" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="middle" colspan="2" class="MiddlePack">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
                                <td class="FieldCellSub">Area&nbsp;</td>
					            <td class="FieldCell">
					            <select name=SelRepArea  class="FormElem" onchange="PopulateContPerson(this)">
					                <option value="0">Area</option>
					            </select>&nbsp;&nbsp;
					        </td>
                           </tr>
                           <tr>
                                <td class="FieldCellSub">Representative&nbsp;</td>
					            <td class="FieldCell">
					            <select name="SelContPerson" class="FormElem">
					                <option value="0">Name</option>
					                    <%
					                    sQuery = "Select AgentEntryID,LocationCode,isNull(ContactPersonName,'') from APP_M_AgentLocations"
					                    rsTemp.Open sQuery,con
					                    Do while Not rsTemp.EOF
									        If rsTemp(2) <> "" Then
									        %>
										        <option value="<%=rstemp(0)%>|<%=rstemp(1)%>"><%=rstemp(2)%></option>
									        <%
									        End If
									        rsTemp.MoveNext
					                    Loop
					                    rsTemp.Close
					                    %>
					            </select>
					        </td>
					        </tr>
					        <tr>
								<td align="middle" colspan="2" class="MiddlePack">
									<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan=2>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td valign="middle" class="ActionCell">
											<p align="center">
											    <input type="button" value="Create Representative" onClick="CreateRep()" name="btnCreate" class="ActionButtonX" tabindex="3" > 
                                                <input type="button" value="Save" onClick="checkSubmit()" name="B2" class="ActionButton" tabindex="3" > 
                                                <input type="button" value="Close" name="B3" class="ActionButton" tabindex="3" onclick="window.close()"> 
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td align="middle" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</TD>
				</TR>
			</TABLE>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
