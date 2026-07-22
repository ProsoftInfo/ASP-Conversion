<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IntRcptRefNoSel.asp
	'Module Name				:	Inventory (Transcation)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Jan 24,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<script type="application/xml" data-itms-xml-island="1" id="TempData"><Root></Root></script>
<title>iTMS-Internal Receipt Reference Selection</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/printwindow.js"></script>
<script LANGUAGE=javascript SRC="../scripts/intRcptRefNoSel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
Dim sRefType,sOrgCode
Dim Query,rsObj
set rsObj = server.createObject("ADODB.Recordset")
                        
sRefType = Request.QueryString("RefType")
sOrgCode = Request.QueryString("OrgID")
%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
	<form method="POST" name="formname" action="<%=Request.ServerVariables("SCRIPTNAME")%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Reference Selection
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>



<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
    <table>
        <tr>
            <td class="FieldCell">
                Reference No
            </td>
            <td>
                <select id="SelReference" name="selReference" class="formelem">
                    <option value="S">Select</option>
                    <%
                        if trim(sRefType)="I" then
                            Query = "Select IssueEntryNo,isNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103) from INV_T_MaterialIssueHeader Order By IssueDate Desc"
                            rsObj.open Query,con
                            if not rsObj.eof then
                                do while not rsObj.eof 
                                    %>
                                        <option value="<%=rsObj(0)%>"><%=rsObj(1)%>-<%=rsObj(2)%></option>
                                    <%
                                    rsObj.movenext
                                loop
                            end if
                            rsObj.close
                        end if
                    %>
                </select>
            </td>
        </tr>
    </table>
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>
<tr>
<td align="center" class="MiddlePack" colspan="3">
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
<input type="button" value="Done" name="B11" class="ActionButton" tabindex="4" onclick="CheckSubmit()">
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
</body>
</html>
