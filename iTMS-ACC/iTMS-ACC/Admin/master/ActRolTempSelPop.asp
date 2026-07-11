<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ActRolTempSelPop.asp
	'Module Name				:	Admin (Role)
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Activity Role Mapping Popup</title>
<base target="_self"/>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML id="ActivityData"><Root></Root></XML>
<XML id="RoleData"><Root></Root></XML>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/ActRolTempSelPopCompat.js"></SCRIPT>
<SCRIPT>
ITMSActRolTempSelPopCompat.install();
</SCRIPT>

<%
    Dim nAppCode,nProcessCode,nActCode
    
    nAppCode = Request.QueryString("AppCode")
	nProcessCode = Request.QueryString("ProcessCode")
	nActCode = Request.QueryString("ActCode")
%>
</head>
<body leftmargin="5" topmargin="0" marginheight="0" marginwidth="0" onload="Init()" >

	<form method="POST" name="formname" action="">
	<Input type="hidden" name="hItemRows" value="">
	<Input type="hidden" name="hLastSelectedPractice" value="">
	<Input type="hidden" name="hAppCode" value="<%=nAppCode%>">
	<Input type="hidden" name="hProcessCode" value="<%=nProcessCode%>">
	<Input type="hidden" name="hActCode" value="<%=nActCode%>">
		
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Role Activity Mapping
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				    <tr>
				        <td class="toppack"></td>
				    </tr>
				    <tr>
						<td class="TabBody">
						    <table border="0" cellpadding="0" cellspacing="1" width="100%">
						        <tr>
						            <td>    
						                <div style="width:470px;height:420px">
							                <table border="0" id="tblTempAct" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
					                            <tr>
					                                <td class="ExcelHeaderCell" align="center" rowspan="2" style="width:20px">S.No.</td>
					                                <td class="ExcelHeaderCell" align="center" rowspan="2">&nbsp;</td>
					                                <td class="ExcelHeaderCell" align="center">Activity Name</td>
					                            </tr>
					                            <tr>
					                                <td class="ExcelHeaderCell" align="center">Template Name</td>
					                            </tr>
							                </table>
							            </div>
						            </td>
						        </tr>
						    </table>
						</td>
					</tr>
					 <tr>
				        <td class="BottomPack"></td>
				    </tr>
					<tr>
					    <td>
					        <table width="100%">
					            <tr>
					                <td class="ActionCell" align="center">
					                    <input type="button" name="btnDone" value="Done" class="ActionButton" onclick="CheckSubmit()">
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
