<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SelAccHeadName.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Maheshwari .S
	'Created On					:	Oct 07 2006
	'Modified On				:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include File="../../include/sessionVerify.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include File="../../include/purpopulate.asp" -->


<%

Dim sQuery,rs1,rs2,rs3
Dim sOrgID,sBookCode
Dim iSchID,iSchSubID,iSchSubSubId,iBkId,sBkHead,sBkSubHead
Dim iBkSchdId,iBkSchdSubId,iBkSchdSubSubId,iBkPara,sTemp,sCode

Set rs1 = server.CreateObject("ADODB.Recordset") 
'Set rs2 = server.CreateObject("ADODB.Recordset") 
'Set rs3 = server.CreateObject("ADODB.Recordset") 
iBkPara = Request("sTemp")
'Response.Write "Parameter=" & iBkPara

sTemp = Split(iBkPara,":")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script src="../../scripts/itms-modern-compat.js"></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/Selection.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "scheduleSelection", mode: "returnList" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</head>
<BODY>
<form method="POST" name="formname">
<input type="hidden" name="hAcclist" value="">
<input type="hidden" name="hPass" value = "<%=iBkPara%>">
<p align="center" class=PageTitle>Schedule Sub Heads </p>
          
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	
	
	<tr>
	<td>&nbsp;</td>
	<td class="FieldCell" colspan="7" align="center">
		<select size="8" name="SelName1" STYLE="WIDTH=300PX" class="FormElem" multiple>
		<%
			sQuery = "Select Distinct BreakupHeading,BreakupID,ScheduleID,ScheduleSubID, "&_
					 "ScheduleSubSubID from Vw_Acc_SchBreakSetup where ScheduleID = "&sTemp(0)&" and  "&_
					 "ScheduleSubID = "&sTemp(1)&" and ScheduleSubSubID = "&sTemp(2)&" "&_
					 "and Useable = 'N' Order By BreakupHeading "
			With rs1
				.ActiveConnection = Con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			End With
			Set rs1.ActiveConnection = Nothing
			Do While Not rs1.EOF
				iSchID = rs1("ScheduleID") 
				iSchSubID = rs1("ScheduleSubID") 
				iSchSubSubId = rs1("ScheduleSubSubID") 
				iBkId = rs1("BreakupID")
				sCode = iSchID&"-"&iSchSubID&"-"&iSchSubSubId&"-"&iBkId
		%>
				<option Value="<%=sCode%>"><%=rs1("BreakupHeading")%></Option>
		<%
				rs1.MoveNext
			Loop
			rs1.Close
			
		%>
		</select>
	</td>
		<td colspan="1">&nbsp;</td>
		<td>
		<td class="FieldCell" colspan="7" align="center">
			<select size="8" name="SelName2" STYLE="WIDTH=300PX" class="FormElem" multiple>		
		<%
			sQuery = "Select Distinct BreakupHeading,BreakupID,ScheduleID,ScheduleSubID, "&_
					 "ScheduleSubSubID from Vw_Acc_SchBreakSetup where ScheduleID = "&sTemp(0)&" and  "&_
					 "ScheduleSubID = "&sTemp(1)&" and ScheduleSubSubID = "&sTemp(2)&" "&_
					 "and Useable = 'Y' Order By BreakupHeading "
			
			With rs1
				.ActiveConnection = Con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			End With
			Set rs1.ActiveConnection = Nothing
			Do While Not rs1.EOF
				iSchID = rs1("ScheduleID") 
				iSchSubID = rs1("ScheduleSubID") 
				iSchSubSubId = rs1("ScheduleSubSubID") 
				iBkId = rs1("BreakupID")
				sCode = iSchID&"-"&iSchSubID&"-"&iSchSubSubId&"-"&iBkId
		%>
				<option Value="<%=sCode%>"><%=Replace(rs1("BreakupHeading"),":"," ")%></Option>
		<%
				rs1.MoveNext
			Loop
			rs1.Close
			
		%>
			</select>
		</td>
		</tr>
		<tr>
		<td colspan="16" align="center">
			<input type="Button" name="ButAdd"  class="AddButton" value="Add >>" onclick = "addclick('SelName2','SelName1','ButRem')">
			<input type="Button" name="ButRem"  class="AddButton" value="<< Remove" onclick = "removeclick('SelName2','SelName1','ButRem')">
		</td>
		</tr>
	</table>
<tr>
<td align="center">
	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td valign="middle" class="ActionCell">
                <p align="center"> 
                   <input type="Button" value="Submit" name="ButSub"  class="ActionButton" onClick="SubmitFun()">
                   <input type="Reset" value="Cancel"  name="ButRes"  class="ActionButton" onclick = "window.close()">
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
 </td>
</tr>
</table>
</form>
</BODY>
</HTML>
