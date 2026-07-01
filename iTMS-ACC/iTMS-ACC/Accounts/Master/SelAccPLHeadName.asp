<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SelAccPLHeadName.asp
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
Dim sOrgID,sBookCode,iPlHead,iPlSubHd,iPlSubSubHd
Dim iSchID,iSchSubID,iSchSubSubId,iBkId,sBkHead,sBkSubHead
Dim iBkSchdId,iBkSchdSubId,iBkSchdSubSubId,iBkPara,sTemp,sCode,sTemp2

Set rs1 = server.CreateObject("ADODB.Recordset") 
iPlHead = Request("PLHead")
iPlSubHd = Request("PLSubHd")
iPlSubSubHd = Request("PLSubSubHd")
sOrgID = Request("orgId")

sTemp = Split(iPlSubHd,",")
sTemp2 = Split(iPlSubSubHd,",")

iPlSubHd = sTemp(0)
iPlSubSubHd = sTemp2(0)

IF CStr(iPlSubHd) = "A" Then
	iPlSubHd = 0
End IF

IF CStr(iPlSubSubHd) = "A" Then
	iPlSubSubHd = 0
End IF

IF CStr(iPlHead) = "A" Then
	iPlHead = 0
End IF



'Response.Write sOrgID &" " & iPlHead &" " & iPlSubHd &" " & iPlSubSubHd

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = {
	type: "scheduleSelection",
	saveUrl: "XMLSave.asp?Name=SchedPLBrkSubHeads&Mod=Acc",
	returnValue: "Y"
};
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</head>
<BODY>
<form method="POST" name="formname">
<input type="hidden" name="hAcclist" value="">
<input type="hidden" name="hPass" value = "<%'=iBkPara%>">
<p align="center" class=PageTitle>Schedule Sub Heads </p>
          
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

	<tr>
	<td>&nbsp;</td>
	<td class="FieldCell" colspan="7" align="center">
		<select size="12" name="SelName1" STYLE="WIDTH=400PX" class="FormElem" multiple>
		<%
			Dim ShDesc,ShValues,i
			i = 1
			'sQuery = "SELECT ScheduleHeading, ScheduleID, ScheduleNumber" &_
			'		" FROM dbo.Acc_M_Schdsetupheads ORDER BY ScheduleNumber"
			
			'sQuery = "Select ScheduleID,ScheduleSubID,ScheduleSubSubID,'ScheduleID '+ Cast(ScheduleNumber As Varchar) +'-'+ScheduleHeading+'-'+SubHeadingName" &_
			'		" From VwAccSchSetup Where ApplicableFor = 'P' Order By ScheduleNumber,ScheduleHeading,SubHeadingName"
			
			sQuery = "Select ScheduleID,ScheduleSubID,ScheduleSubSubID,'ScheduleID '+ Cast(ScheduleNumber As Varchar) +'-'+ScheduleHeading+'-'+SubHeadingName "&_
					 "From Vw_Acc_SchSetup Where ApplicableFor = 'P' and Cast(ScheduleID As VarChar)+''+Cast(ScheduleSubID As VarChar)+''+Cast(ScheduleSubSubID As Varchar) "&_
					 "Not in (Select Cast(ScheduleID As VarChar)+''+Cast(ScheduleSubID As VarChar)+''+Cast(ScheduleSubSubID As Varchar) "&_
					 "From Acc_T_PLAcDetail) Order By ScheduleNumber,ScheduleHeading,SubHeadingName "
					
			With rs1
				.ActiveConnection = Con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			End With
			Set rs1.ActiveConnection = Nothing
			Do While Not rs1.EOF
				'ShDesc = "Schedule "&i&" : "&rs1(0)
				
				
				
				ShDesc =rs1(3) ' "Schedule "&rs1(0)&" : "&
				ShValues = rs1(0)&"-"&rs1(1)&"-"&rs1(2)
		%>
				<option Value="<%=ShValues%>"><%=ShDesc%></Option>
		<%
				rs1.MoveNext
				i = i+1
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
		
		<tr>
	<td>&nbsp;</td>
	<td class="FieldCell" colspan="7" align="center">
			<select size="12" name="SelName2" STYLE="WIDTH=400PX" class="FormElem" multiple>		
		<%
			
			
			sQuery = "Select ScheduleID,ScheduleSubID,ScheduleSubSubID,'ScheduleID '+ Cast(ScheduleNumber As Varchar) +'-'+ScheduleHeading+'-'+SubHeadingName "&_
					 "From Vw_Acc_SchSetup Where ApplicableFor = 'P' and Cast(ScheduleID As VarChar)+''+Cast(ScheduleSubID As VarChar)+''+Cast(ScheduleSubSubID As Varchar) "&_
					 "in (Select Cast(ScheduleID As VarChar)+''+Cast(ScheduleSubID As VarChar)+''+Cast(ScheduleSubSubID As Varchar) "&_
					 "From Acc_T_PLAcDetail Where OrganisationCode = '"&sOrgID&"' and PLHeadID = "&iPlHead&" and PLSubID = "&iPlSubHd&" and PLSubSubID = "&iPlSubSubHd&") "
					
			With rs1
				.ActiveConnection = Con
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.Open
			End With
			Set rs1.ActiveConnection = Nothing
			Do While Not rs1.EOF
				'ShDesc = "Schedule "&i&" : "&rs1(0)
				ShDesc =rs1(3) ' "Schedule "&rs1(0)&" : "&
				ShValues = rs1(0)&"-"&rs1(1)&"-"&rs1(2)
		%>
				<option Value="<%=ShValues%>"><%=ShDesc%></Option>
		<%
				rs1.MoveNext
				
			Loop
			rs1.Close
		%>
		</select>
			</select>
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
                   <input type="Reset" value="Cancel"  name="ButRes"  class="ActionButton" onclick = "Window.close()">
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
