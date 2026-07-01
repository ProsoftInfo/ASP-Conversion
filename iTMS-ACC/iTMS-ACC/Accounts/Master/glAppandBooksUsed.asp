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
	'Program Name				:	glAppandBooksUsed.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	December 02,2003
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/accpopulate.asp"-->
<%
dim sOrgId,objRs,sQuery,sCallTy,Temparr,sSelVal,sTitle
dim sBookCode,iBookNo,sTransType,bFlag

sCallTy=Request("sTempValues")
Temparr = Split(sCallTy,"?")
sCallTy = Temparr(0)
sSelVal = Temparr(1)
'Response.Write sSelVal
Temparr = Split(sSelVal,",")
'sBookCode=Request("BookCode")

IF CStr(sCallTy) = "A" Then
	sTitle = "Application Used "
Else
	sTitle = "Frequently Book Used "
End IF


Set objRs = Server.CreateObject("ADODB.RecordSet")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE><%=sTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript">
window.__itmsPopupCompat = { type: "appBookUsed" };
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="CheckVal('<%=sSelVal%>')">
<form method="POST" name="formname" action="">
<div align="center">
  <center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popuptable">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
						<table border="0" cellpadding="0" cellspacing="0" width="100%" >
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="7">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel" height="2">
                                    &nbsp;
                                    <p>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                    </p>
								</td>
								<td valign="top">
                                    <div align="center">
                                      <center>
                                    <table cellpadding="0" cellspacing="1" border="0" class="ExcelTable">
 
<%
Dim sBookId,sBookLName,sAppID,sAppLName,iRecCount,iCtr,sCheckVal
IF CStr(sCallTy) = "A" Then 
	with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT APPLICATIONCODE,APPLICATIONNAME,APPLNSHORTNAME FROM MS_APPLICATIONS ORDER BY APPLICATIONNAME"
			.ActiveConnection = con
			.Open
		end with
		iRecCount = objRs.RecordCount
		set objRs.ActiveConnection = nothing

		set sAppID = objRs(0)
		set sAppLName = objRs(1)
		If not objRs.EOF then
			Do While Not objRs.EOF
				
%>
	<tr>
    <td class="ExcelDisplayCell">
	<Input type="checkbox" name="chkSelVal" value="<%=sAppID%>:<%=sAppLName%>"><%=sAppLName%> <br>
	</td></tr>
<%
		
				objRs.MoveNext
				sCheckVal = ""
			Loop
		end if
		objRs.Close
Else
	with objRs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT BookCode,BookName,BookshortName FROM Acc_M_DayBooks ORDER BY BookName"
			.ActiveConnection = con
			.Open
		end with
		iRecCount = objRs.RecordCount
		set objRs.ActiveConnection = nothing

		set sBookID = objRs(0)
		set sBookLName = objRs(1)
		If not objRs.EOF then
			Do While Not objRs.EOF
%>
<tr>
    <td class="ExcelDisplayCell">
<Input type="checkbox" name="chkSelVal" value="<%=sBookID%>:<%=sBookLName%>"><%=sBookLName%> <br>
</td>
</tr>
<%
				objRs.MoveNext
			Loop
		end if
		objRs.Close
End IF
%>
                                </tr>
                                    </table>
                                      </center>
                                    </div>
                                    <input type="hidden" name="hRowCount" value="<%=iRecCount%>">
								</td>
								<td align="center" class="ClearPixel" width="5" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
                                                                <input type="button" value="Done" name="B7" onclick="finaldone()" class="ActionButton" >
                                                                <input type="button" value="Cancel" name="B8" onClick="finalcancel()" class="ActionButton">
                                                                 <input type="reset" value="Reset" name="B9" class="ActionButton" >
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
  </center>
</div>
</form>
</BODY>
</HTML>
