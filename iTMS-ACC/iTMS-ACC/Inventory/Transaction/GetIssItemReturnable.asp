<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetIssItemReturnable.asp
	'Module Name				:	Include
	'Author Name				:	Ragavendran
	
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/populate.asp" -->
<%
	Dim rsTemp
	Dim sIssueEntryNo,sQuery,sReturnValue

	Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	
	sIssueEntryNo = Request.QueryString("RefCodes")
	
    sQuery = "Select IsNull(Returnable,'N'),IsNull(ReturnItem,'S') from INV_T_MaterialIssueHeader where IssueEntryNo in ("& sIssueEntryNo &")"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        sReturnValue = trim(rsTemp(0))&":"&trim(rsTemp(1))
    end if
    rsTemp.close
    Response.write  sReturnValue
%>