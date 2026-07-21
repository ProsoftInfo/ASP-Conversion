<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassUpdate.asp
	'Module Name				:	Sales - Gate Pass
	'Author Name				:	Ragavendran R
	'Created On					:	APRIL 03,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	GatePassEntry.asp
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
<!--#include File="../../include/populate.asp" -->
<!--#include File="../../include/NoSeries.asp" -->
<HTML><HEAD><TITLE>iTMS - Gate Pass</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	dim dcrs,iGatePassNo,sSql,objfs,ObjCmd,sForSubConNo,sInvoiceType
	dim sOrgID,sDCNo,sRemarks,sOrgIDTr,sFileName,sRefType
	dim iEntryNo, iPurRetNo, iActionTakenNo, sCreatedBy
	Dim iClassCode,iItemCode, iInspNo, iRcptNo,iLocNo,iBinNo,sRcptDt,iQty,UoM, sQuery
	Dim sGatePassDate, sItemType,sInvNo,Arr1
	Dim iSeriesNo, iSeriesCode,sExistDCNo
		
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set dcrs = CreateObject("ADODB.RecordSet")
	
	iGatePassNo = trim(Request.Form("hGatePassNo"))
	sForSubConNo = trim(Request.Form("hForSubConNo"))
	sDCNo = trim(Request.Form("hDCNo"))
	sOrgID = trim(Request.Form("hOrg"))
	sRemarks = trim(Request.Form("txtRemarks"))
	sGatePassDate = trim(Request.Form("hGatePassDate")) 
	sItemType = trim(Request.Form("hItemType")) 
Response.Write "<font color=Red>"	
'	Response.Write "sDCNo = "& sDCNo
	
	If Trim(sGatePassDate) = "" Then 
		sGatePassDate = Formatdate(date())
	Else
		sGatePassDate = sGatePassDate
	End If

	sInvNo = ""
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT IsNull(ReferenceNo,0),isNull(RefType,''),isNull(InvoiceType,'A'),ISNULL(DCCODE,'-') FROM FORGATEPASSHEADER WHERE GATEPASSNO = " & iGatePassNo & " AND OrganisationCode = " & Pack(sOrgID) 
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sInvNo = dcrs(0)
		sRefType = dcrs(1)
		sInvoiceType = dcrs(2)
		sExistDCNo = dcrs(3)
	end if
	dcrs.Close
			
	iEntryNo = 0
		
	sFileName = "../../Inventory/welcome_Inventory.asp"
	
	if sRemarks = "" then
		sRemarks = "NULL"
	else
		sRemarks = Pack(sRemarks)
	end if

	con.BeginTrans

    if sExistDCNo <>"-" then
	    sSql = "UPDATE FORGATEPASSHEADER SET STATUS = 'Y', GENERATEDON = Convert(datetime,'" & sGatePassDate & "',103) , REMARKS = " &_
		    sRemarks & " WHERE GATEPASSNO = " & iGatePassNo & " AND OrganisationCode = " & Pack(sOrgID) 
	    'Response.Write "sSql = "& sSql
	else
	
	
	    with dcrs
		    .CursorLocation = 3
		    .CursorType = 3
		    '.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'DC' AND ITEMTYPE = " & Pack(sItemType) & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " Order By SeriesCode desc"
		    .Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'DC' AND ORGANISATIONCODE = " & Pack(sOrgID) & " Order By SeriesCode desc"
		    .ActiveConnection = con
		    .Open
	        end with
	    set dcrs.ActiveConnection = nothing

	    if not dcrs.EOF then
		    iSeriesNo = trim(dcrs(0))
		    iSeriesCode = trim(dcrs(1))
		    sDCNo = GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,FormatDate(date))
		    sDCNo = Pack(sDCNo)
	    else
		    sDCNo = "NULL"
	    end if
	    dcrs.close
	
	    sSql = "UPDATE FORGATEPASSHEADER SET STATUS = 'Y', GENERATEDON = Convert(datetime,'" & sGatePassDate & "',103) , REMARKS = " &_
		    sRemarks & ",DCCODE = " & sDCNo & " WHERE GATEPASSNO = " & iGatePassNo & " AND OrganisationCode = " & Pack(sOrgID) 
	    'Response.Write "sSql = "& sSql
	end if
	
	con.Execute sSql
	
	'con.rollbacktrans
	'Response.End 
	Con.CommitTrans
	
%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="post" name="formname" action="<%=sFileName%>">
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	alert("DC / Gatepass has been Modified.")
	document.formname.submit 
</script>
</form>
</Body>
</HTML>