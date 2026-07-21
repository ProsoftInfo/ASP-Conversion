<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassInsert.asp
	'Module Name				:	Sales - Gate Pass
	'Author Name				:	Ragavendran R
	'Created On					:	April 03,2010
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
<!--#include File="../../include/NoSeriesCommonFunctions.asp" -->
<HTML><HEAD><TITLE>iTMS - Gate Pass</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	dim dcrs,iGatePassNo,sSql,objfs,ObjCmd,sForSubConNo,sInvoiceType
	dim sOrgID,sDCNo,sRemarks,sOrgIDTr,sReturnToPage,sRefType,sCallFrom
	dim iEntryNo, iPurRetNo, iActionTakenNo, sCreatedBy
	Dim iClassCode,iItemCode, iInspNo, iRcptNo,iLocNo,iBinNo,sRcptDt,iQty,UoM, sQuery
	Dim sGatePassDate, sItemType,sInvNo,Arr1
	Dim iSeriesNo, iSeriesCode,sArrSeries,sTempSeries
	Dim iNumClassCode,sNumClassname
		
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set dcrs = CreateObject("ADODB.RecordSet")
	
	iGatePassNo = trim(Request.Form("hGatePassNo"))
	sForSubConNo = trim(Request.Form("hForSubConNo"))
	sDCNo = trim(Request.Form("hDCNo"))
	sOrgID = trim(Request.Form("hOrg"))
	sRemarks = trim(Request.Form("txtRemarks"))
	sGatePassDate = trim(Request.Form("hGatePassDate")) 
	sItemType = trim(Request.Form("hItemType")) 
	sReturnToPage = trim(Request.Form("hReturnToPage"))
	sCallFrom  = trim(Request.Form("hCallFrom")) 
	iNumClassCode = Trim(Request.Form("hNumClassCode"))
	
	
	If Trim(sGatePassDate) = "" Then 
		sGatePassDate = Formatdate(date())
	Else
		'sGatePassDate = FormatDate(sGatePassDate)
		sGatePassDate = sGatePassDate
	End If

	sInvNo = ""
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT IsNull(ReferenceNo,0),isNull(RefType,''),isNull(InvoiceType,'A') FROM FORGATEPASSHEADER WHERE GATEPASSNO = " & iGatePassNo & " AND OrganisationCode = " & Pack(sOrgID) 
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	'Response.Write dcrs.Source  
	if not dcrs.EOF then
		sInvNo = dcrs(0)
		sRefType = dcrs(1)
		sInvoiceType = dcrs(2)
	end if
	dcrs.Close
			
	iEntryNo = 0
	
	if sRemarks = "" then
		sRemarks = "NULL"
	else
		sRemarks = Pack(sRemarks)
	end if
	if Trim(sDCNo)="" or isnull(sDCNo) or Trim(sDCNo)="-" then
	
	
	    sTempSeries = GetInvNumberSeriesCodes("DC",sOrgID,iNumClassCode)
	    sArrSeries = Split(sTempSeries,":")
	    iSeriesNo = sArrSeries(0)
	    iSeriesCode = sArrSeries(1)
	    
	    sQuery = "Select GroupName from INV_M_Classification where GroupCode = "& iNumClassCode
	    dcrs1.Open sQuery,con
	    if not dcrs1.EOF then
	        sNumClassName = Trim(dcrs1(0))
	    end if
	    dcrs1.Close 
	    
    	
	    if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
	        sDCNo = "NULL"
	        Response.Clear 
	        Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Gate Pass - "& sNumClassName &"  Classification </H2></p>"
	        Response.End 
	    else
	        sDCNo = GenSeriesNumber(sOrgID,iSeriesNo,iSeriesCode,FormatDate(date))
		    sDCNo = Pack(sDCNo)
	    end if
	end if 'if Trim(sDCNo)="" or isnull(sDCNo) or Trim(sDCNo)="-" then
	con.BeginTrans

	sSql = "UPDATE FORGATEPASSHEADER SET STATUS = 'Y', GENERATEDON = Convert(datetime,'" & sGatePassDate & "',103) , REMARKS = " &_
		sRemarks & ",DCCODE = '" & sDCNo & "' WHERE GATEPASSNO = " & iGatePassNo & " AND Organisationcode = " & Pack(sOrgID) 
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql
	
'	Con.RollbackTrans
'	Response.End 
	Con.CommitTrans
if trim(sReturnToPage) = "" then	
	sReturnToPage = "../../Inventory/welcome_Inventory.asp"	
end if 	

'if sCallFrom = "SaleInvoice" then
	sReturnToPage = sReturnToPage & "?CallFrom=GatePassEntryInsertPrg" & "&InvNo=" & sInvNo 
'end if 
If trim(sCallFrom)="SUB" then
	Response.Redirect "../../Purchase/Transaction/POLIST.ASP?ACTN=L"
end if
If trim(sCallFrom)="SINV" then
	Response.Redirect "../../Sales/Transaction/SALESINVOICES.ASP?ACTN=L"
end if
%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="post" name="formname" action="<%=sReturnToPage%>">
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
	alert("DC / Gatepass has been generated.")
	document.formname.submit 
</script>
</form>
</Body>
</HTML>