<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	InventoryNoSeriesAmendInsert.asp
	'Module Name				:	Inventory (Master Amendment)
	'Author Name				:	TAJUDEEN S
	'Created On					:	April 20, 2004
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
	Dim sUnit,sItem,sActivity,iLength,dcrs,i
	dim sPrefix,sSuffix,sSql,sPeriod,iStartNo,sClassCode,sCatCode,sQuery
	Response.Write "<font color=red>"

	sUnit=trim(Request.Form("selUnit"))
	sItem=trim(Request.Form("selItmType"))
	sActivity=trim(Request.Form("selActType"))
	iLength=trim(Request.Form("hSeriesLen"))
	sClassCode = Trim(Request.Form("hClassCode"))
	sCatCode = Trim(Request.Form("hCatCode"))
	if Trim(sCatCode)=""  then
	    sCatCode = Trim(Request.Form("selCategory"))
	end if 
	
	if Trim(sCatCode)="0" then sCatCode = ""
	if Trim(sClassCode)="0" then sClassCode = ""

	set dcrs=server.CreateObject("ADODB.Recordset")
	
	con.BeginTrans
	
	sQuery = "Select H.SeriesNo,H.SeriesCode from INV_M_NUMBERSERIES H left join INV_M_NoSeriesClass D on H.SeriesCode = D.SeriesCode where ORGANISATIONCODE = '" & sUnit & "' and ActivityType ='"& sActivity &"' "
	if Trim(sCatCode)<>"" then
	    sQuery= sQuery&" and CatCode in ("& sCatCode &")"
	end if 
	if Trim(sClassCode)<>"" then
	    sQuery= sQuery&" and ClassCode in ("& sClassCode &")"
	end if 
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		for i = 1 to iLength
			iStartNo=trim(Request.Form("txtStartNo" & i))
			sPrefix=trim(Request.Form("txtPrefix" & i))
			sSuffix=trim(Request.Form("txtSuffix" & i))
			sPeriod=trim(Request.Form("txtPeriod" & i)) 
			
			ssql = "UPDATE APP_R_NOSERIESMODULEENTRY SET NUMBER="&iStartNo&",PREFIX=" & pack(sPrefix) & ", SUFFIX=" & pack(sSuffix) & " WHERE OUDEFINITIONID = " & pack(sUnit) & " AND SERIESNO = " & dcrs(0) & " AND SERIESCODE = " & dcrs(1) & " AND ENTRYNO=" & i & " AND PERIOD=" & pack(sPeriod)
			Response.Write ssql
			con.Execute sSql
		next
	end if
	dcrs.Close 

	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End
		Response.Clear 
		con.CommitTrans
	end if

	con.close
	set con = nothing
	Response.Redirect "InventoryNoSeriesAmendEntry.asp"
%>

