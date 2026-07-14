<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetNumberSeriesList_Mtn.asp
	'Module Name				:	Maintenance (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	01 Jan 2009
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
<!--#include virtual="/include/purpopulate.asp"-->
<%
	Dim sUnit,sActitity,sTemp,sTempVal,sQuery,Objrs,iCtr,sSerCode
	Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp,sCrValue,sDrValue
	Dim oDom,newElem,Root,sSeriesCode,sSeriesNo,sRetTy,objRs1
	Dim iSeriesCode,iSeriesNo,iAgentCode,iAddSerNo,iAddSerCode,sAgentName
	Dim iEntNo,sWorkGroupType,sInvTy,sSalTy,sAgTy,iSerNo,iSerCode,sWorkGroupValues,sSalVal,sInvVal
	Dim sWorkGroupDesc,sInvDesc,sSalDesc,sNoUsed,Objrs3
	
	
	sTempVal = Request.QueryString("sVal")
	'sTempVal = "010101:2" ' for testing
	
	sTemp = Split(sTempVal,":")
	
	iCtr = 1
	
	sUnit = sTemp(0)
	sActitity = sTemp(1)
	
	Set oDom = server.CreateObject("Microsoft.xmlDom")
	Set Root = oDom.createElement("Root")
	oDom.appendChild Root
	
	sMon = Month(Date)
	sYear = Year(Date)

	IF CInt(sMon) <=9 Then
		sMon = 0&sMon
	End IF
	sMonYr = sMon&sYear
	sFinYear = GetFinancialYear(sMonYr)
	saTemp = Split(sFinYear,":")
	sYear = Right(saTemp(0),4)
	sMon = Mid(saTemp(0),4,2)
	sFinFrom = sYear&sMon

	sYear = Right(saTemp(1),4)
	sMon = Mid(saTemp(0),4,2)
	sFinTo = sYear&sMon
	Set Objrs = Server.CreateObject("ADODB.RecordSet")
	Set Objrs1 = Server.CreateObject("ADODB.RecordSet")
	Set Objrs3 = Server.CreateObject("ADODB.RecordSet")
	
	sQuery = "Select NoSeriesTransactionNo,OrganisationCode,isNull(SeriesNo,0),isNull(SeriesCode,0) "&_
			 "From MTN_M_NoSeries Where NoSeriesStatus <> '1'  and OrganisationCode = '"&sUnit&"' and ActivityType = '"&sActitity&"' "
															 
	With Objrs
		.CursorType = 3
		.CursorLocation = 3
		.ActiveConnection = Con
		.Source = sQuery
		.Open
	End With
	'Response.Write sQuery
	Set Objrs.ActiveConnection = Nothing
	
	Do While Not Objrs.EOF
		iSeriesNo = Objrs(2)
		iSeriesCode = Objrs(3)
		
		sQuery = "Select isNull(WorkGroup,0),isNull(MainSeriesNo,0),isNull(MainSeriesCode,0),isnull(WorkGroupValue,0) "&_
					 " From VwMtnNoSeriesSel Where NoSeriesTransactionNo = "&Objrs(0)&" "
			With Objrs1
				.CursorType = 3
				.CursorLocation = 3
				.ActiveConnection = Con
				.Source = sQuery
				.Open
			End With
			'Response.Write sQuery
			Set Objrs1.ActiveConnection = Nothing
			IF Not Objrs1.EOF Then
				iEntNo = "1"
				sWorkGroupType = objRs1(0)
				sWorkGroupValues = objRs1(3)
			End IF
			objRs1.Close
			
			sNoUsed = "N"
			sNoUsed = CheckNoSerUsedMaintenance(sUnit,sActitity,sWorkGroupType,sWorkGroupValues)
			'Response.Write sNoUsed
			
			IF CStr(sWorkGroupType) = "0" Then
				sWorkGroupType = "All"
				sWorkGroupDesc = ""
			Else
				sWorkGroupType = "Specific"
				sWorkGroupDesc = GetWorkGroup(sWorkGroupValues)
			End IF
		
						
			 
			Set newElem = oDom.createElement("NumSeriesList")
			newElem.setAttribute "TransNo",Objrs(0)
			newElem.setAttribute "NumFor",""
			newElem.setAttribute "SeriesNo",iSeriesNo
			newElem.setAttribute "SeriesCode",iSeriesCode
			newElem.setAttribute "EntryNo",iEntNo
			newElem.setAttribute "WorkGroupType",sWorkGroupType
			newElem.setAttribute "WorkGroupValue",sWorkGroupValues
			newElem.setAttribute "WorkGroupDesc",sWorkGroupDesc
			newElem.setAttribute "EditCheck","N"
			newElem.setAttribute "NoUsed",sNoUsed
				
		Root.appendChild newElem
		iCtr = iCtr + 1
		Objrs.MoveNext
	Loop
	Objrs.Close
		
	
	Response.ContentType="text/xml"
	Response.Write oDom.xml											
	
	
	
%>
<%
Function CheckNoSerUsedMaintenance(Unit,Activity,sWorkGroupType,sWorkGroupValue)
dim sSql,rsTemp,sDesc,sOrgID
sOrgID = getUnitNoOUDefID(Unit,"U")
'Response.Write "sOrgID"&sOrgID


Set rsTemp = server.CreateObject("ADODB.RecordSet")
'-----------------------------------------------------------------------------
if Activity = "1" then ''Activity Plan
	sSql = "select ActivityPlanNo from MTN_M_ActivityPlanHeader where ActivityPlanCode is not null "
elseif Activity = "2" then  '' Work Order
	sSql = "Select WorkOrderNo from MTN_M_WorkOrderHeader where WorkOrderCode is not null "
elseif Activity = "3" then ''Breakdown Instruction
	sSql = "Select BrkDwnInstructionNo from MTN_T_BreakDnWrkInstr where BrkDwnInstructionCode is not null "
elseif Activity = "4" then ''Daily Maintenance
	sSql = "Select MaintenanceNo from MTN_M_DailyMaintenanceHeader where MaintenanceCode is not null "
end if
'--------------------------------------------------------------------------
'Response.Write sSql
with rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sSql
	.Open
end with
set rsTemp.ActiveConnection = nothing

if not rsTemp.EOF then
	CheckNoSerUsedMaintenance = "Y"
else
	CheckNoSerUsedMaintenance = "N"
end if
rsTemp.Close

	
End Function
%>


<%
	Function GetWorkGroup(sVal)
	
		Dim sRetVal,rsTemp
		
		Set rsTemp = Server.CreateObject("ADODB.RecordSet")
		
		sQuery = "Select WorkGroupName from PRD_M_WORKGROUP  Where WorkGroupCode = '"&sVal&"' "
		'Response.Write "<p>" & sQuery
		with rsTemp
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sQuery
			.Open
		end with
		set rsTemp.ActiveConnection = nothing
		if not rsTemp.EOF then
			sRetVal = rsTemp(0)
		end if
		rsTemp.Close
		
		GetWorkGroup = sRetVal
	End Function
	
%>
