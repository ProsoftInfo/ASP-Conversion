

<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	DataImportModuleWiseInsert.asp	
	'Module Name				:	Inventory (Data Import)
	'Author Name				:	Kalaiselvi R
	'Created On					:	July 07, 2011
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
<!--#include file="../../include/populate.asp"-->

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "ItemImportExport.asp"
		}
	}
//-->
</SCRIPT>
<%

Dim sModule,sListOfStoreProcedures,sReturnCode,sError,sSql

Dim Arr

Dim iCounter,nLength

Dim adoCmd,rsTemp

sModule = Request.QueryString("Module")

if sModule = "INV" then
	sListOfStoreProcedures = "Import_DB_ItemClassifications,Import_DB_Inv_ItemMaster,Import_DB_StorageLoc,Import_DB_Inv_Master,Import_DB_Stock,Import_DB_MRS,Import_DB_IssueDetails"
elseif sModule = "PRD" then
	sListOfStoreProcedures = "Import_DB_Prod_Packing,Import_DB_Prod_DailyProduction"
elseif sModule = "PUR" then
	sListOfStoreProcedures = "Import_DB_Pur_PO,Import_DB_Pur_GRN,Import_DB_Pur_ActualReceipt,Import_DB_Pur_Inspection,Import_DB_Pur_Invoice"
elseif sModule = "SAL" then
	sListOfStoreProcedures = ""
end if 

set rsTemp = Server.CreateObject("ADODB.RecordSet")

Set adoCmd = Server.CreateObject("ADODB.Command")
Set adoCmd.ActiveConnection = con
		
con.beginTrans

Response.Write "<p style='color:red'> "

Arr = Split(sListOfStoreProcedures,",")	
	
nLength = ubound(Arr)
	
for iCounter = 0 to nLength
	
	' calling Stored Procedure 

	adoCmd.CommandText = Arr(iCounter)
	adoCmd.CommandType = 4 'adCmdStoredProc

	Response.Write "<p> " & Arr(iCounter)
	
	sReturnCode = 0
	
	'adoCmd.Parameters.Append adoCmd.CreateParameter("@DailyPackCodeToReturn",129,2,10,sDailyPackingCode)
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@ReturnCode",3,2,5,sReturnCode)
	
	adoCmd.Execute()
	
	sReturnCode = trim(adoCmd.Parameters("@ReturnCode").Value)
		
	Response.Write "<p>sReturnCode =  " & sReturnCode
	
	if trim(sReturnCode) = trim("2") then
		sError = "Error in " & Arr(iCounter)
		exit for
	end if 
							
	
next



if trim(sReturnCode) = trim("2") then

	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
else
	'con.RollbackTrans
	'Response.end
	
	sError = "Data Import Sucessfully Completed"	
	con.CommitTrans
end if

con.close
set con = nothing

%>

<BODY onLoad = "msgbox('<%=sError%>','Y')">
