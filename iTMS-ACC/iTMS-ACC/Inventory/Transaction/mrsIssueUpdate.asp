
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsIssueUpdate.asp
	'Module Name				:	Inventory (Issue Edit)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 18,2014
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssueItemEntry.asp
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
<!--#include file="../../include/mrsIssueUpdateCommon.asp"-->

<%
Dim sAppCallFrom,rsIssObj,sIssQuery,sPONO,sIssEntNo
sAppCallFrom = Request("hCallFrom")
set rsIssObj= server.CreateObject("ADODB.Recordset")
sIssEntNo = Request("IssEntNo")
    con.begintrans

    '' To Call the Issue Insert Common Function
    MrsIssueUpdate(sIssEntNo)
    
    
   
    if con.Errors.count <> 0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
		'Response.Write "<p>sSalInvConfirm="&sSalInvConfirm
	'	Response.Clear 
		
		if sSInvType ="X" then sSInvType ="CB"
        if sSInvType ="Y" then sSInvType ="NEB"
        if sSInvType ="Z" then sSInvType ="EB"
        Response.Write "<p>sSinvType = "& sSinvType
        Response.Write "<p> sSalType = "& sSSalType
        Response.Write "<p> POS = "& sSALPOSID
        Response.Write "<p> sSinvType = "& sSinvTypeName
        Response.Write "<p>  sSalType = "& sSSalTypeName
        Response.Write "<p>  POS = "& sSSALPOSIDName
        
        sIssQuery = "Select AppRefNo from INV_T_MRSHeader where MRSNumber = "& sAppRefNo 
	    rsIssObj.Open sIssQuery,con
	    if not rsIssObj.EOF then
	        sPONO = rsIssObj(0)
	    end if 
	    rsIssObj.Close 
	    
	    Response.Write "<p>sAppCallFrom="& sAppCallFrom
        		
	'	con.RollbackTrans
	'	Response.End
	   Response.Clear
	   con.CommitTrans
	    Response.Redirect(sRedirectTo)
	end if 'if con.Errors.count <> 0 then
%>
