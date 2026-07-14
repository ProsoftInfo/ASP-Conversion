<%@ Language=VBScript %>
<%	option explicit	%>
<%
		
	'Program Name				:	PayRecReminderDelete.asp
	'Module Name				:	Accounts (Transaction)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 16,2011
	'Modified On				:   
	'Modified by				:   
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
<!--#include virtual="/include/MatPopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->

<%
    Dim rsObj
    Dim sReminderNo,sQuery,sAction
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sReminderNo = Request("RemNo")
    con.begintrans
    
    sQuery = "Select ActionTaken from APP_R_ApplicationReminders where ReminderNo = "&sReminderNo
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        sAction =rsObj(0)
    end if
    rsObj.Close 
    
    if sAction = "Created" then
        sQuery = "Delete from ACC_T_OverDueReminderDet where ReminderNo = "& sReminderNo
        con.execute sQuery
        
        sQuery = "Delete from APP_R_ApplicationReminders where ReminderNo = "& sReminderNo
        con.execute sQuery
    else
        Response.Write "The Reminder Could not be Deleted"
        Response.End 
    end if
    
'	con.rollbackTrans
'	Response.End 
	 Response.Clear 
	 con.commitTrans
	
%>