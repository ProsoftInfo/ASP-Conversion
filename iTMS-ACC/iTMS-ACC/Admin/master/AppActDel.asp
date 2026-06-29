<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActDel.asp
	'Module Name				:	Admin (Activity Delete)
	'Author Name				:
	'Created On					:
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

<%
dim dcrs,dcrs1,sSql,iInternalActivityID,objDOM,sType,nAppCode,sAppName,nProcessCode
dim sActName,sPrgPath,sStatus,nActCode,sUser,sItemTypeID,sUnit,sQuery
Dim nActTempNo,sActTempName,sActDesc,sDel
dim ndRoot,ndActivity,ndTemplate
con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
objDOM.async = False
objDOM.load(request)

Set ndRoot = objDOM.documentElement

if ndRoot.hasChildNodes() then
    for each ndActivity in ndRoot.childNodes
        if ndActivity.nodeName="Activity" then
            
            sType = ndActivity.getAttribute("TYPE")
            nAppCode = ndActivity.getAttribute("APPCODE")
            sAppName = ndActivity.getAttribute("APPNAME")
            nProcessCode = ndActivity.getAttribute("PROCESSCODE")
            nActCode = ndActivity.getAttribute("ACTIVITYCODE")
            sActName = ndActivity.getAttribute("ACTIVITYNAME")
            for each ndTemplate in ndActivity.childNodes
                if ndTemplate.nodeName="Template" then
                    nActTempNo = ndTemplate.getAttribute("No")
                    sDel = ndTemplate.getAttribute("Del")
                    if sDel="Y" then
                        sQuery = "Delete from Ms_ApplicationActivityTemplates where ApplicationCode = "& nAppCode &" and ProcessCode= "& nProcessCode &" and ActivityCode = "& nActCode &" and ActivityTemplateNo = "& nActTempNo
                        con.execute sQuery    
                    end if
                end if
            next
        end if
    next
end if

sQuery = "Select * from Ms_ApplicationActivityTemplates where ApplicationCode = "& nAppCode &" and ProcessCode= "& nProcessCode &" and ActivityCode = "& nActCode
with dcrs
    .cursorLocation = 3
    .cursorType = 3
    .ActiveConnection = con
    .source = sQuery
    .open
end with
if dcrs.eof then
    sQuery = "Delete from Ms_ApplicationActivity where ApplicationCode = "& nAppCode &" and ProcessCode= "& nProcessCode &" and ActivityCode = "& nActCode
    con.execute sQuery
    Response.write "ActNo:0"
else
    Response.write "ActNo:"& nActCode
end if
dcrs.close

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
