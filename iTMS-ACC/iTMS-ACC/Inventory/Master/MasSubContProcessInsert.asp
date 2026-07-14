<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MasSubContProcessInsert.asp	
	'Module Name				:	Inventory 
	'Author Name				:	Ragavendran R
	'Created On					:	Oct 29,2013
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<%
Dim rsObj
Dim sSubContID,sSubContName,sSubContDesc,sOrgCode,sMode,sQuery

set rsObj = Server.CreateObject("ADODB.Recordset")

sMode = Request("Mode")
sSubContID = Request("hSubContID")
sSubContName = Request("txtSubContName")
sSubContDesc = Request("txtSubContDesc")
sOrgCode = Request("hOrgCode")

if trim(sMode)="" or IsNull(sMode) then sMode="S" 

if sMode="E" then
    sQuery = "Update APP_M_SubContractProcess set SubConProcessName = '"& sSubContName &"',SubConProcessDesc = '"& sSubContDesc & "',OrganisationCode='"& sOrgCode &"' where SubConProcessID = "& sSubContID
    con.execute sQuery
elseif sMode="D" then
    sQuery = "Delete from APP_M_SubContractProcess where SubConProcessID in ("& sSubContID &")"
    con.execute sQuery
else

    sQuery = "Select IsNull(Max(SubConProcessID),0)+1 from APP_M_SubContractProcess"
    rsObj.open sQuery,con
    if not rsObj.eof then
        sSubContID = rsObj(0)
    end if 
    rsObj.close
    
    sQuery = "Insert into APP_M_SubContractProcess(SubConProcessID,SubConProcessName,SubConProcessDesc,OrganisationCode) Values("& sSubContID &",'"& sSubContName &"','"& sSubContDesc &"','"& sOrgCode &"')"
    con.execute sQuery
end if
Response.redirect "MasSubContProcess.asp"
%>
