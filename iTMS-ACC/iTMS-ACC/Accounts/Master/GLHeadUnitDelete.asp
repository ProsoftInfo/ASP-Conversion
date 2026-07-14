<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLHeadUnitDelete.asp
	'Module Name				:	Accounts (Master Delete)
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 02,2010
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
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim sQuery,sAccCode,sUnitID

sAccCode = Request.QueryString("AccHead")
sUnitID = Request.QueryString("UnitCode")

Con.BeginTrans

sQuery = "Delete from ACC_R_GLACCFREQUENTLYUSED where AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_R_ORGGLCOSTCENTRE WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_R_ORGGLANALYTICAL WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_R_ORGPARTYTYPE WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_R_GLACCAPPLICATIONS WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_T_GLACCOPENINGAMT WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from ACC_R_ORGGLACCOUNTHEAD WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

sQuery = "Delete from Acc_M_GLSummaryApp WHERE AccountHead = "& sAccCode &" and OUDefinitionID = '"& sUnitID &"'"
Response.Write sQuery
con.execute sQuery

Response.Clear
Con.CommitTrans

%>
