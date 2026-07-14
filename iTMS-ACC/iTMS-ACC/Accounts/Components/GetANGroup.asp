<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetANALGroup.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Created On					:	January 18,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view for Classification)
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
<%
	dim objRs,objRs1,oDOM,oNodRoot,oNodGroup,oNodHead,sQuery
	dim sGCode,sGName,sPGroup,sOrgId,bHeadFlag,sHeadName,sHeadCode
	
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	Set oNodRoot = oDOM.createElement("Root")												
	oDOM.appendChild oNodRoot
	
	sOrgId=Request.QueryString("orgid")
	bHeadFlag=Request.QueryString("flag")

	sQuery="SELECT AHGroupCode,AHGroupName,AHParentGroup FROM Acc_M_AnalyticalGroup ORDER BY AHGroupCode"
	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	
	set objRs.ActiveConnection = nothing
	
	set sGCode = objRs(0)
	set sGName = objRs(1)
	set sPGroup = objRs(2)

	Response.ContentType = "text/xml"
	if not objRs.EOF then
		while not objRs.EOF
			Set oNodGroup = oDOM.createElement("ANALGroup")
				oNodGroup.setAttribute "GroupCode", trim(sGCode)
				oNodGroup.setAttribute "GroupName", trim(sGName)
				oNodGroup.setAttribute "ParentCode",trim(sPGroup)
			oNodRoot.appendChild oNodGroup
			
			if trim(bHeadFlag)="B" then
				sQuery="select AnalyticalCode,AnalyticalName from VwOrgAnalytical where AHGroupCode='"&sGCode&"'"&_
					" and OUDefinitionID='"&sOrgId&"' "
					with objRs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQuery
						.ActiveConnection = con
						.Open
					end with
	
					set objRs1.ActiveConnection = nothing
	
					set sHeadCode = objRs1(0)
					set sHeadName = objRs1(1)
					while not objRs1.EOF
						Set oNodHead = oDOM.createElement("ANALHead")
							oNodHead.setAttribute "HeadCode", trim(sGCode)&":"&trim(sHeadCode)
							oNodHead.setAttribute "HeadName", trim(sHeadName)
							oNodHead.setAttribute "ParentCode",trim(sGCode)
						oNodGroup.appendChild oNodHead
						objRs1.MoveNext
					wend
					objRs1.Close
			end if
			objRs.MoveNext
		wend
	else
		Set oNodGroup = oDOM.createElement("ACCGroup")
			oNodGroup.setAttribute "GroupCode", "0"
			oNodGroup.setAttribute "GroupName", "No Records Avialable"
			oNodGroup.setAttribute "ParentCode","0"
		oNodRoot.appendChild oNodGroup		
	end if	
			
	Response.ContentType="text/xml"
	Response.Write oDOM.xml		
	
	con.close
	set objRs=nothing
	set objRs1=nothing
	set con = nothing
	set oDOM=nothing
%>

