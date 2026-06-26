<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetCCGroup.asp
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

<!-- #include File="../../include/DatabaseConnection.asp" -->
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
	
	sQuery="SELECT CCGroupCode,CCGroupName,CCParentGroup FROM Acc_M_CostCenterGroup ORDER BY CCGroupCode"
	
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

	if not objRs.EOF then
		while not objRs.EOF
			Set oNodGroup = oDOM.createElement("COSTGroup")
				oNodGroup.setAttribute "GroupCode", trim(sGCode)
				oNodGroup.setAttribute "GroupName", trim(sGName)
				oNodGroup.setAttribute "ParentCode",trim(sPGroup)
			oNodRoot.appendChild oNodGroup
			
			if trim(bHeadFlag)="B" then
				sQuery="select CostCenterHead,CCAccountDescription from VwOrgCostCenter where CCGroupCode='"&sGCode&"'"&_
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
						Set oNodHead = oDOM.createElement("COSTHead")
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
		objRs.close
	
	else
		Set oNodGroup = oDOM.createElement("COSTGroup")
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

