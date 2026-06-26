<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetACCGroup.asp
	'Module Name				:	Accounts (Master Creation)
	'Author Name				:	Senthil E
	'Created On					:	May 09,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view)
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
	dim sGCode,sGName,sPGroup,sOrgId,bHeadFlag,sHeadName,sHeadCode,sCatCode

	Set objRs = Server.CreateObject("ADODB.RecordSet")
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set oNodRoot = oDOM.createElement("Root")
	oDOM.appendChild oNodRoot

	sOrgId=Request.QueryString("orgid")
	bHeadFlag=Request.QueryString("flag")


	sQuery="SELECT CategoryCode, CategoryName from Acc_M_AccountCategory order by CategoryCode"
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

	if not objRs.EOF then
		while not objRs.EOF
			Set oNodGroup = oDOM.createElement("ACCGroup")
				oNodGroup.setAttribute "GroupCode", trim(sGCode)
				oNodGroup.setAttribute "GroupName", trim(sGName)
				oNodGroup.setAttribute "ParentCode",trim(sGCode)
			oNodRoot.appendChild oNodGroup
			objrs.MoveNext
		wend
	end if
	objRs.Close



	sQuery="SELECT AccountsGroupCode,AccountsGroupName,AccountsParentGroup,GroupCategory FROM Acc_M_AccountGroups ORDER BY Grouphierarchy"

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
	set sCatCode=objRs(3)
	if not objRs.EOF then
		while not objRs.EOF
			Set oNodGroup = oDOM.createElement("ACCGroup")
				oNodGroup.setAttribute "GroupCode", trim(sCatCode)& trim(sGCode)
				oNodGroup.setAttribute "GroupName", trim(sGName)
				if trim(sGCode)=trim(sPGroup) then
					oNodGroup.setAttribute "ParentCode", trim(sCatCode)
				else
					oNodGroup.setAttribute "ParentCode",trim(sCatCode)& trim(sPGroup)
				end if
			oNodRoot.appendChild oNodGroup

			if trim(bHeadFlag)="B" then

					sQuery=	"select Distinct AccountHead,AccountDescription from VwOrgGLHeads where AccountsGroupCode='"&sGCode&"' and OUDefinitionID='"&sOrgId&"' ORDER BY 2"

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
						Set oNodHead = oDOM.createElement("ACCHead")
							oNodHead.setAttribute "HeadCode", trim(sGCode)&":"&trim(sHeadCode)
							oNodHead.setAttribute "HeadName", trim(sHeadName)
							oNodHead.setAttribute "ParentCode",trim(sCatCode)& trim(sGCode)
						oNodGroup.appendChild oNodHead
						objRs1.MoveNext
					wend
					objRs1.Close
			end if
			objRs.MoveNext
		wend
		objRs.close


	else
			Set oNodGroup = oDOM.createElement("ACCGroup")
				oNodGroup.setAttribute "BookCode", "0"
				oNodGroup.setAttribute "GroupName", "No Records Avialable"
				oNodGroup.setAttribute "ParentCode","0"
			oNodRoot.appendChild oNodGroup

	end if


	'oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-EnqAmdAgentData.xml")

	Response.ContentType="text/xml"
	Response.Write oDOM.xml

	con.close
	set objRs=nothing
	set objRs1=nothing
	set con = nothing
	set oDOM=nothing
%>

