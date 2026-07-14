<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParTypeCreatUpdatePopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 21,2011
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
<%
Dim sQuery,objRs
dim sName,sShortName,sType,iCode,iRecCount,sAction,sParType,sParSubType

'XML DOM Variables
Dim oDOM,newElem,Root,objfs,HeaderNode,nodetemp


' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set objRs = Server.CreateObject("ADODB.RecordSet")

sType=trim(Request.Form("radParType"))
sName=trim(Request.Form("txtSubTypeName"))
sShortName=trim(Request.Form("txtSubTypeShortName"))
sAction = Request.QueryString("Action")
'Response.Write "sAction="& sAction
if trim(sAction)<>"C" then
sParType = Request.QueryString("ParType")
sParSubType = Request.QueryString("ParSubType")
end if


if trim(sAction)="C" then

    sQuery="select count(1)  from APP_M_PartyTypes where SubTypeName='"&sName&"'"
    with objRs	
	    .CursorLocation = 3
	    .CursorType = 3
	    .Source = sQuery
	    .ActiveConnection = con
	    .Open
    end with
    iRecCount=objRs(0)
    objRs.Close
    %>

    <HTML>
    <head>
        <base target="_self"></base>
        <LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
        <script src="/Scripts/itms-modern-compat.js"></script>

        <SCRIPT>
        <!--
	        function msgbox(strr,flag) {
		        if (flag == "Y") {
	    		        alert(strr);
	    		        window.close();
				        }
    	        }
        //-->
        </SCRIPT>
        <script>
        window.__itmsPopupCompat = { type: "autoClose", returnValue: "Done" };
        </script>
        <script src="../../scripts/PopupModernCompat.js"></script>
        </head>
    <%
    if iRecCount =0 then

	    sQuery="select isnull(max(PartySubType)+1,1)  from APP_M_PartyTypes where PartyType='"&sType&"'"
	    with objRs	
		    .CursorLocation = 3
		    .CursorType = 3
		    .Source = sQuery
		    .ActiveConnection = con
		    .Open
	    end with
	    iCode=objRs(0)

	    sQuery="INSERT APP_M_PartyTypes(PartyType, PartySubType, SubTypeName, SubTypeShortName) "&_
			    "VALUES('"&sType&"',"&iCode&",'"&sName&"','"&sShortName&"')"

	    con.Execute(sQuery)
    '------------------------Update XML File -----------------------
	    if objfs.FileExists(Server.MapPath("../xmldata/PartyType.xml")) then
		    oDOM.Load server.MapPath("../xmldata/PartyType.xml")
		    Set Root = oDOM.documentElement
	    else	
		    Set Root = oDOM.createElement("Root")
		    oDOM.appendChild Root
	    end if

	    Set newElem = oDOM.createElement("PartyType")

	    newElem.setAttribute "Type", sType
	    newElem.setAttribute "Code", iCode
	    newElem.setAttribute "NAME", sName
	    newElem.setAttribute "SHORTNAME", sShortName
	    Root.appendChild newElem
    	
	    oDOM.Save server.MapPath("../xmldata/PartyType.xml")
		
%>
<BODY onLoad = "msgbox('Party Type Created Successfully','Y')">
<%else%>
<BODY onLoad = "msgbox('Party Type Already Exist','N')">
<%
end if	
set objRs=nothing

%>
<form name="formname">
</form>
</BODY>
</HTML>
<%
elseif trim(sAction)="U" then
    sQuery="Update APP_M_PartyTypes set SubTypeName='"&sName&"',SubTypeShortName='"&sShortName&"' where "&_
		    "PartyType='"&sParType&"' and PartySubType="&sParSubType
    		
	    con.Execute(sQuery)
    '------------------------Update XML File -----------------------
	    oDOM.Load server.MapPath("../xmldata/PartyType.xml")
	    Set Root = oDOM.documentElement

    For Each HeaderNode In Root.childNodes
	    if HeaderNode.Attributes.Item(0).nodeValue = sParType and HeaderNode.Attributes.Item(1).nodeValue = sParSubType then
		    HeaderNode.Attributes.Item(2).nodeValue=sName
		    HeaderNode.Attributes.Item(3).nodeValue=sShortName
	    end if
    next
    	
    oDOM.Save server.MapPath("../xmldata/PartyType.xml")
    		
    %>
    <HTML>
    <head>
        <base target="_self"></base>
    <LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
    <script src="/Scripts/itms-modern-compat.js"></script>
    <SCRIPT>
    <!--
	    function msgbox(strr,flag) {
		    if (flag == "Y") {
			    alert(strr);
			    window.close();
		    }
	    }
    //-->
    </SCRIPT>
    <script>
    window.__itmsPopupCompat = { type: "autoClose", returnValue: "Done" };
    </script>
    <script src="../../scripts/PopupModernCompat.js"></script>
    </head>
    <BODY onLoad = "msgbox('Party Type Update Successfully','Y')">
        <form name="formname">
        </form>
    </BODY>
    </HTML> 
<%
elseif trim(sAction)="D" then
    sQuery="select count(1) from APP_R_OrgParty where PartyType='"&sParType&"' and PartySubType="&sParSubType
    'Response.Write sQuery
    With objRs
	    .CursorLocation = 3
	    .CursorType = 3
	    .Source = sQuery
	    .ActiveConnection = con
	    .Open
    End with
    Set objRs.Activeconnection = nothing
    if CDbl(objRs(0))>0 then
    %>
    <HTML>
    <head>
    <base target="_self"></base>
    
    <LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
    <script src="/Scripts/itms-modern-compat.js"></script>
    <SCRIPT>
    <!--
	    function msgbox(strr,flag) {
		    if (flag == "Y") {
			    alert(strr);
			    window.close();
		    }
	    }
    //-->
    </SCRIPT>
    <script>
    window.__itmsPopupCompat = { type: "autoClose", returnValue: "Done" };
    </script>
    <script src="../../scripts/PopupModernCompat.js"></script>
</head>
    <BODY onLoad = "msgbox('Party Type Cannot be Deleted already Related with party','N')">
    <form name="formname">
        </form>
    </BODY>
    </HTML>
    <%
    else

	    sQuery="delete Acc_R_OrgPartyType where PartyType='"&sParType&"' and PartySubType="&sParSubType
	    con.Execute(sQuery)

	    sQuery="delete APP_M_PartyTypes where PartyType='"&sParType&"' and PartySubType="&sParSubType
    			
	    con.Execute(sQuery)
	    '------------------------Update XML File -----------------------
		    oDOM.Load server.MapPath("../xmldata/PartyType.xml")
		    Set Root = oDOM.documentElement

	    For Each HeaderNode In Root.childNodes
		    if HeaderNode.Attributes.Item(0).nodeValue = sParType and HeaderNode.Attributes.Item(1).nodeValue = sParSubType then
			    set nodetemp=Root.removeChild(HeaderNode)
		    end if
	    next
    		
	    oDOM.Save server.MapPath("../xmldata/PartyType.xml")
    		
    %>
    <HTML>
    <head>
        <base target="_self"></base>
    <LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
    <script src="/Scripts/itms-modern-compat.js"></script>
    <SCRIPT>
    <!--
	    function msgbox(strr,flag) {
		    if (flag == "Y") {
			    alert(strr);
			      window.close();
		    }
	    }
    //-->
    </SCRIPT>
    <script>
    window.__itmsPopupCompat = { type: "autoClose", returnValue: "Done" };
    </script>
    <script src="../../scripts/PopupModernCompat.js"></script>
    </head>

    <BODY onLoad = "msgbox('Party Type Deleted Successfully','Y')">
        <form name="formname">
        </form>
    </BODY>
    </HTML>
    <%
    end if
    objRs.Close
end if
%>

<%'ParTypeCreatEntryPopup.asp%>
