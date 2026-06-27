<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLVouAppUpdate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 27, 2003
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<%
dim iTransNo,iUserId,sMode,sVouStatus,sQuery,oDOM,oNodRoot,sBookCode
dim oVouRoot

iTransNo=Request("TransNo")
iUserId=Request("User")
sMode=Request("Mode")
sBookCode=Request("BkCode")


IF sMode="E" THEN
	if CInt(iUserId)=0 then
		sVouStatus="010103" 'FOR ACCOUNTING
	else
		sVouStatus="010101" 'FOR APPROVAL
	end if
	IF sBookCode="BA" THEN
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		oDOM.async = false
		oDOM.load(Request)
		set oNodRoot=oDOM.documentElement
		dim sInsType,sInsNo,sInsDate,sPayableat,sDrawnOn

		sInsType=oNodRoot.Attributes.GetNamedItem("InsType").value
		sInsNo=oNodRoot.Attributes.GetNamedItem("InsNo").value
		sInsDate=oNodRoot.Attributes.GetNamedItem("InsDate").value
		sPayableat=oNodRoot.Attributes.GetNamedItem("Payableat").value
		sDrawnOn=oNodRoot.Attributes.GetNamedItem("Drawnon").value

		sQuery="update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'," &_
			"BankInstrumentType='"&sInsType&"',BankInstrumentNo='"&sInsNo&"',PayableAt='"&sPayableat&"',"&_
			"BankInstrumentDate=convert(datetime,'"&sInsDate&"',103),DrawnOnBank='"&sDrawnOn&"'"&_
			" where CreatedTransNo="&iTransNo

		con.execute(sQuery)
		IF CInt(iUserId)>0 THEN
			sQuery="insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
					"Values("&iTransNo&",1,"&iUserId&")"
			con.execute(sQuery)
		END IF

		oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

		set oVouRoot=oDOM.documentElement
		oVouRoot.appendChild(oNodRoot)

		oDOM.Save server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")

		set oDOM=nothing

	ELSE

		if CInt(iUserId)=0 then
			sVouStatus="010103" 'FOR ACCOUNTING
		else
			sVouStatus="010101" 'FOR APPROVAL
		end if

			sQuery="update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'" &_
			" where CreatedTransNo="&iTransNo
			con.execute(sQuery)

			sQuery="insert into Acc_T_VouchersForApproval(CreatedTransNo,ApprovalLevel,ToBeApprovedBy)"&_
				"Values("&iTransNo&",1,"&iUserId&")"
			con.execute(sQuery)
	END IF

END IF


%>