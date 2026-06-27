
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	InstrumentDetView.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 01, 2011
	'Modified By				:	
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
Dim objRs,sQuery,iTransNo,nInstNo,sInstDate,sPayableAt,sDrOnBank,sInstrType,nInstrAmt,sOrgName
set objRs = Server.CreateObject("ADODB.Recordset")

sOrgName = Session("Orgshortname")
iTransNo=Request("TransNo")
'Response.Write iTransNo

sQuery = " Select isNull(BankInstrumentNo,0),Convert(DateTime,BankInstrumentDate,103),isNull(PayableAt,''),"&_
		 " isNull(DrawnOnBank,''),BankInstrumentType,InstrumentAmount From Acc_T_CreatedVoucherInstrumentDet "&_
		 " Where CREATEDTRANSNO ="& iTransNo & " "

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

If not 	objRs.EOF then
	nInstNo		= objrs(0)
	sInstDate	= objrs(1)
	sPayableAt	= objrs(2)
	sDrOnBank	= objrs(3)
	sInstrType	= objrs(4)
	nInstrAmt	= objrs(5)
End IF
objRs.Close

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Instrument Details</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Instrument Details
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">
											<tr>
												<td class="FieldCellSub">Unit Name </td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sOrgName%></span></td>
											</tr>
											<tr>
												<td class="FieldCellSub">Instrument No</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=nInstNo%></span>
												</td>
											</tr>
											<tr>
												<td class="FieldCellSub">Instrument Date</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sInstDate%></span></td>
											</tr>
											<tr>
												<td class="FieldCellSub">Payable At</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sPayableAt%></span></td>
											</tr>
											<tr>
												<td class="FieldCellSub">Drawn On</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sDrOnBank%></span></td>
											</tr>
											<tr>
												<td class="FieldCellSub">Instrument Type</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=sInstrType%></span></td>
											</tr>
											<tr>
												<td class="FieldCellSub">Instrument Amount</td>
												<td class="FieldCellSub"><span class="DataOnly"><%=FormatNumber(nInstrAmt,2,,,0)%></span></td>
											</tr>
											
										</table>
									</td>
									<td align="center">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Close" name="B3" class="ActionButton" onclick="window.close()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="BottomPack">
									</td>
								</tr>

							</table>
						</td>
					</tr>

				</table>
			</td>
		</tr>

	</table>
	</form>
</body>
