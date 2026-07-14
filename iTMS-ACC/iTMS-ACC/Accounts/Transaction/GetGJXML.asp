<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>


<%
	'Program Name				:	GetGJXML.asp
	'Module Name				:	ACCOUNTS (Transaction)
	'Author Name				:	Ragavendran R 
	'Created On					:	Feb 11,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<%

	Dim sTemp,sBkCode,sTransNo,sPDocType,sPBookCode,sPAdjType
	Dim Root,objhttp,rs1,rs2,rs3,rs4,sQry,iENo
	Dim sRecNo,	sRInvNo,sRInvDate,sRTransAmt,sRAmdToAdj,sRDocType,sReceivableNo,sRAdjType
	Set rs1 = server.CreateObject("ADODB.RecordSet")
	Set rs2 = server.CreateObject("ADODB.RecordSet")
	Set rs3 = server.CreateObject("ADODB.RecordSet")
	sTransNo = Request.QueryString("hTransNo")
	
	sQry = "select Distinct VoucherEntryNumber,TransCrDrIndication,Amount,AccountingUnit,TdsOnAmount,TdsPercentage from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&sTransNo &" and isNull(TDSFlag,'N') = 'N' "

			with rs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con
				.Open
			End With

			if not rs1.EOF then
				iENo	= rs1(0)
			end if
	
	
	sQry ="Select C.ReceivableNumber,C. Narration,Convert(Varchar,C.VoucherDate,103),C.AmountReceivable,C.AmountReceived,C.CreatedTransNo "&_
				      " From Acc_T_CreatedReceivables C Where C.CreatedTransNo = "&sTransNo
				      
	'			Response.Write "Test="&sQry
				
				with rs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con
					.Open
				End with

				if not rs3.EOF then
					sRecNo = rs3(0)
					sRInvNo = rs3(1)
					sRInvDate = rs3(2)
					sRTransAmt = rs3(3)
					sRAmdToAdj = rs3(4)
					sRCTransNo = rs3(5)

					sQry = "select CrDrIndication from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& sTransNo  &" "
					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQry
						.ActiveConnection = con
						.Open
					End with
					If not rs2.EOF then
						sRDocType = rs2(0)
					End If
					rs2.Close

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select ReceivableNumber from Acc_T_Receivables where CreatedReceivable = "& sRecNo &" "
						.ActiveConnection = con
						.Open
					End with

					If not rs2.EOF then
						sReceivableNo = rs2(0)
					End If
					rs2.Close

					If trim(sReceivableNo) = "" then
						With rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = "select ReceivableNumber from Acc_T_Receivables where DRCreatedReceivable = "& sRecNo &" "
							.ActiveConnection = con
							.Open
						End with
						If not rs2.EOF then
							sReceivableNo = rs2(0)
						End If
						rs2.Close
					End If

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = "select BookCode from Acc_T_CreatedVoucherHeader where CreatedTransNo ="& sRCTransNo &" "
						.ActiveConnection = con
						.Open
					End with
					If not rs2.EOF then
						sRBookCode = rs2(0)
					End If
					rs2.close

					If sRBookCode = "04" then sRAdjType = "PI"
					If sRBookCode = "05" then sRAdjType = "I"
					If sRBookCode = "06" then sRAdjType = "D"
					If sRBookCode = "07" then sRAdjType = "C"
					
					sTemp = sRecNo & "#" & sRInvNo & "#" & sRInvDate & "#" & sRTransAmt & "#" & "0" & "#" & sRAmdToAdj & "#" & sRDocType & "#" & "0" & "#" & sReceivableNo & "#" & sRAdjType

					rs3.movenext
				end if
				rs3.Close
	
	Response.ContentType = "text/xml"
	Response.Write sTemp
%>