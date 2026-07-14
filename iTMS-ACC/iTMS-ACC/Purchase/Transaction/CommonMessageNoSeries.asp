<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	CommonMessageNoSeries.asp
	'Module Name				:	Purchase (General)
	'Author Name				:	SRIDEVI PRIYA A.
	'Created On					:	December 26, 2003
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

<!--#include virtual="/include/sessionVerify.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>

<%

Dim sTitle,sHeading,sMsg,sRedirect,sTransaction,sSeriesType

sHeading = trim(request("Heading"))
sRedirect= trim(request("Redirect"))
sTransaction = trim(request("TranName"))
sSeriesType = trim(request("SeriesType"))

sTitle = "Message"

if trim(sSeriesType) <> "INV" then	' Purchase Number series
	sHeading = " Number Series must be defined before <br> <br> starting " & sTransaction & " transaction"
	sRedirect = "../../noseries/PurNoSeriesEntry.asp"
Elseif trim(sSeriesType) = "INV" then	' Inventory Number series
	sHeading = sHeading & " <br> <br> before starting " & sTransaction & " transaction"

	'' redirection to Inventory no. series to be clarified..
	'sRedirect = "../../Inventory/Master/NoSeriesPurchaseEntry.asp"
End if

%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center"><% = sTitle %>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
    <tr>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" class="ClearPixel">
								</td>
								<td width="100%" align="center" height="300">
                                    <table cellpadding="0" cellspacing="0">
                                <tr>
                            <td class="FieldCell" align="center"> <b><%= sHeading %></b>
                            </td>
                                </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                                <tr>
                            <td align="center" class="FieldCell"> <%= sMsg %>
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                </tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">


	<input type="button" value="Home" name="B8" onClick="window.location.href='../Welcome_purchase.asp'"  class="ActionButton" tabindex="3" >
	<% if trim(sSeriesType) <> "INV" then	' Purchase Number series %>
		<input type="button" value="Define Number Series" onClick="window.location.href='<%=sRedirect%>'" name="B11" class="ActionButtonX" tabindex="3" >
	<%else%>
		<!--input type="button" value="Define Number Series"  name="B11" class="ActionButtonX" tabindex="3" -->
	<%end if%>
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                                </tr>
						</table>
					</td>
    </tr>
</table>
</form>
</BODY>
</HTML>