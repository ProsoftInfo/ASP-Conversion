<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppRoleCreationAndAmendEntry.asp
	'Module Name				:	Admin (Role Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	December 08, 2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	RoleCreationInsert.asp
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Role Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData"><Root/></XML>
<XML ID="RetData"><Root Done=""/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
	function appRoleXml(name) {
		var element;
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function appRoleRoot(name) {
		var data = appRoleXml(name);
		var doc = data && (data.XMLDocument || data._doc || data);
		return data && data.documentElement || doc && doc.documentElement || null;
	}

	function appRoleDocument(name) {
		var data = appRoleXml(name);
		return data && (data.XMLDocument || data._doc || data) || null;
	}

	function appRoleReturnValue() {
		var root = appRoleRoot("RetData");
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window.returnValue = root;
			window.returnvalue = root;
		}
	}

	function CheckSubmit(sPara) {
		var root;
		var responseText;
		var request;
		var description = String(document.formname.txtDesc.value || "").replace(/^\s+|\s+$/g, "");
		if (description === "") {
			alert("Enter Role Description");
			document.formname.txtDesc.select();
			return false;
		}
		root = appRoleRoot("OutData");
		root.setAttribute("TYPE", String(document.formname.hPassType.value || "").replace(/^\s+|\s+$/g, ""));
		root.setAttribute("ROLEID", String(document.formname.hRoleID.value || "").replace(/^\s+|\s+$/g, ""));
		root.setAttribute("ROLEDESC", description);
		request = new XMLHttpRequest();
		request.open("POST", "AppRoleCreationInsert.asp", false);
		request.send(appRoleDocument("OutData"));
		responseText = request.responseText || "";
		if (responseText !== "") {
			alert(responseText);
			return false;
		}
		if (sPara === "CRN") {
			alert("Role has been inserted SuccessFully");
		} else {
			alert("Role has been Amend SuccessFully");
		}
		appRoleRoot("RetData").setAttribute("Done", "Y");
		appRoleReturnValue();
		window.close();
		return true;
	}

	window.addEventListener("beforeunload", appRoleReturnValue);
</SCRIPT>
<%
	Dim sTemp,sArr,sPassType,sSql,nRoleID,sRoleName

	Dim dcrs
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	sTemp = Request.QueryString("PassData")
	sArr  = Split(sTemp,":")
	sPassType = Trim(sArr(0))

	If sPassType = "EDT" Then
		nRoleID = sArr(1)

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ROLEID,ROLEDESCRIPTION FROM MS_ROLES WHERE ROLEID = "& nRoleID &" "
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		If not dcrs.EOF then
			sRoleName = Trim(dcrs(1))
		End IF
		dcrs.Close
	End IF

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="POST" name="formname" action="">

	<Input type=hidden name=hPassType value="<%=sPassType%>">
	<Input type=hidden name=hRoleID value="<%=nRoleID%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<%If sPassType = "CRN" Then%>
					<p align="center">Role Creation
				<%Else%>
					<p align="center">Role Amendment
				<%End IF%>
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td class="TabCell" valign="bottom" align="center" width="95">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">

										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>

					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<div align="left">
											<table cellpadding="0" cellspacing="0">
												<tr>
													<td class="FieldCell">Role Description</td>
													<td class="FieldCellSub">
														<%If sPassType = "CRN" Then%>
															<input type="text" name="txtDesc" maxlength=250 size="50" class="FormElem">
														<%Else%>
															<input type="text" name="txtDesc" value="<%=UCase(sRoleName)%>" maxlength=250 size="50" class="FormElem">
														<%End IF%>
													</td>
												</tr>
											</table>
										</div>
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
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<%If sPassType = "CRN" Then%>
														<input type="button" value="Create" name="B2" class="ActionButton" onClick="CheckSubmit('<%=sPassType%>')">
													<%Else%>
														<input type="button" value="Amend" name="B2" class="ActionButton" onClick="CheckSubmit('<%=sPassType%>')">
													<%End If%>
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
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
</html>
