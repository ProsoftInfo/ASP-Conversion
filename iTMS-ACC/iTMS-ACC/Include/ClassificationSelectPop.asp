<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ClassificationSelectPop.asp
	'Module Name				:	Inventory (Classification Selection)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 10, 2004
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
<!--#include virtual="/Include/GetSettings.asp"-->
<%
	dim sIP
	sIP = GetSettings("IP")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification Selection</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" ID="Data"><root/></script>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../scripts/ModalReturnCompat.js"></script>
<SCRIPT>
var sRet = "-1*****0";

function form() {
	return document.forms.formname || document.forms[0];
}

function treeControl() {
	var frm = form();
	return frm && frm.elements ? frm.elements.ctlCategoryTree : document.getElementById("ctlCategoryTree");
}

function hiddenValue(name) {
	var frm = form();
	var item = frm && frm.elements ? frm.elements[name] : null;
	return item ? item.value : "";
}

function sendValue(sValue) {
	sRet = sValue || "-1*****0";
	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.returnAndClose(sRet);
		return;
	}
	if (window.ITMSModernCompat) {
		window.ITMSModernCompat.returnModalValue(sRet);
	}
	window.close();
}

function CheckSubmit() {
	var tree = treeControl();
	var classSelected = tree && tree.classification || "";
	var className = tree && tree.classificationName || tree && tree.GetText || "";
	if (classSelected === "" || classSelected.indexOf(":") === -1) {
		alert("Select Classification");
		return false;
	}
	sendValue(classSelected + "*****" + className);
	return true;
}

function Init() {
	var sIType = hiddenValue("hIType") || "NO";
	var sOrgID = hiddenValue("hOrgID") || "NO";
	var tree = treeControl();
	var hITypeName = hiddenValue("hITypeName") || "NO";
	if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
		window.ITMSModernCompat.init(document);
		tree = treeControl();
	}
	if (tree) {
		tree.IType = (sIType || "NO") + ":" + hITypeName.replace(/:/g, " - ") + ":" + (sOrgID || "NO");
	}
}

window.addEventListener("beforeunload", function () {
	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.returnValue(sRet);
	} else if (window.ITMSModernCompat) {
		window.ITMSModernCompat.returnModalValue(sRet);
	}
});
</SCRIPT>
</HEAD>
<%
	Dim sIType,sOrgID,sITypeName,sSelectMode
	sIType = trim(Request("sIType"))
	sOrgID = trim(Request("sOrgID"))
	sITypeName = trim(Request("sITypename"))
	sSelectMode = Trim(Request("SelMode"))
	if Trim(sSelectMode)="" or IsNull(sSelectMode) or sSelectMode="M" then sSelectMode = "S"

%>
<BODY leftMargin=15 topMargin=10  onLoad="Init()">
<form method="POST" name="formname" action="">
<input type=hidden name="hIType" value="<%=Server.HTMLEncode(sIType)%>">
<input type=hidden name="hOrgID" value="<%=Server.HTMLEncode(sOrgID)%>">
<input type=hidden name="hITypeName" value="<%=Server.HTMLEncode(sITypeName)%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Classification Selection
		</td>
    </tr>

	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>

	<tr>
		<td valign="top">
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border=0 width="100%">
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>

								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td align="center" colspan="3" class="MiddlePack">
												<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>

                                        <tr>
											<td>
                                                <table border="0" class="TableOutlineOnly" width="589">
                                                  <!--<tr>
                                                    <td>
														<IFRAME NAME="ifr "ID="iframe" FRAMEBORDER=0 SCROLLING=AUTO SRC="../inventory/master/comItemClassificationTree.asp" NORESIZE="RESIZE" STYLE="WIDTH=100%; HEIGHT=350"></IFRAME>
                                                  </tr>-->
                                                  <tr>


                                                    <td>
                                                        <div id="ctlCategoryTree" data-itms-tree-control data-tree-kind="item-classification"
	                                                    data-dsn="../Inventory/Components/GetCategoryGroup.asp" data-itype="<%=Server.HTMLEncode(sIType & ":" & Replace(sITypeName, ":", " - ") & ":" & sOrgID)%>"
	                                                    data-width="552px" data-height="340px"></div>
                                                    </td>

                                                  </tr>

                                                </table>
											</td>
                                        </tr>

										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>

										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
                                                              <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
															  <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="sendValue('-1*****0')">
														</td>
													</tr>
												</table>
											</td>
										</tr>

                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>

								<td align="center">
                                    <img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<INPUT TYPE=HIDDEN NAME="hClassSelected" VALUE="">
</form>
</BODY>
</HTML>
