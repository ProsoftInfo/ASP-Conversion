(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function controls(name) {
		var frm = form();
		var item = frm && (frm.elements[name] || frm[name]);
		if (!item) {
			return [];
		}
		if (item.length && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function field(name) {
		return controls(name)[0] || null;
	}

	function setValue(name, value) {
		controls(name).forEach(function (item) {
			item.value = value == null ? "" : String(value);
		});
	}

	function submitForm() {
		var frm = form();
		if (frm) {
			frm.submit();
		}
	}

	function encodeParam(value) {
		return encodeURIComponent(value == null ? "" : String(value));
	}

	function done(value) {
		var root = value && (value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value);
		return !!(root && typeof root.getAttribute === "function" && root.getAttribute("Done") === "Y");
	}

	function openDialog(url, features, callback) {
		var popup;
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			popup = window.ITMSModernCompat.openModalDialog(url, "", features, callback || function () {});
			return !!popup;
		}
		popup = window.open(url, "_blank", "width=600,height=500,resizable=yes,scrollbars=yes");
		if (!popup) {
			alert("Popup was blocked. Please allow popups for this site and try again.");
			return false;
		}
		return true;
	}

	function sequentialSelections(prefix) {
		var countField = field("hCnt");
		var count = parseInt(countField && countField.value, 10) || 0;
		var result = [];
		var item;
		var i;
		for (i = 1; i <= count; i += 1) {
			item = field(prefix + i) || field(prefix.toLowerCase() + i);
			if (item && item.checked && trim(item.value) !== "") {
				result.push(item.value);
			}
		}
		return result;
	}

	function alertRoleActivityMap(iCount, nRoleID, nActivityCode) {
		alert(iCount + "  " + nRoleID + " " + nActivityCode);
	}

	function applicationRolesPaginate(nPage) {
		setValue("hPageSelection", nPage);
		setValue("selRole", field("hRoleID").value);
		submitForm();
	}

	function gotoRoleAction(sPara) {
		var selected = sequentialSelections("Chkbox");
		var para = trim(sPara);
		var roleId = selected.length ? selected[selected.length - 1] : "";
		var passData;
		if (selected.length > 1) {
			alert("Select any one Role");
			return false;
		}
		if (selected.length === 0 && para !== "CRN") {
			alert("Select Role");
			return false;
		}
		passData = para + ":" + roleId;
		if (para === "CRN" || para === "EDT") {
			openDialog("AppRoleCreationAndAmendEntry.asp?PassData=" + encodeParam(passData), "dialogHeight:200px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (value) {
				setValue("selRole", field("hRoleID").value);
				if (done(value)) {
					submitForm();
				}
			});
		}
		return false;
	}

	function showApplicationRoleActivityMap(nRoleID, sRoleName) {
		var passData = nRoleID + ":" + sRoleName;
		openDialog("AppRoleActivityMappingPopUp.asp?PassData=" + encodeParam(passData), "dialogHeight:400px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No;scroll:YES", function () {
			submitForm();
		});
		return false;
	}

	function assignRoleID(nRoleID) {
		setValue("hRoleID", nRoleID);
		setValue("selProcess", "0");
		submitForm();
	}

	function assignProcessCode(nProcessCode) {
		setValue("hProcessCode", nProcessCode);
		submitForm();
	}

	function deleteRole() {
		var selected = sequentialSelections("Chkbox");
		var roleId = selected.length ? selected[selected.length - 1] : "";
		var frm = form();
		if (!roleId) {
			alert("Select any one Role For Deletion");
			return false;
		}
		frm.action = "RoleDeletion.asp?RoleID=" + encodeParam(roleId);
		submitForm();
		return false;
	}

	function validateActivity() {
		setValue("hProcessName", trim(field("selProcess").value));
		setValue("hPracticeName", trim(field("selActivityType").value));
		setValue("hActivityName", trim(field("txtActivityName").value));
		setValue("selApplication", trim(field("selProcess").value));
		submitForm();
	}

	function childElements(node) {
		var result = [];
		var i;
		for (i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function attrAt(node, index) {
		var item = node && node.attributes && node.attributes.item(index);
		return item ? item.value : "";
	}

	function responseRoot(xhr) {
		var doc = xhr.responseXML;
		if (!doc || !doc.documentElement) {
			doc = new DOMParser().parseFromString(xhr.responseText || "<Root/>", "text/xml");
		}
		return doc.documentElement;
	}

	function populatePractice(value) {
		var process = typeof value === "string" ? value : value && value.value;
		var practice = field("selActivityType");
		var xhr;
		var root;
		if (process === "0") {
			alert("Select Prcoess Name");
			field("selProcess").focus();
			return false;
		}
		if (!practice) {
			return false;
		}
		practice.options.length = 1;
		xhr = new XMLHttpRequest();
		xhr.open("GET", "XMLSelect.asp?sWho=PR&sProcess=" + encodeParam(process), false);
		xhr.send(null);
		if (trim(xhr.responseText) === "") {
			alert("No Practice defined for the Process Selected");
			field("selProcess").focus();
			return false;
		}
		root = responseRoot(xhr);
		childElements(root).forEach(function (node) {
			practice.options[practice.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
		});
		return true;
	}

	function chkReset() {
		if (field("TxtProcessName")) {
			setValue("TxtProcessName", "");
		}
		setValue("txtActivityName", "");
		submitForm();
	}

	function activityPaginate(nPage) {
		setValue("hPageSelection", nPage);
		setValue("selApplication", field("hAppCode").value);
		setValue("selProcess", field("hAppCode").value);
		submitForm();
	}

	function selectedActivityParts(requireSelection, multiMessage, emptyMessage) {
		var selected = sequentialSelections("Chkbox");
		if (selected.length > 1) {
			alert(multiMessage);
			return null;
		}
		if (requireSelection && selected.length === 0) {
			alert(emptyMessage);
			return null;
		}
		return selected.length ? selected[0].split(":") : [];
	}

	function showRoleActivityMapOld(sCallFrom, sTempVal) {
		var callFrom = trim(sCallFrom);
		var parts = selectedActivityParts(callFrom !== "ADD", "Select any one Activty", "Select Activity");
		var passData;
		if (!parts) {
			return false;
		}
		if (callFrom === "ADD") {
			passData = callFrom + ":" + sTempVal;
		} else if (callFrom === "EDT") {
			passData = callFrom + ":" + sTempVal + ":" + (parts[0] || "") + ":" + (parts[1] || "");
		}
		if (passData) {
			openDialog("ActivityMappingPopUp.asp?PassData=" + encodeParam(passData), "dialogHeight:220px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (value) {
				setValue("selApplication", field("hAppCode").value);
				if (done(value)) {
					submitForm();
				}
			});
		}
		return false;
	}

	function showRoleActivityMap(sCallFrom, sTempVal) {
		var passData = trim(sCallFrom) + ":" + sTempVal;
		openDialog("ActivityMappingPopUp.asp?PassData=" + encodeParam(passData), "dialogHeight:450px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (value) {
			if (done(value)) {
				submitForm();
			}
		});
		return false;
	}

	function showDetails(sType, sTemp) {
		var passData = sTemp + ":" + sType;
		openDialog("AppActivityDetailsPopUp.asp?PassData=" + encodeParam(passData), "dialogHeight:320px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (value) {
			if (done(value)) {
				submitForm();
			}
		});
		return false;
	}

	function delActivity() {
		var parts = selectedActivityParts(true, "Select any one Activty For Delete", "Select Activity For Delete");
		var frm = form();
		var passData;
		if (!parts) {
			return false;
		}
		passData = (parts[4] || "") + ":" + (parts[2] || "") + ":" + (parts[3] || "") + ":" + (parts[0] || "");
		frm.action = "ActivityDelete.asp?sData=" + encodeParam(passData);
		submitForm();
		return false;
	}

	function assignApplicationCode(nAppcode) {
		setValue("hAppCode", nAppcode);
		submitForm();
	}

	function installApplicationRoles() {
		window.Paginate = applicationRolesPaginate;
		window.popRoleActivityMap = alertRoleActivityMap;
		window.GotoAction = gotoRoleAction;
		window.ShowRoleActivityMap = showApplicationRoleActivityMap;
		window.AssaignRoleID = assignRoleID;
		window.AssaignProcessCode = assignProcessCode;
		window.DeleteRole = deleteRole;
	}

	function installActivityCreationMain() {
		window.Validate = validateActivity;
		window.PopulatePractice = populatePractice;
		window.ChkReset = chkReset;
		window.Paginate = activityPaginate;
		window.popRoleActivityMap = alertRoleActivityMap;
		window.ShowRoleActivityMapOld = showRoleActivityMapOld;
		window.ShowRoleActivityMap = showRoleActivityMap;
		window.ShowDetails = showDetails;
		window.DelActivity = delActivity;
		window.AssaignApplicationCode = assignApplicationCode;
	}

	window.ITMSAdminRoleActivityCompat = {
		installApplicationRoles: installApplicationRoles,
		installActivityCreationMain: installActivityCreationMain
	};
}(window, document));
