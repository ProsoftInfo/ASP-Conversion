(function (window, document) {
	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function upper(value) {
		return trim(value).toUpperCase();
	}

	function compat() {
		return window.ITMSModernCompat;
	}

	function openDialog(url, args, features, callback) {
		if (!compat() || !compat().openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		compat().openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function xmlText(node) {
		if (!node) {
			return "";
		}
		if (typeof node.xml === "string") {
			return node.xml;
		}
		if (window.XMLSerializer) {
			return new XMLSerializer().serializeToString(node);
		}
		return "";
	}

	function loadOutData(node, target) {
		var xml = xmlText(node);
		var data = target || window.OutData;
		if (xml && data && typeof data.loadXML === "function") {
			data.loadXML(xml);
		}
	}

	function form() {
		return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return document.getElementById(name) || frm && frm.elements && frm.elements[name] || window[name] || null;
	}

	function setFieldText(name, value) {
		var item = field(name);
		if (!item) {
			return;
		}
		if ("value" in item) {
			item.value = value == null ? "" : String(value);
		} else {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(String(type)) : "";
		var parts = String(value || "").split(":");
		return {
			program: parts[0] || fallbackProgram,
			height: parts[1] || fallbackHeight || "500",
			width: parts[2] || fallbackWidth || "750"
		};
	}

	function loopModal(baseUrl, args, features, callback) {
		openDialog(baseUrl, args, features, function handleResult(node) {
			if (!node) {
				if (callback) {
					callback(node);
				}
				return;
			}
			var action = upper(getAttr(node, "Action"));
			if (action && action !== "DONE" && action !== "CLOSE") {
				var passQuery = getAttr(node, "PassQuery");
				var nextUrl = baseUrl.split("?")[0] + "?" + passQuery;
				openDialog(nextUrl, args, features, handleResult);
				return;
			}
			if (callback) {
				callback(node);
			}
		});
	}

	window.RefTypeSelectionSupp = function (refType, orgId, partyCode, stock, flag, addButton, dispItem, callback) {
		if (trim(refType) !== "N") {
			loopModal(
				"/Common/DynamicNoSelection.asp?orgID=" + encodeURIComponent(orgId) + "&RefType=" + encodeURIComponent(refType) + "&ParCode=" + encodeURIComponent(partyCode),
				window.OutData,
				"dialogHeight:500px;dialogWidth:700px;status:no",
				function (node) {
					loadOutData(node);
					if (callback) {
						callback(node);
					}
				}
			);
			return;
		}
		loopModal(
			"../../Common/SuppItemSelectCommon.asp?orgID=" + encodeURIComponent(orgId) + "&Stock=" + encodeURIComponent(stock) + "&hSelectMode=M&Flag=" + encodeURIComponent(flag) + "&hDispButt=" + encodeURIComponent(addButton) + "&hDispItem=" + encodeURIComponent(dispItem) + "&hPartyCode=" + encodeURIComponent(partyCode),
			window.OutData,
			"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No",
			callback
		);
	};

	window.RefTypeSelection = function (refType, orgId, partyCode, stock, flag, addButton, dispItem, callFrom, callback) {
		var size;
		if (trim(refType) !== "N") {
			size = popupSize("3", "DynamicNoSelection.asp", "500", "850");
			loopModal(
				"/Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&RefType=" + encodeURIComponent(refType) + "&ParCode=" + encodeURIComponent(partyCode),
				window.OutData,
				"dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No",
				function (node) {
					loadOutData(node);
					if (callback) {
						callback(node);
					}
				}
			);
			return;
		}

		size = popupSize("1", "ItemSelectRelPartyCommon.asp", "500", "850");
		openDialog(
			"../../Common/" + size.program + "?orgID=" + encodeURIComponent(window.sUnit || orgId || "") +
				"&sIType=" + encodeURIComponent(window.sIType || "") +
				"&Stock=" + encodeURIComponent(stock) +
				"&hSelectMode=M&Flag=" + encodeURIComponent(flag) +
				"&hDispButt=" + encodeURIComponent(addButton) +
				"&hDispItem=" + encodeURIComponent(dispItem) +
				"&CallFrom=" + encodeURIComponent(callFrom || ""),
			window.OutData,
			"dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No",
			callback
		);
	};

	window.MixSelection = function (orgId, callback) {
		loopModal(
			"../../Common/MixSelectCommon.asp?orgID=" + encodeURIComponent(orgId) + "&hSelectMode=M",
			window.MixData,
			"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No",
			callback
		);
	};

	window.popIssueTo = function () {
		var form = document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
		var issueSelect = form && form.selIssueTo;
		var issueValue = issueSelect && issueSelect.options ? issueSelect.options[issueSelect.selectedIndex].value : "";
		var lower = trim(issueValue).toLowerCase();
		var size = popupSize("2", "PartySelection.asp", "500", "500");

		if (lower === "party") {
			openDialog(
				"/Common/" + size.program + "?orgID=" + encodeURIComponent(form.hUnit.value),
				window.PartyData,
				"dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No",
				function processParty(node) {
					if (!node || upper(getAttr(node, "Action")) === "CLOSE") {
						return;
					}
					var action = upper(getAttr(node, "Action"));
					if (action && action !== "DONE") {
						openDialog("/Common/" + size.program + "?" + getAttr(node, "PassQuery"), window.PartyData, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", processParty);
						return;
					}
					Array.prototype.forEach.call(node.childNodes || [], function (child) {
						setFieldText("txtParty", getAttr(child, "RetField0"));
						form.hIssueToType.value = "Party";
						form.hIssueToCode.value = getAttr(child, "RetField1");
					});
				}
			);
			return;
		}

		if (lower === "dept:prd") {
			openDialog("/Common/WorkCenterPopup.asp", "", "dialogHeight:150px;dialogWidth:300px;", function (workCenter) {
				var parts = issueValue.split(":");
				Array.prototype.some.call((workCenter && workCenter.childNodes) || [], function (child) {
					form.hIssueToType.value = parts[0] || "";
					form.hIssueToCode.value = parts[1] || "";
					form.hIssueToSubCode.value = getAttr(child, "Code");
					setFieldText("txtParty", getAttr(child, "Name"));
					return true;
				});
			});
			return;
		}

		if (lower === "unit" || lower === "pos") {
			alert("Select the Sub Level");
			issueSelect.selectedIndex = 0;
			return;
		}

		var parts = issueValue.split(":");
		form.hIssueToType.value = parts[0] || "";
		form.hIssueToCode.value = parts[1] || "";
	};

	window.popIssueToWithOutSubLevel = function () {
		var form = document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
		var issueSelect = form && form.selIssueTo;
		var issueValue = issueSelect && issueSelect.options ? issueSelect.options[issueSelect.selectedIndex].value : "";
		var lower = trim(issueValue).toLowerCase();
		if (lower === "party") {
			form.hIssueToType.value = "Party";
			form.hIssueToCode.value = "";
			return;
		}
		if (lower === "unit" || lower === "pos") {
			form.hIssueToType.value = lower;
			form.hIssueToCode.value = "";
			return;
		}
		var parts = issueValue.split(":");
		form.hIssueToType.value = parts[0] || "";
		form.hIssueToCode.value = parts[1] || "";
	};
})(window, document);
