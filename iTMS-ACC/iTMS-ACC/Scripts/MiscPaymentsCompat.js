(function (window, document) {
	"use strict";

	var sFlag = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		var target = String(name).toLowerCase();
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				return frm.elements[index];
			}
		}
		return document.getElementsByName(name)[0] || document.getElementById(name) || null;
	}

	function fields(name) {
		var item = field(name);
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return item ? [item] : [];
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item && item.value != null ? item.value : fallback || "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function setHtml(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function checkedValue(name) {
		var found = fields(name).filter(function (item) {
			return item.checked;
		})[0];
		return found ? found.value : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function request(method, url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open(method, url, false);
		xhr.send(body || null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function popupInfo(type, fallback) {
		var text = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(String(type)) : "";
		var parts = String(text || fallback || "").split(":");
		return {
			program: parts[0] || "",
			height: parts[1] || "500",
			width: parts[2] || "500",
			features: "dialogHeight:" + (parts[1] || "500") + "px;dialogWidth:" + (parts[2] || "500") + "px;Status:No"
		};
	}

	function getDate(name) {
		var control = field(name) || byId(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate ? window.ITMSModernCompat.toDisplayDate(control.value) : control.value || "";
	}

	function setDate(name, value) {
		var control = field(name) || byId(name);
		if (!control) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = value || "";
		}
	}

	function parseDate(value) {
		var match = trim(value).match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		var year;
		if (!match) {
			return null;
		}
		year = Number(match[3]);
		if (year < 100) {
			year += 2000;
		}
		return new Date(year, Number(match[2]) - 1, Number(match[1]));
	}

	function addOption(select, text, value) {
		if (select) {
			select.options[select.options.length] = new Option(text || "", value || "");
		}
	}

	function populateBookSelect(xmlDoc) {
		var select = field("selBook");
		var root = xmlDoc && xmlDoc.documentElement;
		childElements(root).forEach(function (node) {
			addOption(select, attr(node, 1), attr(node, 0));
		});
	}

	function populatePartyTypes(xmlDoc) {
		var select = field("SelAccHead");
		var root = xmlDoc && xmlDoc.documentElement;
		childElements(root).forEach(function (node) {
			addOption(select, node.textContent || "", attr(node, "ParType"));
		});
	}

	function applyAccHeadSelection(name, value) {
		setValue("hAccHead", value || "0");
		setHtml("spAccHead", name ? name + "&nbsp;" : "");
	}

	function retFromDialogRoot(root) {
		var entry = childElements(root, "Entry")[0] || childElements(root)[0];
		if (!entry) {
			return "";
		}
		return [0, 1, 2, 3, 4, 5, 6].map(function (index) {
			return attr(entry, "RetField" + index);
		}).join(":");
	}

	function runXmlDialog(baseUrl, args, xmlName, info, callback) {
		function handle(value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var action = root ? trim(attr(root, "Action")).toUpperCase() : "";
			var query = root ? trim(attr(root, "PassQuery")) : "";
			if (action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE" && query) {
				openDialog(baseUrl + "?" + query, xmlObject(xmlName), info.features, handle);
				return;
			}
			callback(root);
		}
		openDialog(baseUrl + "?" + args, xmlObject(xmlName), info.features, handle);
	}

	window.ChkSubmit = function () {
		var count = Number(valueOf("hCnt", 0));
		var selected = null;
		var selectedCount = 0;
		var i;
		var box;
		var parts;
		for (i = 1; i <= count; i += 1) {
			box = field("ChkMiscZ" + i);
			if (box && box.checked) {
				selected = box.value;
				selectedCount += 1;
			}
		}
		if (selectedCount !== 1) {
			alert("Select any one to create");
			return false;
		}
		parts = String(selected || "").split("Z");
		if (parts[4] !== "010101") {
			alert("Already Entry Created for the Invoice");
			return false;
		}
		form().action = (parts[0] === "C" ? "MsiVouEntry.asp" : "MsiVouEntryForBank.asp") +
			"?TransNo=" + encodeURIComponent(parts[1] || "") +
			"&OrgName=" + encodeURIComponent(parts[2] || "") +
			"&OrgID=" + encodeURIComponent(parts[3] || "");
		form().submit();
		return false;
	};

	window.AssignPage = function (page) {
		setValue("hPage", page);
		form().action = "MsiVouBookSelection.asp";
		form().submit();
		return false;
	};

	window.popPartType = function () {
		var xhr = request("GET", "XMLGetOrgParType.asp?orgID=" + encodeURIComponent(valueOf("hUnitId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			populatePartyTypes(xhr.responseXML);
		}
	};

	window.DisplayBook = function () {
		var select = field("selBook");
		var xhr;
		if (select) {
			select.options.length = 1;
		}
		xhr = request("GET", "XMLGetOrgBook.asp?BkCode=" + encodeURIComponent(valueOf("selVoucher")) + "&orgID=" + encodeURIComponent(valueOf("hUnitId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			populateBookSelect(xhr.responseXML);
		}
	};

	window.Validate = window.validate = function () {
		var criteria = checkedValue("optCriteria") || valueOf("hoptCriteria");
		var fromDate = getDate("ctlVouFromDate");
		var toDate = getDate("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		var fromAmount = trim(valueOf("txtGAmount"));
		var toAmount = trim(valueOf("txtLAmount"));
		if (criteria === "VouDate" && from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}
		if (criteria === "Amount") {
			if (fromAmount === "") {
				alert("Enter From Amount");
				if (field("txtGAmount")) {
					field("txtGAmount").select();
				}
				return false;
			}
			if (toAmount === "") {
				alert("Enter To Amount");
				if (field("txtLAmount")) {
					field("txtLAmount").select();
				}
				return false;
			}
			if (isNaN(Number(fromAmount)) || isNaN(Number(toAmount))) {
				alert("Enter Numbers Only");
				return false;
			}
			if (Number(fromAmount) > Number(toAmount)) {
				alert("To Amount Should be Greater Than From Amount ");
				setValue("txtLAmount", "");
				if (field("txtLAmount")) {
					field("txtLAmount").select();
				}
				return false;
			}
		}
		setValue("hoptCriteria", criteria);
		setValue("horgName", valueOf("hOrgName"));
		setValue("hFDate", fromDate);
		setValue("hTDate", toDate);
		form().action = "MiscPayments.asp";
		form().submit();
		return false;
	};

	window.OptSelection = function () {
		var criteria = checkedValue("optCriteria");
		var amountMode = criteria === "Amount";
		sFlag = criteria;
		setValue("txtGAmount", "");
		setValue("txtLAmount", "");
		if (field("txtGAmount")) {
			field("txtGAmount").readOnly = !amountMode;
		}
		if (field("txtLAmount")) {
			field("txtLAmount").readOnly = !amountMode;
		}
		setValue("hoptCriteria", criteria);
	};

	window.SelNew = function () {
		if (field("SelAccHead")) {
			field("SelAccHead").selectedIndex = 0;
		}
		["txtGAmount", "txtLAmount", "txtNoFrom", "txtNoTo"].forEach(function (name) {
			setValue(name, "");
		});
		setHtml("spAccHead", "");
	};

	window.SelectAccHead = function () {
		var select = field("SelAccHead");
		var orgId = valueOf("hUnitId");
		var bookId = valueOf("selVoucher");
		var bookNo = valueOf("selBook");
		var info;
		var partyType;
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		if (bookId === "0") {
			alert("Select Voucher Type");
			if (field("selVoucher")) {
				field("selVoucher").focus();
			}
			select.selectedIndex = 0;
			return false;
		}
		if (bookNo === "S") {
			alert("Select Book");
			if (field("selBook")) {
				field("selBook").focus();
			}
			select.selectedIndex = 0;
			return false;
		}
		if (select.value === "G") {
			info = popupInfo("5", "GLHeadSelection.asp:500:500");
			runXmlDialog("../../Common/" + info.program, "orgID=" + encodeURIComponent(orgId) + "&BookId=" + encodeURIComponent(bookId) + "&BookNo=" + encodeURIComponent(bookNo), "TempXMLData", info, function (root) {
				var ret = retFromDialogRoot(root);
				var accNode;
				clearChildren(xmlRoot("AccHeadData"));
				if (ret && typeof window.GetGlHeadXml === "function") {
					window.GetGlHeadXml(ret);
				}
				accNode = childElements(xmlRoot("AccHeadData"), "AccHead")[0];
				if (accNode) {
					applyAccHeadSelection(attr(accNode, "Name"), attr(accNode, "No"));
				} else {
					select.selectedIndex = 0;
					applyAccHeadSelection("", "0");
				}
			});
		} else {
			info = popupInfo("2", "PartySelectionAcc.asp:500:500");
			partyType = select.value + "?" + selectedText(select);
			runXmlDialog("../../Common/" + info.program, "orgid=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType), "PartyData", info, function (root) {
				var entry = childElements(root, "Entry")[0] || childElements(root)[0];
				var partyName = attr(entry, "RetField0");
				var partyCode = attr(entry, "RetField1");
				var partyTypeValue = attr(entry, "RetField3");
				var partySubType = attr(entry, "RetField4");
				var xhr;
				var accNode;
				if (!entry) {
					select.selectedIndex = 0;
					applyAccHeadSelection("", "0");
					return;
				}
				xhr = request("GET", "XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(orgId) + "&ParSubType=" + encodeURIComponent(partySubType) + "&ParType=" + encodeURIComponent(partyTypeValue) + "&PartyCode=" + encodeURIComponent(partyCode));
				clearChildren(xmlRoot("AccHeadData"));
				if (trim(xhr.responseText) !== "" && typeof window.GetPartyHeadXml === "function") {
					window.GetPartyHeadXml(partyCode, partyName, xhr.responseText);
				}
				accNode = childElements(xmlRoot("AccHeadData"), "AccHead")[0];
				if (accNode) {
					applyAccHeadSelection(attr(accNode, "Name"), partyType + "?" + attr(accNode, "No"));
				} else {
					select.selectedIndex = 0;
					applyAccHeadSelection("", "0");
				}
			});
		}
		return false;
	};

	window.SetDate = function () {
		var criteria = valueOf("hoptCriteria");
		ensureCompat();
		setDate("ctlVouFromDate", valueOf("hFDate"));
		setDate("ctlVouToDate", valueOf("hTDate"));
		if (criteria) {
			fields("optCriteria").forEach(function (item) {
				item.checked = item.value === criteria;
			});
			window.OptSelection();
		}
	};

	window.MinDate = function () {
		var from = parseDate(getDate("ctlVouFromDate"));
		var to = parseDate(getDate("ctlVouToDate"));
		var min = parseDate(valueOf("hFDate"));
		var max = parseDate(valueOf("hTDate"));
		if (from && min && max && (from < min || from > max)) {
			alert("Date Should be within the Financial Year  " + valueOf("hFDate") + " to " + valueOf("hTDate"));
			setDate("ctlVouFromDate", valueOf("hFDate"));
			return false;
		}
		if (to && min && max && (to < min || to > max)) {
			alert("Date Should be within the Financial Year  " + valueOf("hFDate") + " to " + valueOf("hTDate"));
			setDate("ctlVouToDate", valueOf("hTDate"));
			return false;
		}
		return true;
	};

	window.ChkReset = function () {
		setDate("ctlVouFromDate", valueOf("hFDate"));
		setDate("ctlVouToDate", valueOf("hTDate"));
		fields("optCriteria").forEach(function (item) {
			item.checked = false;
		});
		["txtGAmount", "txtLAmount", "hoptCriteria"].forEach(function (name) {
			setValue(name, "");
		});
		if (field("txtGAmount")) {
			field("txtGAmount").readOnly = true;
		}
		if (field("txtLAmount")) {
			field("txtLAmount").readOnly = true;
		}
	};
}(window, document));
