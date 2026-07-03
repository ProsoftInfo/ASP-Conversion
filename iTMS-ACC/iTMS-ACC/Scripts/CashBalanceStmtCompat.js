(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback || "";
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
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function loadXml(name, xmlText) {
		var object = xmlObject(name);
		var doc;
		if (object && typeof object.loadXML === "function") {
			object.loadXML(xmlText || "<Root/>");
			return object;
		}
		doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		if (object) {
			object._doc = doc;
		}
		return object || doc;
	}

	function serializeXml(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		return new XMLSerializer().serializeToString(nodeOrDoc);
	}

	function childElements(node, name) {
		var wanted = name ? String(name).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
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

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var parts;
		if (window.GetWindowSizeForPopup) {
			parts = String(window.GetWindowSizeForPopup(type)).split(":");
			if (parts.length >= 3) {
				return { program: parts[0], height: parts[1], width: parts[2] };
			}
		}
		return { program: fallbackProgram, height: fallbackHeight, width: fallbackWidth };
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		if (callback) {
			callback(window.showModalDialog ? window.showModalDialog(url, args || "", features || "") : "");
			return;
		}
		window.open(url, "_blank", "width=700,height=450,resizable=yes,status=no");
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate ? window.ITMSModernCompat.toIsoDate(value) : value;
		}
	}

	function getDateControl(name) {
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
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return control.value || "";
	}

	function parseDate(value) {
		var text = trim(value);
		var parts;
		if (!text) {
			return null;
		}
		if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
			parts = text.split("-");
			return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
		}
		parts = text.split(/[\/\-]/);
		if (parts.length >= 3) {
			return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
		}
		return null;
	}

	function todayDisplay() {
		var date = new Date();
		var day = String(date.getDate()).padStart(2, "0");
		var month = String(date.getMonth() + 1).padStart(2, "0");
		return day + "/" + month + "/" + date.getFullYear();
	}

	function fillPartyTypes() {
		var select = field("selAccHead");
		var xhr;
		var root;
		if (!select) {
			return;
		}
		while (select.options.length > 2) {
			select.remove(2);
		}
		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(valueOf("hUnitNo")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("OutData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("OutData", xhr.responseText);
		}
		root = xmlRoot("OutData");
		childElements(root).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, "ParType") || attr(node, 0)));
		});
	}

	function fillBooks() {
		var select = field("selBook");
		var xhr;
		var root;
		if (!select) {
			return;
		}
		select.options.length = 1;
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=01&orgID=" + encodeURIComponent(valueOf("hUnitNo")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		root = xmlRoot("UnitBookData");
		childElements(root).forEach(function (node) {
			select.add(new Option(attr(node, 1), attr(node, 0)));
		});
	}

	function continueSelection(url, args, features, callback) {
		openDialog(url, args, features, function (value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var action = String(attr(root, "Action")).toUpperCase();
			var query = trim(attr(root, "PassQuery"));
			if (action && action !== "DONE" && action !== "CLOSE" && query) {
				continueSelection(url.replace(/\?.*$/, "?" + query), args, features, callback);
				return;
			}
			callback(root || value);
		});
	}

	function finishAccHeadSelection(root) {
		var nodes = childElements(root || xmlRoot("AccHeadData"), "Entry");
		var codes = [];
		var names = [];
		nodes.forEach(function (node) {
			codes.push(attr(node, "RetField0"));
			names.push(attr(node, "RetField5"));
		});
		if (codes.length) {
			setValue("hAccHead", codes.join(","));
			setValue("txtAccHead", names.join(","));
		} else {
			field("selAccHead").selectedIndex = 0;
			setValue("hAccHead", "0");
		}
	}

	function finishPartySelection(value) {
		var parts = String(value || "").split(":");
		var partyValue;
		if (parts.length <= 2) {
			field("selAccHead").selectedIndex = 0;
			field("selAccHead").focus();
			return;
		}
		partyValue = (parts[4] || "") + "?" + (parts[3] || "") + "?" + (parts[0] || "") + "?" + (parts[1] || "");
		setValue("hAccHead", partyValue);
		setValue("txtAccHead", parts[0] || "");
	}

	window.ResetAccHead = function () {
		setValue("hAccHead", "");
		setValue("txtAccHead", "");
		setValue("hAccIndex", "");
		if (field("selAccHead")) {
			field("selAccHead").value = "0";
		}
	};

	window.DisplayBook = function () {
		fillPartyTypes();
		fillBooks();
	};

	window.GetBookNo = function () {
		var select = field("selBook");
		setValue("hBookNo", select ? select.selectedIndex : 0);
		setValue("hBookVal", select ? select.value : "");
	};

	window.SelectAccHead = function () {
		var select = field("selAccHead");
		var book = field("selBook");
		var size;
		var partyType;
		if (!select || !book) {
			return false;
		}
		if (book.value === "S") {
			alert("Select Book");
			book.focus();
			select.selectedIndex = 0;
			return false;
		}
		if (select.value === "0") {
			return false;
		}
		setValue("hAccHead", "");
		setValue("txtAccHead", "");
		if (select.value === "G") {
			size = popupSize("5", "GLHeadSelection.asp", "500", "550");
			continueSelection(
				"../../Common/" + size.program + "?orgID=" + encodeURIComponent(valueOf("hUnitNo")) + "&BookId=01&BookNo=" + encodeURIComponent(book.value) + "&hSelectMode=M",
				xmlObject("AccHeadData"),
				"dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No",
				finishAccHeadSelection
			);
		} else {
			size = popupSize("12", "PartySelection.asp", "500", "420");
			partyType = select.value + "?" + select.options[select.selectedIndex].text;
			continueSelection(
				"../../Common/" + size.program + "?orgID=" + encodeURIComponent(valueOf("hUnitNo")) + "&Party=" + encodeURIComponent(partyType),
				"",
				"dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No",
				finishPartySelection
			);
		}
		setValue("hAccIndex", select.selectedIndex);
		setValue("hAccTxt", valueOf("txtAccHead"));
		return false;
	};

	window.GetFormDet = function () {
		var book = field("selBook");
		var formValue = valueOf("hUnitNo");
		formValue += "|" + (book ? book.value : "");
		formValue += "|" + (book ? book.selectedIndex : "");
		formValue += "|" + getDateControl("ctlVouFromDate");
		formValue += "|" + getDateControl("ctlVouToDate");
		formValue += "|" + valueOf("hAccIndex");
		formValue += "|" + (trim(valueOf("hAccIndex")) === "" ? "" : valueOf("hAccHead"));
		formValue += "|" + valueOf("txtAccHead");
		formValue += "|" + (window.sVouType || "");
		setValue("hFormVal", formValue);
	};

	window.Validate = function () {
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		if (field("selBook") && field("selBook").selectedIndex === 0) {
			alert("Select Cash Book");
			field("selBook").focus();
			return false;
		}
		if (from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}
		setValue("hFromDate", fromDate);
		setValue("hToDate", toDate);
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.ChkReset = function () {
		setDateControl("ctlVouFromDate", todayDisplay());
		setDateControl("ctlVouToDate", todayDisplay());
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
			field("selAccHead").disabled = true;
		}
		if (field("selBook")) {
			field("selBook").selectedIndex = 0;
		}
		["txtFromAmount", "txtToAmount", "txtVouNoFrom", "txtVouNoTo", "txtAccHead", "hFlag"].forEach(function (name) {
			setValue(name, "");
		});
		window.DisplayBook();
	};

	window.SetDate = function () {
		if (trim(valueOf("hFromDate")) && trim(valueOf("hToDate"))) {
			setDateControl("ctlVouFromDate", valueOf("hFromDate"));
			setDateControl("ctlVouToDate", valueOf("hToDate"));
		}
		window.DisplayBook();
		if (field("selBook") && field("selBook").options.length > 1) {
			field("selBook").selectedIndex = Number(valueOf("hBookNo", 0)) || 0;
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = Number(valueOf("hAccIndex", 0)) || 0;
			}
			setValue("txtAccHead", valueOf("hAccTxt"));
		} else if (field("selBook")) {
			field("selBook").selectedIndex = 0;
		}
	};

	window.PrintWindow = function () {
		var pass = trim(valueOf("hUnitNo")) + "|" + valueOf("hUnitName") + "| |" + trim(valueOf("hBookVal")) + "|VouDate|" + valueOf("hFromDate") + "|" + valueOf("hToDate") + "|B";
		openDialog("../Reports/PRNDBCashStmentView.asp?Value=" + encodeURIComponent(pass), "A", "dialogHeight:150px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};
}(window, document));
