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
		var target;
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		target = String(name).toLowerCase();
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				return frm.elements[index];
			}
		}
		return null;
	}

	function fields(name) {
		var item = field(name);
		var frm = form();
		var result = [];
		var target;
		var index;
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		if (item) {
			return [item];
		}
		if (!frm || !frm.elements) {
			return result;
		}
		target = String(name).toLowerCase();
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				result.push(frm.elements[index]);
			}
		}
		return result;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		if (!item || item.value == null) {
			return fallback || "";
		}
		return item.value;
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

	function xmlDocument(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return doc && doc.documentElement || object && object.documentElement || null;
	}

	function serializeXml(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		try {
			return new XMLSerializer().serializeToString(nodeOrDoc);
		} catch (ignore) {
			return "";
		}
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

	function childElements(node) {
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1;
		});
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function createElement(root, name) {
		var doc = root && root.ownerDocument || xmlDocument("SearchData") || document.implementation.createDocument("", "", null);
		return doc.createElement(name);
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

	function syncPost(url, payload) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(payload || "");
		return xhr;
	}

	function saveXmlIsland(name, url) {
		var object = xmlObject(name);
		var doc = xmlDocument(object);
		var root = xmlRoot(object);
		syncPost(url, serializeXml(doc || root));
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

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || value == null || value === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
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

	function dateYearMonth(name) {
		var parsed = parseDate(getDateControl(name));
		if (!parsed) {
			return "";
		}
		return String(parsed.getFullYear()) + String(parsed.getMonth() + 1).padStart(2, "0");
	}

	function finPeriodBounds() {
		var period = valueOf("hFinPeriod");
		var parts = period.split(":");
		var start = parts[0] || period.substring(0, 4);
		var end = parts[1] || period.slice(-4);
		return {
			fromText: "01/04/" + start,
			toText: "31/03/" + end,
			min: start + "04",
			max: end + "03"
		};
	}

	function pageKind() {
		var grid = valueOf("hGridName");
		if (/purchase/i.test(grid) || /purchase/i.test(location.pathname)) {
			return "purchase";
		}
		return "sales";
	}

	function config() {
		var purchase = pageKind() === "purchase";
		return {
			isPurchase: purchase,
			bookCode: purchase ? "04" : "05",
			module: purchase ? "PUR" : "SAL",
			viewPage: purchase ? "PurchaseVouchView_San.asp" : "SalesVouchView_San.asp",
			listPage: purchase ? "PurchaseVouchers.asp" : "SalesVouchers.asp",
			otherAppNo: purchase ? "2" : "3",
			otherAppPage: purchase ? "AppOtherPURViewWithMulTax.asp" : "AppOtherSALView.asp",
			editPage: purchase ? "PURCHASEVOUCHERENTRYEDIT.ASP" : "SALESVOUCHERENTRYEDIT.asp",
			deletePage: purchase ? "PurVouDeletion.asp" : "SalVouDeletion.asp",
			accountPage: purchase ? "AccPurVouGenerate.asp" : "AccSalVouGenerate.asp",
			typeField: purchase ? "selPurType" : "selSalType",
			typeAttr: purchase ? "SelPurType" : "selSalType",
			title: purchase ? "Purchase Vouchers" : "Sales Vouchers"
		};
	}

	function listAction() {
		var cfg = config();
		if (cfg.isPurchase) {
			return cfg.listPage + "?ACTN=" + encodeURIComponent(valueOf("hAction"));
		}
		return cfg.listPage;
	}

	function submitTo(action) {
		if (action) {
			form().action = action;
		}
		form().submit();
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return null;
		}
		window.open(url, "_blank", "width=700,height=500,resizable=yes,scrollbars=yes");
		return null;
	}

	function popupInfo(type, fallback) {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(type) : fallback;
		var parts = String(value || fallback).split(":");
		return {
			program: parts[0],
			features: "dialogHeight:" + (parts[1] || "500") + "px;dialogWidth:" + (parts[2] || "500") + "px;Status:No"
		};
	}

	function runStringDialog(page, query, features, done) {
		openDialog(page + "?" + query, "", features, function (outValue) {
			var text = trim(outValue);
			var parts = text.split(":");
			if (text !== "" && parts.length === 1) {
				runStringDialog(page, text, features, done);
				return;
			}
			done(text, parts);
		});
	}

	function trimLeadingComma(value) {
		var text = String(value || "");
		return text.charAt(0) === "," ? text.substring(1) : text;
	}

	function selectedVouchers(separator) {
		var checks = fields("Chkbox");
		var appNos = fields("hFrmAppNo");
		var values = [];
		var appNo = "";
		var lastText = "";
		checks.forEach(function (item, index) {
			if (item.checked) {
				values.push(item.value);
				lastText = item.getAttribute("text") || item.text || "";
				if (appNos[index]) {
					appNo = appNos[index].value || "";
				}
			}
		});
		return {
			values: values,
			joined: values.join(separator),
			appNo: appNo,
			lastText: lastText
		};
	}

	function requireSelected(selected) {
		if (trim(selected.joined) === "") {
			alert("Select Voucher");
			return false;
		}
		return true;
	}

	function requireSingle(selected, message) {
		if (!requireSelected(selected)) {
			return false;
		}
		if (selected.values.length > 1 || selected.joined.indexOf("~") !== -1) {
			alert(message || "Edit One by One");
			return false;
		}
		return true;
	}

	function checkedTransType() {
		var selected = [];
		fields("voutype").forEach(function (item) {
			if (item.checked) {
				selected.push(item.value);
			}
		});
		return selected.join(",");
	}

	function buildSearchOption(root, cfg) {
		var item = createElement(root, "SearchOption");
		item.setAttribute("SelUnitId", valueOf("hOrgID"));
		item.setAttribute("SelBook", valueOf("selBook"));
		item.setAttribute("SelUser", valueOf("selUser"));
		item.setAttribute(cfg.typeAttr, valueOf(cfg.typeField));
		item.setAttribute("SelVouFrom", valueOf("selVouFrm"));
		item.setAttribute("VoucherType", "");
		root.appendChild(item);
	}

	function appendRange(root, name, fromValue, toValue) {
		var item = createElement(root, name);
		item.setAttribute("From", fromValue || "");
		item.setAttribute("To", toValue || "");
		root.appendChild(item);
	}

	function saveSearchXml() {
		var cfg = config();
		var root = xmlRoot("SearchData");
		var acc;
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var accHead = trim(valueOf("txtAccHead")) !== "" ? valueOf("hAccHead") : "0";
		var subTypeText = byId("spParSubType") ? trim(byId("spParSubType").innerHTML || byId("spParSubType").textContent) : "";
		if (!root) {
			loadXml("SearchData", "<Root/>");
			root = xmlRoot("SearchData");
		}
		acc = createElement(root, "AccHead");
		clearChildren(root);
		root.setAttribute("Src", valueOf("hGridName"));
		root.setAttribute("TransType", checkedTransType());
		buildSearchOption(root, cfg);
		appendRange(root, "VoucherNo", field("ChkVouNo") && field("ChkVouNo").checked ? valueOf("txtVouNoFrom") : "", field("ChkVouNo") && field("ChkVouNo").checked ? valueOf("txtVouNoTo") : "");
		appendRange(root, "VoucherDate", field("ChkVouDt") && field("ChkVouDt").checked ? fromDate : "", field("ChkVouDt") && field("ChkVouDt").checked ? toDate : "");
		appendRange(root, "VoucherAmount", field("ChkVouAmt") && field("ChkVouAmt").checked ? valueOf("txtFromAmount") : "", field("ChkVouAmt") && field("ChkVouAmt").checked ? valueOf("txtToAmount") : "");
		acc.setAttribute("No", accHead);
		acc.setAttribute("Name", trim(valueOf("txtAccHead")));
		acc.setAttribute("ParSubTypeValue", subTypeText ? valueOf("hParSubTypeVal") : "");
		acc.setAttribute("ParSubTypeName", subTypeText ? valueOf("hParSubTypeName") : "");
		root.appendChild(acc);
		saveXmlIsland("SearchData", "XMLSave.asp?Name=SearchCriteria&Mod=ACC");
	}

	function setTypeSelection(value) {
		var items = fields("voutype");
		var selected = "," + trim(value) + ",";
		items.forEach(function (item) {
			item.checked = selected.indexOf("," + trim(item.value) + ",") !== -1;
		});
	}

	window.Sort = function (fieldNo, orderByField, order) {
		setValue("hField" + trim(fieldNo), trim(orderByField) + ":" + trim(order));
		setValue("hFieldSelected", fieldNo);
		submitTo(listAction());
		return false;
	};

	window.ChangeStatusOfInputFields = function (item) {
		var disabled = !(item && item.checked);
		if (item && item.name === "ChkVouNo") {
			if (field("txtVouNoFrom")) {
				field("txtVouNoFrom").disabled = disabled;
			}
			if (field("txtVouNoTo")) {
				field("txtVouNoTo").disabled = disabled;
			}
		}
		if (item && item.name === "ChkVouAmt") {
			if (field("txtFromAmount")) {
				field("txtFromAmount").disabled = disabled;
			}
			if (field("txtToAmount")) {
				field("txtToAmount").disabled = disabled;
			}
		}
		return true;
	};

	window.PaginateAcc = function (pageNo) {
		setValue("hPageSelection", pageNo);
		submitTo(listAction());
		return false;
	};

	function checkPeriodDate(controlName, label) {
		var bounds = finPeriodBounds();
		var yearMonth = dateYearMonth(controlName);
		if (yearMonth && (yearMonth < bounds.min || yearMonth > bounds.max)) {
			alert(label + " Date must be Between " + bounds.fromText + " and " + bounds.toText);
			setDateControl(controlName, todayDisplay());
			if (field(controlName) && field(controlName).focus) {
				field(controlName).focus();
			}
			return false;
		}
		return true;
	}

	window.CheckFromDate = function () {
		return checkPeriodDate("ctlVouFromDate", "From");
	};

	window.CheckToDate = function () {
		return checkPeriodDate("ctlVouToDate", "To");
	};

	window.DisplayBook = function () {
		var cfg = config();
		var bookSelect = field("selBook");
		var current = bookSelect ? bookSelect.value : "";
		var xhr;
		var root;
		if (!bookSelect) {
			return true;
		}
		bookSelect.options.length = 1;
		if (trim(valueOf("hOrgID")) === "") {
			return true;
		}
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=" + cfg.bookCode + "&orgID=" + encodeURIComponent(valueOf("hOrgID")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		root = xmlRoot("UnitBookData");
		childElements(root).forEach(function (node) {
			bookSelect.add(new Option(attr(node, 1), attr(node, 0)));
		});
		if (current) {
			bookSelect.value = current;
		}
		return true;
	};

	window.ChkVouType = function () {
		var root = xmlRoot("SearchData");
		if (!root) {
			loadXml("SearchData", "<Root/>");
			root = xmlRoot("SearchData");
		}
		root.setAttribute("Src", valueOf("hGridName"));
		root.setAttribute("TransType", checkedTransType());
		saveXmlIsland("SearchData", "XMLSave.asp?Name=SearchCriteria&Mod=ACC");
		submitTo(listAction());
		return false;
	};

	window.ChkforApprove = function () {
		var cfg = config();
		var selected = selectedVouchers(":");
		if (!requireSelected(selected)) {
			return false;
		}
		if (trim(selected.appNo) === cfg.otherAppNo) {
			setValue("hAppNo", selected.appNo);
			submitTo(cfg.otherAppPage + "?TransNo=" + encodeURIComponent(selected.joined) + "&sPara=App");
			return false;
		}
		setValue("hTransNo", selected.joined);
		if (confirm("Do you want to Approve")) {
			submitTo("AppVouStatusUpdateAll.asp");
		}
		return false;
	};

	window.ChkforDelete = function () {
		var cfg = config();
		var selected = selectedVouchers("|");
		if (!requireSelected(selected)) {
			return false;
		}
		if (trim(selected.appNo) === cfg.otherAppNo) {
			alert("Deletion Is Not Possible");
			return false;
		}
		if (!confirm("Do you want to delete the selected voucher")) {
			return false;
		}
		setValue("hTransNo", cfg.isPurchase ? selected.joined + "|0" : "0|" + selected.joined);
		if (confirm("This will Permanently Delete the Voucher(s)\nClick OK to Delete")) {
			submitTo(cfg.deletePage);
		}
		return false;
	};

	window.ChkforEdit = function () {
		var cfg = config();
		var selected = selectedVouchers("~");
		if (!requireSingle(selected, "Edit One by One")) {
			return false;
		}
		if (trim(selected.appNo) === cfg.otherAppNo) {
			setValue("hAppNo", selected.appNo);
			submitTo(cfg.otherAppPage + "?TransNo=" + encodeURIComponent(selected.joined) + "&sPara=Edt");
			return false;
		}
		setValue("hTransNo", selected.joined);
		if (cfg.isPurchase) {
			submitTo(cfg.editPage + "?ACTN=E:" + encodeURIComponent(selected.joined));
		} else {
			submitTo(cfg.editPage + "?nTransNo=" + encodeURIComponent(selected.joined));
		}
		return false;
	};

	function accountVoucher() {
		var cfg = config();
		var selected = selectedVouchers("~");
		var statusField;
		if (field("btnAcc")) {
			field("btnAcc").disabled = true;
		}
		if (!requireSingle(selected, "Account One by One")) {
			if (field("btnAcc")) {
				field("btnAcc").disabled = false;
			}
			return false;
		}
		if (trim(selected.appNo) === cfg.otherAppNo) {
			setValue("hAppNo", selected.appNo);
			submitTo(cfg.otherAppPage + "?TransNo=" + encodeURIComponent(selected.joined) + "&sPara=Acc");
			return false;
		}
		statusField = field("hVouSts" + selected.joined);
		if (statusField && statusField.value === "01") {
			alert("Only Approved Vouchers can be Accounted ");
			if (field("btnAcc")) {
				field("btnAcc").disabled = false;
			}
			return false;
		}
		if (statusField && statusField.value === "04") {
			alert("Voucher already Accounted ");
			if (field("btnAcc")) {
				field("btnAcc").disabled = false;
			}
			return false;
		}
		setValue("hTransNo", selected.joined);
		submitTo(cfg.accountPage);
		return false;
	}

	window.ChkforAcc = accountVoucher;
	window.ChkforACC = accountVoucher;

	window.SelectAccHead = function () {
		if (field("selBook") && field("selBook").value === "S") {
			alert("Select Book");
			if (field("selBook").focus) {
				field("selBook").focus();
			}
			return false;
		}
		return true;
	};

	window.Validate = function () {
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		if (field("ChkVouDt") && field("ChkVouDt").checked && from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}
		if (field("ChkVouNo") && field("ChkVouNo").checked && (trim(valueOf("txtVouNoFrom")) === "" || trim(valueOf("txtVouNoTo")) === "")) {
			alert("Voucher No Empty");
			return false;
		}
		if (field("ChkVouAmt") && field("ChkVouAmt").checked && (trim(valueOf("txtFromAmount")) === "" || trim(valueOf("txtToAmount")) === "")) {
			alert("Voucher Amount Empty");
			return false;
		}
		setValue("hFromDate", fromDate);
		setValue("hToDate", toDate);
		setValue("hXmlAccFlag", byId("spParSubType") && trim(byId("spParSubType").innerHTML || byId("spParSubType").textContent) !== "" ? "True" : "");
		setValue("hXmlParFlag", trim(valueOf("txtAccHead")) !== "" ? "True" : "");
		saveSearchXml();
		submitTo(listAction());
		return false;
	};

	window.ChkReset = function () {
		setDateControl("ctlVouFromDate", todayDisplay());
		setDateControl("ctlVouToDate", todayDisplay());
		if (field(config().typeField)) {
			field(config().typeField).selectedIndex = 0;
		}
		if (field("selBook")) {
			field("selBook").selectedIndex = 0;
		}
		["txtFromAmount", "txtToAmount", "txtVouNoFrom", "txtVouNoTo", "txtAccHead", "hParSubTypeVal", "hParSubTypeName", "hAccHead", "hParCode", "hPartyName"].forEach(function (name) {
			setValue(name, "");
		});
		if (byId("spParSubType")) {
			byId("spParSubType").innerHTML = "";
		}
		window.DisplayBook();
		return false;
	};

	window.ClearSearch = function () {
		var root = xmlRoot("SearchData");
		if (!root) {
			loadXml("SearchData", "<Root/>");
			root = xmlRoot("SearchData");
		}
		clearChildren(root);
		root.setAttribute("Src", valueOf("hGridName"));
		root.setAttribute("TransType", "");
		saveXmlIsland("SearchData", "XMLSave.asp?Name=SearchCriteria&Mod=ACC");
		submitTo(listAction());
		return false;
	};

	window.ShowVouch = function (transNo) {
		openDialog(config().viewPage + "?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	};

	function applySearchData() {
		var root = xmlRoot("SearchData");
		var vouFromData = "";
		if (!root || attr(root, "Src") !== valueOf("hGridName")) {
			window.DisplayBook();
			return;
		}
		setTypeSelection(attr(root, "TransType"));
		childElements(root).forEach(function (node) {
			if (node.nodeName === "SearchOption") {
				window.DisplayBook();
				setValue("selBook", attr(node, "SelBook"));
				setValue("selUser", attr(node, "SelUser"));
				setValue(config().typeField, attr(node, config().typeAttr));
				vouFromData = attr(node, "SelVouFrom");
			}
			if (node.nodeName === "VoucherNo" && trim(attr(node, "From")) !== "" && trim(attr(node, "To")) !== "") {
				setValue("txtVouNoFrom", attr(node, "From"));
				setValue("txtVouNoTo", attr(node, "To"));
				if (field("ChkVouNo")) {
					field("ChkVouNo").checked = true;
				}
			}
			if (node.nodeName === "VoucherDate" && trim(attr(node, "From")) !== "" && trim(attr(node, "To")) !== "") {
				setValue("hFromDate", attr(node, "From"));
				setValue("hToDate", attr(node, "To"));
				setDateControl("ctlVouFromDate", attr(node, "From"));
				setDateControl("ctlVouToDate", attr(node, "To"));
				if (field("ChkVouDt")) {
					field("ChkVouDt").checked = true;
				}
			}
			if (node.nodeName === "VoucherAmount" && trim(attr(node, "From")) !== "" && trim(attr(node, "To")) !== "") {
				setValue("txtFromAmount", attr(node, "From"));
				setValue("txtToAmount", attr(node, "To"));
				if (field("ChkVouAmt")) {
					field("ChkVouAmt").checked = true;
				}
			}
			if (node.nodeName === "AccHead" && trim(attr(node, "No")) !== "") {
				setValue("hAccHead", attr(node, "No"));
				setValue("txtAccHead", attr(node, "Name"));
				setValue("hParSubTypeVal", attr(node, "ParSubTypeValue"));
				setValue("hParSubTypeName", attr(node, "ParSubTypeName"));
				if (byId("spParSubType")) {
					byId("spParSubType").innerHTML = attr(node, "ParSubTypeName");
				}
			}
		});
		setValue("selVouFrm", vouFromData || "P");
		window.ChangeStatusOfInputFields(field("ChkVouNo"));
		window.ChangeStatusOfInputFields(field("ChkVouAmt"));
	}

	window.SetDate = function () {
		var fromDate = valueOf("hFromDate");
		var toDate = valueOf("hToDate");
		ensureCompat();
		if (trim(fromDate) !== "" && trim(toDate) !== "") {
			setDateControl("ctlVouFromDate", fromDate);
			setDateControl("ctlVouToDate", toDate);
		}
		applySearchData();
		if (!valueOf("selVouFrm")) {
			setValue("selVouFrm", "P");
		}
		return true;
	};

	window.ResetAccHead = function () {
		if (byId("spParSubType")) {
			byId("spParSubType").innerHTML = "";
		}
		["hParSubTypeVal", "hParSubTypeName", "hAccHead", "txtAccHead"].forEach(function (name) {
			setValue(name, "");
		});
		return false;
	};

	function applySubTypeSelection() {
		var root = xmlRoot("AccHeadData");
		var names = [];
		var values = [];
		childElements(root).forEach(function (node) {
			if (node.nodeName === "Details") {
				names.push(attr(node, "PartySubTypeName"));
				values.push("'" + attr(node, "PartyType") + attr(node, "PartySubType") + "'");
			}
		});
		setValue("hParSubTypeVal", values.join(","));
		setValue("hParSubTypeName", names.join(","));
		if (byId("spParSubType")) {
			byId("spParSubType").innerHTML = names.join(",");
		}
		saveXmlIsland("AccHeadData", "XMLSave.asp?Name=PartySubType&Mod=" + config().module);
	}

	window.SelAccHeadPopup = function () {
		var root = xmlRoot("AccHeadData");
		openDialog("PartySubTypeMultipleSel.asp?Unit=" + encodeURIComponent(valueOf("hOrgID")), xmlObject("AccHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;status:no", function () {
			root = xmlRoot("AccHeadData");
			if (!root || !childElements(root).length) {
				setValue("hParSubTypeVal", "");
				setValue("hParSubTypeName", "");
				if (byId("spParSubType")) {
					byId("spParSubType").innerHTML = "";
				}
				return;
			}
			applySubTypeSelection();
		});
		return false;
	};

	function partyTypeForPopup() {
		var root = xmlRoot("AccHeadData");
		var types = [];
		var subTypes = [];
		var names = [];
		childElements(root).forEach(function (node) {
			if (node.nodeName === "Details") {
				types.push(attr(node, "PartyType"));
				subTypes.push(attr(node, "PartySubType"));
				names.push(attr(node, "PartySubTypeName"));
			}
		});
		if (!types.length) {
			return "0";
		}
		return types.join(",") + "?" + subTypes.join(",") + "?" + names.join(",");
	}

	window.SelPartyPopup = function () {
		var info = popupInfo("12", "PartySelectionAcc.asp:500:500");
		setValue("hPartyName", "");
		setValue("txtAccHead", "");
		runStringDialog("../../Common/" + info.program, "orgID=" + encodeURIComponent(valueOf("hOrgID")) + "&Party=" + encodeURIComponent(partyTypeForPopup()), info.features, function (outValue, parts) {
			var partyName;
			var partyCode;
			if (parts.length <= 1 || outValue === "") {
				return;
			}
			partyName = trimLeadingComma(parts[0] || "");
			partyCode = trimLeadingComma(parts[1] || "");
			saveXmlIsland("PartyData", "XMLSave.asp?Name=PartyType&Mod=" + config().module);
			setValue("hPartyName", partyName);
			setValue("hParCode", partyCode);
			setValue("hAccHead", partyCode);
			setValue("txtAccHead", partyName);
		});
		return false;
	};

	function wireDateBlur() {
		var from = field("ctlVouFromDate") || byId("ctlVouFromDate");
		var to = field("ctlVouToDate") || byId("ctlVouToDate");
		ensureCompat();
		if (from) {
			from.addEventListener("blur", window.CheckFromDate);
		}
		if (to) {
			to.addEventListener("blur", window.CheckToDate);
		}
	}

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", wireDateBlur);
	} else {
		wireDateBlur();
	}
}(window, document));
