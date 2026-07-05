(function (window, document) {
	"use strict";

	var sBookCode = "08";
	var sFlag = "VouNo";
	var sVouType = "B";
	var sHead = "";
	var sHeadCode = "";
	var sHeadDesc = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var elements;
		var target;
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		elements = frm.elements;
		target = String(name).toLowerCase();
		for (index = 0; index < elements.length; index += 1) {
			if (String(elements[index].name || elements[index].id || "").toLowerCase() === target) {
				return elements[index];
			}
		}
		return null;
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function valueOf(name, fallback) {
		var item = field(name) || byId(name);
		return item ? item.value : fallback || "";
	}

	function setValue(name, value) {
		var item = field(name) || byId(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedValue(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : "";
	}

	function focusField(name) {
		var item = field(name) || byId(name);
		if (item && typeof item.focus === "function") {
			item.focus();
		}
		if (item && typeof item.select === "function") {
			item.select();
		}
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var parts;
		if (!text) {
			return null;
		}
		if (value instanceof Date && !isNaN(value.getTime())) {
			return new Date(value.getFullYear(), value.getMonth(), value.getDate());
		}
		if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
			parts = text.split("-");
			return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
		}
		parts = text.split(/[\/.-]/);
		if (parts.length >= 3) {
			return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
		}
		return null;
	}

	function toIsoDate(value) {
		var date = parseDate(value);
		if (!date) {
			return "";
		}
		return date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate());
	}

	function toDisplayDate(value) {
		var date = parseDate(value);
		if (!date) {
			return "";
		}
		return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear();
	}

	function todayDisplay() {
		return toDisplayDate(new Date());
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
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
		return toDisplayDate(control.value || "");
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = toIsoDate(value) || value;
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function loadXml(name, text) {
		var object = xmlObject(name);
		var doc;
		if (!object || !trim(text)) {
			return false;
		}
		if (typeof object.loadXML === "function") {
			return object.loadXML(text);
		}
		doc = xmlDocument(object);
		if (doc && typeof doc.loadXML === "function") {
			return doc.loadXML(text);
		}
		return false;
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

	function selectNodes(context, expression) {
		var doc;
		var result;
		var nodes = [];
		var index;
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		try {
			result = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
			for (index = 0; index < result.snapshotLength; index += 1) {
				nodes.push(result.snapshotItem(index));
			}
		} catch (ignore) {}
		return nodes;
	}

	function childElements(node) {
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1;
		});
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "height=600,width=800,resizable=no,status=no,scrollbars=yes");
	}

	function continueSelection(url, args, features, callback) {
		openDialog(url, args, features, function (value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var action = String(attr(root, "Action")).toUpperCase();
			var query = trim(attr(root, "PassQuery"));
			if (action && action !== "DONE" && action !== "CLOSE" && query) {
				continueSelection("../Reports/GLHeadSelectionMultiple.asp?" + query, args, features, callback);
				return;
			}
			callback(root || value);
		});
	}

	function checkNumbers(value) {
		return /^[0-9]+$/.test(trim(value));
	}

	function GetVouType() {
		var choices = form() ? form().elements.optVoutype : null;
		var index;
		if (!choices) {
			return sVouType;
		}
		if (!choices.length) {
			sVouType = choices.checked ? choices.value : sVouType;
			return sVouType;
		}
		for (index = 0; index < choices.length; index += 1) {
			if (choices[index].checked) {
				sVouType = choices[index].value;
				break;
			}
		}
		return sVouType;
	}

	function finishAccHeadSelection(returnRoot) {
		var root = xmlRoot("AccHeadData");
		var entries = selectNodes(root, "//Entry");
		var heads = [];
		var headCodes = [];
		var headDescs = [];
		if (String(attr(returnRoot, "Action")).toUpperCase() === "CLOSE") {
			loadXml("AccHeadData", "<account/>");
			setText("spAccHead", "");
			if (field("SelAccHead")) {
				field("SelAccHead").selectedIndex = 0;
			}
			return;
		}
		entries.forEach(function (entry) {
			heads.push(attr(entry, "RetField1"));
			headCodes.push(attr(entry, "RetField2"));
			headDescs.push(attr(entry, "RetField0"));
		});
		if (entries.length) {
			sHead = heads.join(",");
			sHeadCode = headCodes.join(",");
			sHeadDesc = headDescs.join(",");
			setText("spAccHead", sHeadDesc);
			if (field("SelAccHead")) {
				field("SelAccHead").selectedIndex = 1;
			}
		} else if (field("SelAccHead")) {
			field("SelAccHead").selectedIndex = 0;
		}
	}

	function SelectAccHead() {
		var book = field("selBook");
		var orgId = valueOf("hUnitId");
		var bookNo;
		var url;
		sHead = "";
		sHeadCode = "";
		sHeadDesc = "";
		if (!book || book.value === "S") {
			alert("Select Book");
			if (book) {
				book.focus();
			}
			return false;
		}
		bookNo = selectedValue(book);
		url = "../Reports/GLHeadSelectionMultiple.asp?orgid=" + encodeURIComponent(orgId) + "&BookId=" + encodeURIComponent(sBookCode) + "&BookNo=" + encodeURIComponent(bookNo) + "&hSelectMode=M";
		continueSelection(url, xmlObject("AccHeadData"), "dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", finishAccHeadSelection);
		return false;
	}

	function SelNew() {
		if (field("SelAccHead")) {
			field("SelAccHead").selectedIndex = 0;
		}
		setValue("txtGAmount", "");
		setValue("txtLAmount", "");
		setValue("txtNoFrom", "");
		setValue("txtNoTo", "");
		setText("spAccHead", "");
		sHead = "";
		sHeadCode = "";
		sHeadDesc = "";
		return true;
	}

	function openReport(sCallTy, passString) {
		var url;
		if (String(sCallTy) === "S") {
			url = "DBGJView.asp?" + passString;
			openDialog(url, "A", "dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No", function () {});
		} else {
			url = "../Reports/PrnDJView.asp?" + passString;
			openDialog(url, "A", "dialogHeight:150px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No", function () {});
		}
	}

	function CheckSubmit(sCallTy) {
		var book = field("selBook");
		var orgId;
		var orgName;
		var bookName;
		var bookNo;
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		var voucherFrom = "";
		var voucherTo = "";
		var greaterAmount = "";
		var lesserAmount = "";
		var flags = [];
		var hasFilter = false;
		var pass;
		GetVouType();
		if (!book || book.selectedIndex === 0) {
			alert("Select Book ");
			if (book) {
				book.focus();
			}
			return false;
		}
		orgId = valueOf("hUnitId");
		orgName = valueOf("hUnitName");
		bookName = selectedText(book);
		bookNo = selectedValue(book);

		if (field("chkBox1") && field("chkBox1").checked) {
			hasFilter = true;
			flags.push(field("chkBox1").value);
			voucherFrom = trim(valueOf("txtNoFrom"));
			voucherTo = trim(valueOf("txtNoTo"));
			if (!voucherFrom) {
				alert("Enter Voucher No. From ");
				focusField("txtNoFrom");
				return false;
			}
			if (!voucherTo) {
				alert("Enter Voucher No. To ");
				focusField("txtNoTo");
				return false;
			}
			if (voucherFrom > voucherTo) {
				alert("Voucher Number Should be Greater Than Previous");
				focusField("txtNoTo");
				return false;
			}
		}

		if (from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}

		if (field("chkBox2") && field("chkBox2").checked) {
			hasFilter = true;
			flags.push(field("chkBox2").value);
			greaterAmount = trim(valueOf("txtGAmount"));
			lesserAmount = trim(valueOf("txtLAmount"));
			if (!greaterAmount) {
				alert("Enter From Amount");
				focusField("txtGAmount");
				return false;
			}
			if (!lesserAmount) {
				alert("Enter To Amount");
				focusField("txtLAmount");
				return false;
			}
			if (!checkNumbers(greaterAmount)) {
				alert("Enter Numbers Only");
				focusField("txtGAmount");
				return false;
			}
			if (!checkNumbers(lesserAmount)) {
				alert("Enter Numbers Only");
				focusField("txtLAmount");
				return false;
			}
			if (toNumber(greaterAmount) > toNumber(lesserAmount)) {
				alert("To Amount Should be Greater From Amount");
				setValue("txtLAmount", "");
				focusField("txtLAmount");
				return false;
			}
		}

		if (field("SelAccHead") && field("SelAccHead").value === "S") {
			hasFilter = true;
		}

		if (!hasFilter) {
			alert("Select any one option to View Data");
			return false;
		}

		pass = "Value=" + encodeURIComponent([
			orgId,
			orgName,
			bookName,
			bookNo,
			flags.join(","),
			voucherFrom,
			voucherTo,
			sVouType,
			fromDate,
			toDate,
			greaterAmount,
			lesserAmount,
			sHead,
			sHeadCode,
			sHeadDesc
		].join("|"));
		openReport(sCallTy, pass);
		return false;
	}

	function DisplayBook() {
		var select = field("selBook");
		var xhr;
		var root;
		if (!select) {
			return false;
		}
		select.options.length = 1;
		SelNew();
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=08&orgID=" + encodeURIComponent(valueOf("hUnitId")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", new XMLSerializer().serializeToString(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		root = xmlRoot("UnitBookData");
		childElements(root).forEach(function (node) {
			select.add(new Option(attr(node, 1), attr(node, 0)));
		});
		return true;
	}

	function MinDate() {
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var minText = valueOf("hFDate");
		var maxText = valueOf("hTDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		var min = parseDate(minText);
		var max = parseDate(maxText);
		if (from && min && max && (from < min || from > max)) {
			alert("Date Should be within the Financial Year  " + minText + " to " + maxText);
			setDateControl("ctlVouFromDate", minText);
			return false;
		}
		if (to && min && max && (to < min || to > max)) {
			alert("Date Should be within the Financial Year  " + minText + " to " + maxText);
			setDateControl("ctlVouToDate", maxText);
			return false;
		}
		return true;
	}

	function OptSelection() {
		var byVoucher = field("chkBox1") && field("chkBox1").checked;
		var byAmount = field("chkBox2") && field("chkBox2").checked;
		if (byVoucher) {
			sFlag = field("chkBox1").value;
			if (field("txtNoFrom")) {
				field("txtNoFrom").readOnly = false;
			}
			if (field("txtNoTo")) {
				field("txtNoTo").readOnly = false;
			}
			setValue("txtGAmount", "");
			setValue("txtLAmount", "");
		} else {
			if (field("txtNoFrom")) {
				field("txtNoFrom").readOnly = true;
			}
			if (field("txtNoTo")) {
				field("txtNoTo").readOnly = true;
			}
			setValue("txtNoFrom", "");
			setValue("txtNoTo", "");
		}
		if (byAmount) {
			sFlag = field("chkBox2").value;
			setValue("txtGAmount", "");
			setValue("txtLAmount", "");
			if (field("txtGAmount")) {
				field("txtGAmount").readOnly = false;
			}
			if (field("txtLAmount")) {
				field("txtLAmount").readOnly = false;
			}
		} else {
			setValue("txtGAmount", "");
			setValue("txtLAmount", "");
			if (field("txtGAmount")) {
				field("txtGAmount").readOnly = true;
			}
			if (field("txtLAmount")) {
				field("txtLAmount").readOnly = true;
			}
		}
		return true;
	}

	function setdate() {
		var fromDate = valueOf("hFDate");
		var toDate = valueOf("hTDate");
		var today = parseDate(todayDisplay());
		var max = parseDate(toDate);
		setDateControl("ctlVouFromDate", fromDate);
		setDateControl("ctlVouToDate", today && max && today <= max ? todayDisplay() : toDate);
		return true;
	}

	window.GetVouType = GetVouType;
	window.SelectAccHead = SelectAccHead;
	window.checkNumbers = checkNumbers;
	window.SelNew = SelNew;
	window.CheckSubmit = CheckSubmit;
	window.DisplayBook = DisplayBook;
	window.MinDate = MinDate;
	window.OptSelection = OptSelection;
	window.setdate = setdate;
}(window, document));
