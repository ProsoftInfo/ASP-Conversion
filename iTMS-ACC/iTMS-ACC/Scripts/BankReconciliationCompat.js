(function (window, document) {
	"use strict";

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

	function textOf(id) {
		var item = byId(id);
		return item ? trim(item.textContent != null ? item.textContent : item.innerText) : "";
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

	function setChecked(name, index, checked) {
		var group = form() && form().elements[name];
		if (!group) {
			return;
		}
		if (group.length) {
			group[index].checked = checked;
		} else if (index === 0) {
			group.checked = checked;
		}
	}

	function checkedValue(name) {
		var group = form() && form().elements[name];
		var index;
		if (!group) {
			return "";
		}
		if (!group.length) {
			return group.checked ? group.value : "";
		}
		for (index = 0; index < group.length; index += 1) {
			if (group[index].checked) {
				return group[index].value;
			}
		}
		return "";
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, "").replace(/[A-Za-z]/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function abs(value) {
		return Math.abs(toNumber(value));
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
	}

	function formatSigned(value, positiveSuffix, negativeSuffix, negativeAbs) {
		var number = toNumber(value);
		if (number < 0) {
			return formatNumber(negativeAbs ? Math.abs(number) : number) + " " + negativeSuffix;
		}
		return formatNumber(number) + " " + positiveSuffix;
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var parts;
		if (value instanceof Date && !isNaN(value.getTime())) {
			return new Date(value.getFullYear(), value.getMonth(), value.getDate());
		}
		if (!text) {
			return null;
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
		if (!control || !value) {
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
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

	function syncPost(url, payload) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		if (payload != null) {
			try {
				xhr.setRequestHeader("Content-Type", "text/xml");
			} catch (ignore) {}
		}
		xhr.send(payload == null ? null : payload);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "height=500,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function continueSelection(url, args, features, callback) {
		openDialog(url, args, features, function (value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var action = String(attr(root, "Action")).toUpperCase();
			var query = trim(attr(root, "PassQuery"));
			var base = String(url).replace(/\?.*$/, "");
			if (action && action !== "DONE" && action !== "CLOSE" && query) {
				continueSelection(base + "?" + query, args, features, callback);
				return;
			}
			callback(root || value);
		});
	}

	function splitBookValue(value) {
		var parts = String(value || "").split("~");
		return {
			accHead: trim(parts[0]),
			bookNo: trim(parts[1])
		};
	}

	function selectBookNoFromControl(control) {
		var source = control;
		if (source && source.value !== undefined) {
			return splitBookValue(source.value).bookNo;
		}
		return trim(source);
	}

	function getOrgId() {
		return valueOf("horgId") || valueOf("hOrgId") || valueOf("selUnitId");
	}

	function isMonthModePage() {
		return !!field("SelToMonth");
	}

	function vd(value) {
		var date = parseDate(value);
		var text = trim(value);
		var parts = text.split(/[\/.-]/);
		if (!date || parts.length < 3) {
			return false;
		}
		return date.getFullYear() >= 1900 && Number(parts[0]) === date.getDate() && Number(parts[1]) === date.getMonth() + 1;
	}

	function checkValidDate(value, compareDate, flag) {
		var date = parseDate(value);
		var compare = parseDate(compareDate);
		if (!date || !compare) {
			return false;
		}
		if (Number(flag) === 0 && date > compare) {
			alert("Date entered should be less than or equal to today's date");
			return false;
		}
		if (Number(flag) === 1 && date < compare) {
			alert("Date entered should be greater than or equal to " + compareDate);
			return false;
		}
		return true;
	}

	function DisplayBook() {
		var select = field("selBook");
		var orgField = field("selUnitId");
		var orgId = orgField ? orgField.value : getOrgId();
		var xhr;
		var root;
		if (!select || !orgId) {
			return false;
		}
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=02&orgID=" + encodeURIComponent(orgId));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", new XMLSerializer().serializeToString(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		root = xmlRoot("UnitBookData");
		if (!root) {
			return false;
		}
		select.options.length = 0;
		childElements(root).forEach(function (node) {
			var bookName = attr(node, "BookName") || attr(node, 1);
			var accHead = attr(node, "AccHead") || attr(node, "BookAccountHead") || attr(node, 0);
			var bookNo = attr(node, "BookNumber") || attr(node, 0);
			select.add(new Option(bookName, accHead + "~" + bookNo));
			setValue("hBookAccHead", accHead);
			setValue("hBookNo", bookNo);
		});
		return true;
	}

	function FnInit() {
		var frm = form();
		var outRoot;
		var book = field("selBook");
		var bookParts = splitBookValue(book && book.value);
		var orgField = field("selUnitId");
		var passBalance = trim(valueOf("txtPassBalance"));
		var response;
		if (!frm) {
			return false;
		}
		if (field("SelToMonth")) {
			setValue("hMonthYr", field("SelToMonth").value);
		}
		if (byId("ctlFrmDate")) {
			setValue("hSelFromDt", getDateControl("ctlFrmDate"));
		}
		if (byId("ctlToDate")) {
			setValue("hSelToDt", getDateControl("ctlToDate"));
		}
		if (!passBalance) {
			alert("Enter Passbook Balance");
			if (field("txtPassBalance")) {
				field("txtPassBalance").select();
			}
			return false;
		}
		if (isNaN(Number(passBalance.replace(/,/g, "")))) {
			alert("Enter numeric value in Passbook Balance");
			if (field("txtPassBalance")) {
				field("txtPassBalance").select();
			}
			return false;
		}
		if (toNumber(passBalance) < 0 || toNumber(passBalance) > 9999999999.99) {
			alert("Passbook Balance should be between 0 and 9999999999.99");
			if (field("txtPassBalance")) {
				field("txtPassBalance").select();
			}
			return false;
		}
		if (!bookParts.bookNo) {
			alert("Select Bank");
			if (book) {
				book.focus();
			}
			return false;
		}
		outRoot = xmlRoot("OutData");
		if (!outRoot) {
			alert("Bank reconciliation data is not available.");
			return false;
		}
		setValue("hOrgName", orgField ? selectedText(orgField) : valueOf("hOrgName"));
		setAttr(outRoot, "orgId", orgField ? orgField.value : getOrgId());
		setAttr(outRoot, "orgName", orgField ? selectedText(orgField) : valueOf("hOrgName"));
		setAttr(outRoot, "BookNo", bookParts.bookNo);
		setAttr(outRoot, "AccHead", bookParts.accHead);
		setAttr(outRoot, "BookName", selectedText(book));
		setAttr(outRoot, "FormDt", valueOf("hSelFromDt"));
		setAttr(outRoot, "ToDt", valueOf("hSelToDt"));
		setAttr(outRoot, "PassBal", passBalance);
		setAttr(outRoot, "PassBalCRDR", String(checkedValue("optPassBalCrDR")).toUpperCase() === "DR" ? "DR" : "CR");
		setValue("hBookNo", bookParts.bookNo);
		setValue("hBookName", selectedText(book));
		setValue("hBookAccHead", bookParts.accHead);
		response = syncPost("XMLSave.asp?Name=Bank Recon&Mod=BA", serializeXml("OutData"));
		if (trim(response.responseText)) {
			alert(response.responseText);
		} else {
			frm.submit();
		}
		return false;
	}

	function monthYearFromDateText(value) {
		var date = parseDate(value);
		if (!date) {
			return trim(value);
		}
		return String(date.getFullYear()) + pad2(date.getMonth() + 1);
	}

	function GetPassBookBal(typeOrDate, source) {
		var bookNo;
		var monthValue;
		var requestValue;
		var responseText;
		if (String(typeOrDate) === "S") {
			bookNo = selectBookNoFromControl(source);
		} else if (String(typeOrDate) === "M") {
			bookNo = trim(source);
		} else if (isMonthModePage()) {
			bookNo = trim(source);
		} else {
			bookNo = selectBookNoFromControl(source);
		}

		if (isMonthModePage()) {
			if (String(typeOrDate) === "M") {
				monthValue = valueOf("SelToMonth");
			} else {
				monthValue = monthYearFromDateText(valueOf("hToDt"));
			}
		} else {
			monthValue = trim(typeOrDate) || getDateControl("ctlFrmDate");
		}
		if (!bookNo) {
			return false;
		}
		requestValue = monthValue + ":" + bookNo + ":" + getOrgId();
		setValue("hBookNo", bookNo);
		responseText = trim(syncPost("XMLGETPASSBkBAL.asp?MonYr=" + encodeURIComponent(requestValue)).responseText);
		if (responseText && !isNaN(Number(responseText))) {
			if (toNumber(responseText) < 0) {
				setChecked("optPassBalCrDR", 1, true);
				setChecked("optPassBalCrDR", 0, false);
			} else {
				setChecked("optPassBalCrDR", 0, true);
				setChecked("optPassBalCrDR", 1, false);
			}
			setValue("txtPassBalance", formatNumber(abs(responseText)));
		}
		return true;
	}

	function clearedOnField(slNo, transNo) {
		return field("txtClearedOn" + slNo + "Z" + transNo);
	}

	function voucherNodes() {
		return childElements(xmlRoot("VoucherData"));
	}

	function updateVoucherNode(slNo, flag, clearedOn) {
		voucherNodes().forEach(function (node) {
			if (attr(node, "SlNo") === String(slNo)) {
				setAttr(node, "Flag", flag);
				setAttr(node, "ClearedOn", clearedOn || "");
			}
		});
	}

	function updateDifferenceSpans(transType, amount, checked) {
		var passAmount;
		var diffAmount;
		var newDiff;
		if (trim(valueOf("hChkVal")) === "BR") {
			return;
		}
		passAmount = toNumber(textOf("dPassBook"));
		diffAmount = toNumber(textOf("dDiffAmt"));
		if (checked) {
			diffAmount += String(transType) === "R" ? toNumber(amount) : -toNumber(amount);
		} else {
			diffAmount += String(transType) === "R" ? -toNumber(amount) : toNumber(amount);
		}
		newDiff = diffAmount - passAmount;
		setText("dDiffAmt", formatSigned(diffAmount, "CR", "DR", false));
		setText("dNewDiff", formatSigned(newDiff, "CR", "DR", true));
	}

	function setClearOn(checkbox, transNo, clearDate, transType, amount, slNo) {
		var clearedField = clearedOnField(slNo, transNo);
		updateVoucherNode(slNo, checkbox && checkbox.checked ? "Y" : "N", checkbox && checkbox.checked ? clearDate : "");
		if (trim(valueOf("hChkVal")) !== "BR" && clearedField) {
			clearedField.readOnly = !checkbox.checked;
			if (!checkbox.checked) {
				clearedField.value = "";
			} else {
				clearedField.value = "";
				clearedField.focus();
			}
			updateDifferenceSpans(transType, amount, !!checkbox.checked);
		}
		return true;
	}

	function anyTransactionChecked() {
		var max = toNumber(valueOf("hSno")) - 1;
		var index;
		var checkbox;
		for (index = 1; index <= max; index += 1) {
			checkbox = field("chkTransNo" + index);
			if (checkbox && checkbox.checked) {
				return true;
			}
		}
		return false;
	}

	function validateClearedNode(node) {
		var slNo = attr(node, "SlNo");
		var transNo = attr(node, "TransNo");
		var control = clearedOnField(slNo, transNo);
		var value = trim(control && control.value);
		if (!value) {
			alert("Enter Cleared On Date");
			if (control) {
				control.focus();
			}
			return false;
		}
		if (!vd(value)) {
			alert("Enter Valid Date - dd/mm/yyyy format");
			if (control) {
				control.focus();
			}
			return false;
		}
		if (!checkValidDate(value, todayDisplay(), 0)) {
			if (control) {
				control.focus();
			}
			return false;
		}
		if (!checkValidDate(value, valueOf("hFromDt"), 1)) {
			if (control) {
				control.focus();
			}
			return false;
		}
		setAttr(node, "ClearedOn", value);
		return true;
	}

	function disableActionButtons() {
		var buttons = form() && form().elements.B6;
		var index;
		if (!buttons) {
			return;
		}
		if (!buttons.length) {
			buttons.disabled = true;
			return;
		}
		for (index = 0; index < buttons.length; index += 1) {
			buttons[index].disabled = true;
		}
	}

	function finalDone() {
		var nodes;
		var response;
		var frm = form();
		if (!frm) {
			return false;
		}
		if (!anyTransactionChecked()) {
			alert("Select Voucher No");
			return false;
		}
		nodes = voucherNodes();
		for (var index = 0; index < nodes.length; index += 1) {
			if (attr(nodes[index], "Flag") === "Y" && !validateClearedNode(nodes[index])) {
				return false;
			}
		}
		disableActionButtons();
		response = syncPost("XMLUpdate.asp?Name=Bank Recon&Mod=BA", serializeXml("VoucherData"));
		if (trim(response.responseText)) {
			alert(response.responseText);
		} else {
			frm.action = "BrsDisplay.asp";
			frm.submit();
		}
		return false;
	}

	function showCharges() {
		var frm = form();
		if (!frm) {
			return false;
		}
		frm.action = "BrsCommEntry.asp";
		frm.submit();
		return false;
	}

	function SetDate() {
		var book = field("selBook");
		if (valueOf("hSelFromDt")) {
			setDateControl("ctlFrmDate", valueOf("hSelFromDt"));
		}
		if (valueOf("hSelToDt")) {
			setDateControl("ctlToDate", valueOf("hSelToDt"));
		}
		if (book) {
			book.value = valueOf("hBookAccHead") + "~" + valueOf("hBookNo");
		}
		if (byId("PayToId")) {
			setText("PayToId", valueOf("hPartyName"));
		}
		if (form() && form().elements.BRType) {
			if (trim(valueOf("hChkVal")) === "BNR") {
				setChecked("BRType", 0, true);
				setChecked("BRType", 1, false);
			} else {
				setChecked("BRType", 1, true);
				setChecked("BRType", 0, false);
			}
		}
		return true;
	}

	function MinDate() {
		var fromDate = getDateControl("ctlFrmDate");
		var toDate = getDateControl("ctlToDate");
		var minText = valueOf("hMinDt");
		var maxText = valueOf("hMaxDt");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		var min = parseDate(minText);
		var max = parseDate(maxText);
		if (from && min && max && (from < min || from > max)) {
			alert("Date Should be within the Financial Year  " + minText + " to " + maxText);
			setDateControl("ctlFrmDate", minText);
			return false;
		}
		if (to && min && max && (to < min || to > max)) {
			alert("Date Should be within the Financial Year  " + minText + " to " + maxText);
			setDateControl("ctlToDate", maxText);
			return false;
		}
		return true;
	}

	function Recon(source) {
		var value = trim(source && source.value);
		setChecked("BRType", 0, value === "BNR");
		setChecked("BRType", 1, value === "BR");
		setValue("hChkVal", value);
		if (field("B6")) {
			field("B6").disabled = false;
		}
		setValue("hPartyCode", "");
		setValue("hPartyName", "");
		setText("PayToId", "");
		form().submit();
		return false;
	}

	function popupSize() {
		var size = window.GetWindowSizeForPopup ? window.GetWindowSizeForPopup("2") : "PartySelection.asp:500:420";
		var parts = String(size).split(":");
		return {
			program: parts[0] || "PartySelection.asp",
			height: parts[1] || "500",
			width: parts[2] || "420"
		};
	}

	function finishPartySelection(root) {
		var entries = childElements(root);
		var index;
		var entry;
		for (index = 0; index < entries.length; index += 1) {
			entry = entries[index];
			if (String(entry.nodeName).toLowerCase() === "entry") {
				setText("PayToId", attr(entry, "RetField0"));
				setValue("hPartyName", attr(entry, "RetField0"));
				setValue("hPartyCode", attr(entry, "RetField1"));
				setValue("hParSubType", attr(entry, "RetField4"));
				break;
			}
		}
	}

	function PartySelect() {
		var orgId = getOrgId();
		var size = popupSize();
		var url = "../../Common/" + size.program + "?orgid=" + encodeURIComponent(orgId);
		var features = "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No";
		continueSelection(url, xmlObject("PartyData") || xmlObject("OutData"), features, finishPartySelection);
		return false;
	}

	function ShowData() {
		var frm = form();
		if (!frm) {
			return false;
		}
		frm.action = "PendingForClearance.asp";
		frm.submit();
		return false;
	}

	window.vd = vd;
	window.checkValidDate = checkValidDate;
	window.DisplayBook = DisplayBook;
	window.FnInit = FnInit;
	window.GetPassBookBal = GetPassBookBal;
	window.setClearOn = setClearOn;
	window.finalDone = finalDone;
	window.showCharges = showCharges;
	window.SetDate = SetDate;
	window.MinDate = MinDate;
	window.Recon = Recon;
	window.PartySelect = PartySelect;
	window.ShowData = ShowData;
}(window, document));
