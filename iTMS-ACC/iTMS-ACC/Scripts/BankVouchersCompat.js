(function (window, document) {
	"use strict";

	var sFlag = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var elements;
		var target;
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		return frm.elements[name] || frm.elements[String(name).toLowerCase()] || frm.elements[String(name).toUpperCase()] || (function () {
			elements = frm.elements;
			target = String(name).toLowerCase();
			for (index = 0; index < elements.length; index += 1) {
				if (String(elements[index].name || "").toLowerCase() === target) {
					return elements[index];
				}
			}
			return null;
		}());
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
		if (!item) {
			return fallback || "";
		}
		return item.value == null ? fallback || "" : item.value;
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
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
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

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
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
			legacyMin: start + "03",
			legacyMax: end + "04"
		};
	}

	function checkPeriodDate(name, label) {
		var bounds = finPeriodBounds();
		var yearMonth = dateYearMonth(name);
		var control = field(name) || byId(name);
		if (yearMonth && (yearMonth < bounds.legacyMin || yearMonth > bounds.legacyMax)) {
			alert(label + " Date must be Between " + bounds.fromText + " and " + bounds.toText);
			setDateControl(name, todayDisplay());
			if (control && control.focus) {
				control.focus();
			}
			return false;
		}
		return true;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		window.open(url, "_blank", "height=480,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function runStringDialog(page, query, features, done) {
		openDialog(page + "?" + query, "", features, function (outValue) {
			var text = trim(outValue);
			var parts;
			if (text === "") {
				done("", []);
				return;
			}
			parts = text.split(":");
			if (parts.length === 1) {
				runStringDialog(page, text, features, done);
				return;
			}
			done(text, parts);
		});
	}

	function popupInfo(type, fallback) {
		var raw = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(type) : fallback;
		var parts = String(raw || fallback).split(":");
		return {
			program: parts[0],
			height: parts[1] || "500",
			width: parts[2] || "500",
			features: "dialogHeight:" + (parts[1] || "500") + "px;dialogWidth:" + (parts[2] || "500") + "px;Status:No"
		};
	}

	function dialogNode(value) {
		return xmlRoot(value) || value && value.nodeType === 1 && value || null;
	}

	function runGlDialog(url, args, features, program, depth, done) {
		if (depth > 12) {
			done(null);
			return;
		}
		openDialog(url, args, features, function (outValue) {
			var node = dialogNode(outValue);
			var action = trim(attr(node, "Action")).toUpperCase();
			var query = trim(attr(node, "PassQuery"));
			if (action && action !== "CLOSE" && action !== "DONE" && query) {
				runGlDialog("../../Common/" + program + "?" + query, args, features, program, depth + 1, done);
				return;
			}
			done(node);
		});
	}

	function selectEntries(root) {
		var nodes = [];
		if (!root) {
			return nodes;
		}
		if (typeof root.selectNodes === "function") {
			nodes = root.selectNodes("//account/Entry");
			return Array.prototype.slice.call(nodes);
		}
		nodes = root.getElementsByTagName ? root.getElementsByTagName("Entry") : [];
		return Array.prototype.slice.call(nodes);
	}

	function resetAccHead() {
		var select = field("selAccHead");
		if (select) {
			select.value = "0";
			select.selectedIndex = 0;
		}
		setValue("hAccHead", "");
		setValue("txtAccHead", "");
	}

	function rememberAccHead() {
		var select = field("selAccHead");
		setValue("hAccIndex", select ? select.selectedIndex : "");
		setValue("hAccTxt", valueOf("txtAccHead"));
	}

	function applyGlSelection(outNode) {
		var entries = selectEntries(xmlRoot("AccHeadData"));
		var codes = [];
		var names = [];
		entries.forEach(function (entry) {
			if (trim(attr(entry, "RetField0")) !== "") {
				codes.push(attr(entry, "RetField0"));
			}
			if (trim(attr(entry, "RetField5")) !== "") {
				names.push(attr(entry, "RetField5"));
			}
		});
		if (entries.length || outNode && childElements(outNode).length) {
			setValue("hAccHead", codes.join(","));
			setValue("txtAccHead", names.join(","));
		} else {
			setValue("hAccHead", "0");
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = 0;
			}
		}
		rememberAccHead();
	}

	function voucherCheckboxes() {
		return fields("Chkbox");
	}

	function voucherText(item) {
		return item ? item.getAttribute("text") || item.text || "" : "";
	}

	function rowIndex(item) {
		var checks = voucherCheckboxes();
		var index = checks.indexOf(item);
		return index === -1 ? 1 : index + 1;
	}

	function rowInfo(item) {
		var index = rowIndex(item);
		return {
			frmAppNo: valueOf("hFrmAppNo" + index),
			purBillType: valueOf("hPurBillType" + index),
			text: voucherText(item)
		};
	}

	function collectVouchers(separator, allowSingleUnchecked) {
		var checks = voucherCheckboxes();
		var count = trim(valueOf("hCnt", valueOf("hcnt", String(checks.length))));
		var chosen = checks.filter(function (item) {
			return item.checked;
		});
		var result = {
			items: [],
			values: [],
			joined: "",
			lastInfo: { frmAppNo: "", purBillType: "", text: "" }
		};
		if (allowSingleUnchecked && count === "1" && chosen.length === 0 && checks[0]) {
			chosen = [checks[0]];
		}
		chosen.forEach(function (item) {
			result.items.push(item);
			result.values.push(item.value);
			result.lastInfo = rowInfo(item);
		});
		result.joined = result.values.join(separator);
		return result;
	}

	function setBtnAcc(disabled) {
		var button = field("btnAcc");
		if (button) {
			button.disabled = !!disabled;
		}
	}

	function checkedValue(name) {
		var checked = fields(name).filter(function (item) {
			return item.checked;
		})[0];
		return checked ? checked.value : "";
	}

	window.ResetAccHead = resetAccHead;

	window.CheckFromDate = function () {
		return checkPeriodDate("ctlVouFromDate", "From");
	};

	window.CheckToDate = function () {
		return checkPeriodDate("ctlVouToDate", "To");
	};

	window.Sort = function (fieldNo, orderByField, order) {
		setValue("hField" + trim(fieldNo), trim(orderByField) + ":" + trim(order));
		setValue("hFieldSelected", fieldNo);
		form().submit();
		return false;
	};

	window.DisplayBook = function () {
		var unitNo = valueOf("hUnitNo");
		var partySelect = field("selAccHead");
		var bookSelect = field("selBook");
		var xhr;
		var root;

		if (partySelect) {
			partySelect.options.length = Math.min(partySelect.options.length, 2);
		}
		if (bookSelect) {
			bookSelect.options.length = 1;
		}

		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(unitNo));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("OutData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("OutData", xhr.responseText);
		}
		root = xmlRoot("OutData");
		childElements(root).forEach(function (node) {
			if (partySelect) {
				partySelect.add(new Option(trim(node.textContent || node.text || ""), attr(node, "ParType")));
			}
		});

		xhr = syncGet("XMLGetOrgBook.asp?BkCode=02&orgID=" + encodeURIComponent(unitNo));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("UnitBookData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("UnitBookData", xhr.responseText);
		}
		root = xmlRoot("UnitBookData");
		childElements(root).forEach(function (node) {
			if (bookSelect) {
				bookSelect.add(new Option(attr(node, 1), attr(node, 0)));
			}
		});
		return true;
	};

	window.GetBookNo = function () {
		var book = field("selBook");
		setValue("hBookNo", book ? book.selectedIndex : "");
		setValue("hBookVal", book ? book.value : "");
		return true;
	};

	window.ChkVouType = function () {
		var vouchers = fields("voutype");
		vouchers.slice(1).forEach(function (item) {
			if (item.checked && vouchers[0]) {
				vouchers[0].checked = false;
			}
		});
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.ChkforApprove = function () {
		var selected = collectVouchers(":", true);
		var textParts = selected.lastInfo.text.split("@");
		if (selected.items.length === 0) {
			return false;
		}
		if (trim(valueOf("hCnt")) !== "1" && trim(textParts[3]) === "") {
			alert("Approve is not Possible.Select A/c. Head or Party.");
			return false;
		}
		window.GetFormDet();
		if (trim(selected.joined) !== "" && confirm("Do you want to Approve")) {
			setValue("hTransNo", selected.joined);
			form().action = "AppVouStatusUpdateAll.asp";
			form().submit();
		}
		return false;
	};

	window.ChkforDelete = function () {
		var selected = collectVouchers("|", true);
		if (trim(selected.joined) === "") {
			alert("Select Voucher");
			return false;
		}
		setValue("hTransNo", "0|" + selected.joined);
		if (confirm("This will Permanently Delete the Voucher(s)\nClick OK to Delete")) {
			form().action = "VouDeletionAll.asp";
			window.GetFormDet();
			form().submit();
		}
		return false;
	};

	window.ChkforEdit = function () {
		var selected = collectVouchers("~", true);
		var textParts;
		var voucherParts;
		var actionValue;
		if (selected.values.length > 1) {
			alert("Edit One by One");
			return false;
		}
		if (trim(selected.joined) === "") {
			return false;
		}
		textParts = selected.lastInfo.text.split("&");
		voucherParts = String(textParts[1] || "").split("@");
		setValue("hTransNo", selected.joined);
		actionValue = selected.joined + "~" + (textParts[0] || "") + "~A";
		form().action = "BankVoucher.asp?Val=" + encodeURIComponent(actionValue) + "&VouTy=" + encodeURIComponent(voucherParts[0] || "");
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.ChkforAccount = function () {
		var selected;
		var statusParts;
		var voucherParts;
		setBtnAcc(true);
		selected = collectVouchers("~", true);
		statusParts = selected.lastInfo.text.split("@");
		voucherParts = String(statusParts[0] || "").split("&");
		if (trim(valueOf("hCnt")) !== "1" && selected.items.length > 0 && String(statusParts[1] || "") === "01") {
			alert("Only Approved Vouchers can be Accounted");
			setBtnAcc(false);
			return false;
		}
		if (selected.items.length > 0 && voucherParts[1] === "R" && Number(statusParts[2] || 0) > 0) {
			alert("Only Bank Reconciliation Vouchers can be Accounted");
			setBtnAcc(false);
			return false;
		}
		if (selected.values.length > 1) {
			alert("Account One by One");
			setBtnAcc(false);
			return false;
		}
		window.GetFormDet();
		if (trim(selected.joined) !== "") {
			setValue("hTransNo", selected.joined);
			form().action = "AccVouGenerate.asp";
			window.GetFormDet();
			setBtnAcc(true);
			form().submit();
		} else {
			setBtnAcc(false);
		}
		return false;
	};

	window.ShowVouch = function (transNo) {
		openDialog("BankVouchView_San.asp?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:600px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	};

	window.OptSelection = function () {
		var vouNo = field("ChkVouNo");
		var vouDate = field("ChkVouDt");
		var vouAmount = field("ChkVouAmt");
		var cheque = field("ChkChq");
		var accHead = field("selAccHead");
		setValue("hVouNoFlag", vouNo && vouNo.checked ? vouNo.value : "");
		setValue("hVouDtFlag", vouDate && vouDate.checked ? vouDate.value : "");
		setValue("hVouAmtFlag", vouAmount && vouAmount.checked ? vouAmount.value : "");
		setValue("hChqFlag", cheque && cheque.checked ? cheque.value : "");
		sFlag = accHead && trim(accHead.value) !== "0" ? "AccHead" : "";
		setValue("hFlag", sFlag);
		return true;
	};
	window.Optselection = window.OptSelection;

	window.SelectAccHead = function () {
		var book = field("selBook");
		var select = field("selAccHead");
		var orgId = valueOf("hUnitNo");
		var bookNo;
		var info;
		var partyType;

		if (!book || book.value === "S") {
			alert("Select Book");
			if (book && book.focus) {
				book.focus();
			}
			if (select) {
				select.selectedIndex = 0;
			}
			return false;
		}
		if (!select || trim(select.value) === "0") {
			return false;
		}
		bookNo = book.value;
		setValue("hAccHead", "");
		setValue("txtAccHead", "");
		if (select.value === "G") {
			info = popupInfo("5", "GLHeadSelection.asp:500:350");
			runGlDialog("../../Common/" + info.program + "?orgID=" + encodeURIComponent(orgId) + "&BookId=&BookNo=" + encodeURIComponent(bookNo) + "&hSelectMode=M", xmlObject("AccHeadData"), info.features, info.program, 0, applyGlSelection);
		} else {
			info = popupInfo("12", "PartySelectionAcc.asp:500:500");
			partyType = select.value + "?" + selectedText(select);
			runStringDialog("../../Common/" + info.program, "orgID=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType), info.features, function (outValue, parts) {
				var partyName;
				var partyCode;
				var partySubType;
				var partyTypeValue;
				if (parts.length <= 2) {
					if (select) {
						select.selectedIndex = 0;
					}
					if (select && select.focus) {
						select.focus();
					}
					return;
				}
				if (outValue !== "") {
					partyName = parts[0] || "";
					partyCode = parts[1] || "";
					partySubType = parts[3] || "";
					partyTypeValue = parts[4] || "";
					setValue("hAccHead", partyTypeValue + "?" + partySubType + "?" + partyName + "?" + partyCode);
					setValue("txtAccHead", partyName);
				} else {
					if (select) {
						select.selectedIndex = 0;
					}
					setValue("hAccHead", "0");
					setValue("txtAccHead", "");
				}
				rememberAccHead();
			});
		}
		return false;
	};

	window.Validate = function () {
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		var book = field("selBook");
		var vouNo = field("ChkVouNo");
		var vouAmount = field("ChkVouAmt");
		var cheque = field("ChkChq");
		if (!book || book.value === "S") {
			alert("Select Book");
			if (book && book.focus) {
				book.focus();
			}
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = 0;
			}
			return false;
		}
		if (from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}
		window.OptSelection();
		if (vouNo && vouNo.checked) {
			if (trim(valueOf("txtVouNoFrom")) === "" || trim(valueOf("txtVouNoTo")) === "") {
				alert("Voucher No Empty");
				return false;
			}
			setValue("hVouNoFlag", "VouNo");
			setValue("hVouFrom", valueOf("txtVouNoFrom"));
			setValue("hVouTo", valueOf("txtVouNoTo"));
		}
		if (vouAmount && vouAmount.checked) {
			if (trim(valueOf("txtFromAmount")) === "" || trim(valueOf("txtToAmount")) === "") {
				alert("Voucher Amount Empty");
				return false;
			}
			setValue("hVouAmtFlag", "VouAmount");
			setValue("hAmtFrom", valueOf("txtFromAmount"));
			setValue("hAmtTo", valueOf("txtToAmount"));
		}
		if (cheque && cheque.checked) {
			if (trim(valueOf("txtFromChqNo")) === "" || trim(valueOf("txtToChqNo")) === "") {
				alert("Enter Cheque No From To ");
				return false;
			}
			if (!/^\d+$/.test(trim(valueOf("txtFromChqNo"))) || !/^\d+$/.test(trim(valueOf("txtToChqNo")))) {
				alert("Enter Only Numeric Value in Cheque No");
				return false;
			}
			setValue("hChqFlag", "Cheque");
			setValue("hChqFrom", valueOf("txtFromChqNo"));
			setValue("hChqTo", valueOf("txtToChqNo"));
		}
		setValue("hFromDate", fromDate);
		setValue("hToDate", toDate);
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.ChkReset = function () {
		var acc = field("selAccHead");
		var book = field("selBook");
		setDateControl("ctlVouFromDate", todayDisplay());
		setDateControl("ctlVouToDate", todayDisplay());
		if (acc) {
			acc.selectedIndex = 0;
			acc.disabled = true;
		}
		if (book) {
			book.selectedIndex = 0;
		}
		["txtFromAmount", "txtToAmount", "txtVouNoFrom", "txtVouNoTo", "txtAccHead", "txtFromChqNo", "txtToChqNo", "hFlag", "hChqFlag", "hChqFrom", "hChqTo"].forEach(function (name) {
			setValue(name, "");
		});
		sFlag = "";
		window.DisplayBook();
		return false;
	};

	window.SetDate = function () {
		var fromDate;
		var toDate;
		var book;
		var acc;
		ensureCompat();
		sFlag = valueOf("hFlag");
		if (valueOf("hVouNoFlag") === "VouNo") {
			setValue("txtVouNoFrom", valueOf("hVouFrom"));
			setValue("txtVouNoTo", valueOf("hVouTo"));
			if (field("ChkVouNo")) {
				field("ChkVouNo").checked = true;
			}
		}
		if (valueOf("hVouAmtFlag") === "VouAmount") {
			setValue("txtFromAmount", valueOf("hAmtFrom"));
			setValue("txtToAmount", valueOf("hAmtTo"));
			if (field("ChkVouAmt")) {
				field("ChkVouAmt").checked = true;
			}
		}
		if (valueOf("hVouDtFlag") === "VouDate" && field("ChkVouDt")) {
			field("ChkVouDt").checked = true;
		}
		if (valueOf("hChqFlag") === "Cheque") {
			if (field("ChkChq")) {
				field("ChkChq").checked = true;
			}
			setValue("txtFromChqNo", valueOf("hChqFrom"));
			setValue("txtToChqNo", valueOf("hChqTo"));
		}
		fromDate = valueOf("hFromDate");
		toDate = valueOf("hToDate");
		if (trim(fromDate) !== "" && trim(toDate) !== "") {
			setDateControl("ctlVouFromDate", fromDate);
			setDateControl("ctlVouToDate", toDate);
		}
		window.DisplayBook();
		book = field("selBook");
		acc = field("selAccHead");
		if (book && book.options.length > 1) {
			book.selectedIndex = Number(valueOf("hBookNo", 0)) || 0;
			if (acc) {
				acc.selectedIndex = Number(valueOf("hAccIndex", 0)) || 0;
			}
			setValue("txtAccHead", valueOf("hAccTxt"));
		} else if (book) {
			book.selectedIndex = 0;
		}
		return true;
	};

	window.GetFormDet = function () {
		var book = field("selBook");
		var criteria = "0";
		var formValue = valueOf("hUnitNo");
		var vouType = checkedValue("OptVouTy") || valueOf("hVocType");
		if (field("ChkVouNo") && field("ChkVouNo").checked) {
			criteria = field("ChkVouNo").value;
		} else if (field("ChkVouDt") && field("ChkVouDt").checked) {
			criteria = field("ChkVouDt").value;
		} else if (field("ChkVouAmt") && field("ChkVouAmt").checked) {
			criteria = field("ChkVouAmt").value;
		} else if (field("ChkChq") && field("ChkChq").checked) {
			criteria = field("ChkChq").value;
		}
		formValue += "|" + valueOf("hUnitNo");
		formValue += "|" + (book ? book.value : "");
		formValue += "|" + (book ? book.selectedIndex : "");
		formValue += "|" + getDateControl("ctlVouFromDate");
		formValue += "|" + getDateControl("ctlVouToDate");
		formValue += "|" + valueOf("txtFromAmount");
		formValue += "|" + valueOf("txtToAmount");
		formValue += "|" + valueOf("txtVouNoFrom");
		formValue += "|" + valueOf("txtVouNoTo");
		formValue += "|" + valueOf("hAccIndex");
		formValue += "|" + criteria;
		formValue += "|" + valueOf("hAccHead");
		formValue += "|" + valueOf("txtAccHead");
		formValue += "|" + valueOf("txtFromChqNo");
		formValue += "|" + valueOf("txtToChqNo");
		setValue("hFormVal", formValue);
		return formValue;
	};

	window.GetUser = function () {
		setValue("hUserID", valueOf("selUser"));
		return true;
	};

	window.ChkforPrint = function () {
		var selected = collectVouchers(":", true);
		if (trim(selected.joined) !== "") {
			openDialog("BankVouchView_San.asp?TransNo=" + encodeURIComponent(selected.joined), "", "dialogHeight:600px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {});
		} else {
			alert("select any one voucher for printing");
		}
		return false;
	};

	window.RegVoucherNo = function () {
		var book = field("selBook");
		var types = fields("OptVouTy");
		if (!book || book.value === "S") {
			alert("Select a book");
			return false;
		}
		if ((!types[1] || !types[1].checked) && (!types[2] || !types[2].checked)) {
			alert("Select Voucher Type - Receipts / Payments");
			return false;
		}
		if (types[1] && types[1].checked) {
			setValue("hVocType", "C");
		} else if (types[2] && types[2].checked) {
			setValue("hVocType", "D");
		}
		setValue("hFromDate", getDateControl("ctlVouFromDate"));
		form().action = "AccVoucherNo_Generate.asp?BookCode=02";
		form().submit();
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
