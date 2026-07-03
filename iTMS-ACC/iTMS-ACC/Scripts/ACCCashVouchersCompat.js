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
		return frm.elements[name] || frm.elements[String(name).toLowerCase()] || frm.elements[String(name).toUpperCase()] || frm.elements[String(name).charAt(0).toUpperCase() + String(name).slice(1)] || (function () {
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
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
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

	function syncPost(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(null);
		return xhr;
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumeric(value) {
		return trim(value) !== "" && !isNaN(Number(String(value).replace(/,/g, "")));
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
		var parts = String(period).split(":");
		var start = parts[0] || String(period).substring(0, 4);
		var end = parts[1] || String(period).slice(-4);
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
		if (window.showModalDialog && callback) {
			callback(window.showModalDialog(url, args || "", features || ""));
			return;
		}
		window.open(url, "_blank", "height=480,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function runStringDialog(page, query, features, done) {
		openDialog(page + "?" + query, "", features, function (outValue) {
			var text = trim(outValue);
			var parts;
			if (text === "") {
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

	function clearAccHeadXml() {
		clearChildren(xmlRoot("AccHeadData"));
	}

	function lastAccHeadNode() {
		var nodes = childElements(xmlRoot("AccHeadData"));
		return nodes.length ? nodes[nodes.length - 1] : null;
	}

	function resetAccHead() {
		var select = field("selAccHead");
		if (select) {
			select.selectedIndex = 0;
		}
		setValue("hAccHead", "0");
		setValue("txtAccHead", "");
	}

	function applyGlSelection() {
		var node = lastAccHeadNode();
		if (node) {
			setValue("hAccHead", attr(node, "No"));
			setValue("txtAccHead", attr(node, "Name"));
		} else {
			resetAccHead();
		}
		setValue("hAccIndex", field("selAccHead") ? field("selAccHead").selectedIndex : "");
		setValue("hAccTxt", valueOf("txtAccHead"));
	}

	function applyPartySelection(partyPrefix) {
		var node = lastAccHeadNode();
		if (node) {
			setValue("hAccHead", partyPrefix + "?" + attr(node, "No"));
			setValue("txtAccHead", attr(node, "Name"));
		} else {
			resetAccHead();
		}
		setValue("hAccIndex", field("selAccHead") ? field("selAccHead").selectedIndex : "");
		setValue("hAccTxt", valueOf("txtAccHead"));
	}

	function currentCriteria() {
		var radios = fields("OptCriteria");
		var checked = radios.filter(function (radio) {
			return radio.checked;
		})[0];
		return checked ? checked.value : valueOf("hFlag");
	}

	function setReadOnly(names, readOnly) {
		names.forEach(function (name) {
			var item = field(name);
			if (item) {
				item.readOnly = !!readOnly;
			}
		});
	}

	window.CheckFromDate = function () {
		return checkPeriodDate("ctlVouFromDate", "From");
	};

	window.CheckToDate = function () {
		return checkPeriodDate("ctlVouToDate", "To");
	};

	window.DisplayBook = function () {
		var unit = field("selUnitId");
		var partySelect = field("selAccHead");
		var bookSelect = field("selBook");
		var unitNo = unit ? unit.value : "";
		var xhr;
		var root;

		if (partySelect) {
			partySelect.options.length = Math.min(partySelect.options.length, 2);
		}
		if (bookSelect) {
			bookSelect.options.length = 1;
		}
		setValue("hUnitNo", unitNo);
		if (!unit || unit.selectedIndex === 0) {
			return true;
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

		xhr = syncGet("XMLGetOrgBook.asp?BkCode=01&orgID=" + encodeURIComponent(unitNo));
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

	window.ChkforEdit = function () {
		var count = valueOf("hCnt", valueOf("hcnt", "0"));
		var checks = fields("Chkbox");
		var checked = checks.filter(function (item) {
			return item.checked;
		});
		var selected;
		var textParts;
		var voucherParts;
		var actionValue;

		if (String(count) !== "1" && checked.length > 1) {
			alert("Edit One by One");
			return false;
		}
		selected = checked[0] || (String(count) === "1" ? checks[0] : null);
		if (!selected) {
			return false;
		}
		textParts = String(selected.getAttribute("text") || selected.text || "").split("&");
		voucherParts = String(textParts[1] || "").split("@");
		setValue("hTransNo", selected.value);
		actionValue = selected.value + "~" + (textParts[0] || "") + "~A";
		form().action = "AMDCASHVOUCHER.ASP?Val=" + encodeURIComponent(actionValue) + "&VouTy=" + encodeURIComponent(voucherParts[0] || "");
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.OptSelection = function () {
		var criteria = currentCriteria();
		var accHead = field("selAccHead");
		sFlag = criteria;
		if (criteria === "VouNo") {
			setReadOnly(["txtVouNoFrom", "txtVouNoTo"], false);
			setReadOnly(["txtFromAmount", "txtToAmount"], true);
			setValue("txtFromAmount", "");
			setValue("txtToAmount", "");
			if (accHead) {
				accHead.disabled = true;
			}
		} else if (criteria === "VouDate") {
			setReadOnly(["txtVouNoFrom", "txtVouNoTo", "txtFromAmount", "txtToAmount"], true);
			["txtVouNoFrom", "txtVouNoTo", "txtFromAmount", "txtToAmount"].forEach(function (name) {
				setValue(name, "");
			});
			if (accHead) {
				accHead.disabled = true;
			}
		} else if (criteria === "VouAmount") {
			setReadOnly(["txtVouNoFrom", "txtVouNoTo"], true);
			setReadOnly(["txtFromAmount", "txtToAmount"], false);
			setValue("txtVouNoFrom", "");
			setValue("txtVouNoTo", "");
			if (accHead) {
				accHead.disabled = true;
			}
		} else if (criteria === "AccHead") {
			setReadOnly(["txtVouNoFrom", "txtVouNoTo", "txtFromAmount", "txtToAmount"], true);
			["txtVouNoFrom", "txtVouNoTo", "txtFromAmount", "txtToAmount"].forEach(function (name) {
				setValue(name, "");
			});
			if (accHead) {
				accHead.disabled = false;
			}
		}
		setValue("hFlag", criteria);
		return true;
	};
	window.Optselection = window.OptSelection;

	window.SelectAccHead = function () {
		var unit = field("selUnitId");
		var book = field("selBook");
		var select = field("selAccHead");
		var orgId;
		var bookNo;
		var partyPrefix;
		var features = "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No";

		if (!select || select.selectedIndex === 0) {
			return false;
		}
		if (!unit || unit.selectedIndex === 0) {
			alert("Select Organaisation Id");
			resetAccHead();
			if (unit) {
				unit.focus();
			}
			return false;
		}
		if (!book || book.value === "S" || book.selectedIndex === 0) {
			alert("Select Book");
			if (book) {
				book.focus();
			}
			resetAccHead();
			return false;
		}

		orgId = unit.value;
		bookNo = book.value;
		clearAccHeadXml();
		if (select.value === "G") {
			runStringDialog("GLHeadSelection.asp", "orgid=" + encodeURIComponent(orgId) + "&BookId=&BookNo=" + encodeURIComponent(bookNo), features, function (outValue, parts) {
				if (parts.length <= 2 || parts[0] === "-1") {
					return;
				}
				if (typeof window.GetGlHeadXml === "function") {
					window.GetGlHeadXml(outValue);
				}
				applyGlSelection();
			});
		} else {
			partyPrefix = select.value + "?" + selectedText(select);
			runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyPrefix), "dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue, parts) {
				var partyName;
				var partyCode;
				var partySubType;
				var partyType;
				var response;
				if (parts.length <= 4 || parts[0] === "-1") {
					resetAccHead();
					return;
				}
				partyName = parts[0] || "";
				partyCode = parts[1] || "";
				partySubType = parts[3] || "";
				partyType = parts[4] || "";
				response = syncGet("XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(orgId) + "&ParSubType=" + encodeURIComponent(partySubType) + "&ParType=" + encodeURIComponent(partyType) + "&PartyCode=" + encodeURIComponent(partyCode));
				if (trim(response.responseText) !== "" && typeof window.GetPartyHeadXml === "function") {
					window.GetPartyHeadXml(partyCode, partyName, response.responseText);
				}
				applyPartySelection(partyPrefix);
			});
		}
		return false;
	};

	window.Validate = function () {
		var criteria = currentCriteria();
		var fromDate = getDateControl("ctlVouFromDate");
		var toDate = getDateControl("ctlVouToDate");
		var from = parseDate(fromDate);
		var to = parseDate(toDate);
		if (criteria === "VouDate" && from && to && from > to) {
			alert("To Date Should be Greater than From Date");
			return false;
		}
		if (criteria === "VouNo") {
			if (trim(valueOf("txtVouNoFrom")) === "" || trim(valueOf("txtVouNoTo")) === "") {
				alert("Voucher No Empty");
				return false;
			}
			setValue("hVouFrom", valueOf("txtVouNoFrom"));
			setValue("hVouTo", valueOf("txtVouNoTo"));
		} else if (criteria === "VouAmount") {
			if (trim(valueOf("txtFromAmount")) === "" || trim(valueOf("txtToAmount")) === "") {
				alert("Voucher Amount Empty");
				return false;
			}
			if (!isNumeric(valueOf("txtFromAmount")) || !isNumeric(valueOf("txtToAmount"))) {
				alert("Enter Numbers Only");
				return false;
			}
			if (toNumber(valueOf("txtFromAmount")) > toNumber(valueOf("txtToAmount"))) {
				alert("To Amount Should be Greater Than From Amount ");
				return false;
			}
			setValue("hAmtFrom", valueOf("txtFromAmount"));
			setValue("hAmtTo", valueOf("txtToAmount"));
		}
		setValue("hFromDate", fromDate);
		setValue("hToDate", toDate);
		window.GetFormDet();
		form().submit();
		return false;
	};

	window.ChkReset = function () {
		var unit = field("selUnitId");
		var book = field("selBook");
		var acc = field("selAccHead");
		setDateControl("ctlVouFromDate", todayDisplay());
		setDateControl("ctlVouToDate", todayDisplay());
		fields("OptCriteria").forEach(function (radio) {
			radio.checked = false;
		});
		if (unit) {
			unit.selectedIndex = 0;
		}
		if (acc) {
			acc.selectedIndex = 0;
			acc.disabled = true;
		}
		if (book) {
			book.selectedIndex = 0;
		}
		["txtFromAmount", "txtToAmount", "txtVouNoFrom", "txtVouNoTo", "txtAccHead", "hFlag"].forEach(function (name) {
			setValue(name, "");
		});
		sFlag = "";
		window.DisplayBook();
		return false;
	};

	window.ShowVouch = function (transNo) {
		openDialog("CashVouchView_San.asp?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No", function () {});
		return false;
	};

	window.SetDate = function () {
		var criteria = valueOf("hFlag");
		var radios = fields("OptCriteria");
		var fromDate = valueOf("hFromDate");
		var toDate = valueOf("hToDate");
		if (criteria === "VouAmount") {
			setValue("txtFromAmount", valueOf("hAmtFrom"));
			setValue("txtToAmount", valueOf("hAmtTo"));
			if (radios[2]) {
				radios[2].checked = true;
			}
		} else if (criteria === "VouNo") {
			setValue("txtVouNoFrom", valueOf("hVouFrom"));
			setValue("txtVouNoTo", valueOf("hVouTo"));
			if (radios[0]) {
				radios[0].checked = true;
			}
		} else if (criteria === "VouDate") {
			if (radios[1]) {
				radios[1].checked = true;
			}
		} else if (criteria === "AccHead") {
			if (radios[3]) {
				radios[3].checked = true;
			}
		}
		if (trim(fromDate) !== "" && trim(toDate) !== "") {
			setDateControl("ctlVouFromDate", fromDate);
			setDateControl("ctlVouToDate", toDate);
		}
		window.DisplayBook();
		window.OptSelection();
		if (field("selBook") && field("selBook").options.length > 1) {
			field("selBook").selectedIndex = Number(valueOf("hBookNo", 0)) || 0;
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = Number(valueOf("hAccIndex", 0)) || 0;
			}
			setValue("txtAccHead", valueOf("hAccTxt"));
		} else if (field("selBook")) {
			field("selBook").selectedIndex = 0;
		}
		sFlag = currentCriteria();
		return true;
	};

	window.GetFormDet = function () {
		var unit = field("selUnitId");
		var book = field("selBook");
		var criteria = fields("OptCriteria").filter(function (radio) {
			return radio.checked;
		})[0];
		var value = unit ? unit.value : "";
		value += "|" + (unit ? unit.selectedIndex : "");
		value += "|" + (book ? book.value : "");
		value += "|" + (book ? book.selectedIndex : "");
		value += "|" + getDateControl("ctlVouFromDate");
		value += "|" + getDateControl("ctlVouToDate");
		value += "|" + valueOf("txtFromAmount");
		value += "|" + valueOf("txtToAmount");
		value += "|" + valueOf("txtVouNoFrom");
		value += "|" + valueOf("txtVouNoTo");
		value += "|" + valueOf("hAccIndex");
		value += "|" + (criteria ? criteria.value : "0");
		value += "|" + valueOf("hAccHead");
		value += "|" + valueOf("txtAccHead");
		setValue("hFormVal", value);
		return value;
	};

	window.CheckVouForEdit = function () {
		var checks = fields("Chkbox");
		var checked = checks.filter(function (item) {
			return item.checked;
		});
		var selected = checked[0] || checks[0];
		var xhr;
		if (!selected) {
			return true;
		}
		setValue("hTransNo", selected.value);
		xhr = syncPost("CheckVouAdjDet.asp?BookCode=01&TransNo=" + encodeURIComponent(selected.value));
		if (trim(xhr.responseText) !== "") {
			window.DispVal(trim(xhr.responseText));
			return false;
		}
		return true;
	};

	window.DispVal = function (value) {
		alert("Voucher is adjusted with the following Bills! \n Amendment Not allowed \n " + value);
	};

	function wireDateBlur() {
		var from = field("ctlVouFromDate") || byId("ctlVouFromDate");
		var to = field("ctlVouToDate") || byId("ctlVouToDate");
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
