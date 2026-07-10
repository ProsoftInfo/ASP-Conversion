(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
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

	function valueOf(name) {
		var item = field(name);
		return item && item.value != null ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function xmlObject(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function createNode(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		if (object && object.XMLDocument) {
			return object.XMLDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
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

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function requestXml(url) {
		var xhr = new XMLHttpRequest();
		var text;
		xhr.open("GET", url, false);
		xhr.send();
		text = xhr.responseText || "";
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			return xhr.responseXML;
		}
		return text ? new DOMParser().parseFromString(text, "text/xml") : null;
	}

	function clearOptions(select) {
		if (select) {
			select.options.length = 0;
		}
	}

	function addOption(select, text, value) {
		if (select) {
			select.options[select.options.length] = new Option(text || "", value || "");
		}
	}

	function fillOptionsFromXml(selectName, xmlDoc) {
		var select = field(selectName);
		var root = xmlDoc && xmlDoc.documentElement;
		childElements(root).forEach(function (node) {
			addOption(select, attr(node, 2), attr(node, 0));
		});
	}

	function voucherNodes() {
		return childElements(xmlRoot("DetData"), "voucher");
	}

	function syncPartyCode() {
		var party = xmlDocument("DetData") ? xmlDocument("DetData").getElementsByTagName("Party")[0] : null;
		if (party) {
			setValue("hPartyCode", [attr(party, "ParType"), attr(party, "ParSubType"), "Fill", attr(party, "ParCode")].join("?"));
		}
	}

	function syncOldInvoiceCodes() {
		var codes = voucherNodes().map(function (voucher) {
			return attr(voucher, "SalTransNo") || attr(voucher, 6);
		}).filter(Boolean);
		setValue("hOldInvCode", codes.join(","));
	}

	function setDateLimits() {
		var control = field("ctlDate") || document.getElementById("ctlDate");
		var minDate = "01/04/" + trim(valueOf("hFromYr"));
		var maxDate = "31/03/" + trim(valueOf("hToYr"));
		if (!control) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(minDate);
		}
		if (typeof control.SetMaxDate === "function") {
			control.SetMaxDate(maxDate);
		}
	}

	function setMoveButtonsDisabled(disabled) {
		var add = field("add");
		var remove = field("remove");
		if (add) {
			add.disabled = !!disabled;
		}
		if (remove) {
			remove.disabled = !!disabled;
		}
	}

	window.PopulateInvoices = function () {
		var orgId = trim(valueOf("hOrgId"));
		var partyCode = valueOf("hPartyCode");
		var selected = voucherNodes().map(function (voucher) {
			return attr(voucher, "SalTransNo") || attr(voucher, 6);
		}).filter(Boolean).join(",");
		clearOptions(field("selFrombox"));
		clearOptions(field("selTobox"));
		fillOptionsFromXml("selFrombox", requestXml("XMLCommisionDetails.asp?OrgId=" + encodeURIComponent(orgId) + "&AgentCode=" + encodeURIComponent(partyCode)));
		if (selected) {
			fillOptionsFromXml("selTobox", requestXml("XMLCommisionDetails.asp?OrgId=" + encodeURIComponent(orgId) + "&AgentCode=" + encodeURIComponent(partyCode) + "&sSelInv=" + encodeURIComponent(selected)));
		}
		setMoveButtonsDisabled(true);
	};

	window.UpdateXml = function () {
		var root = xmlRoot("DetData");
		var selected = field("selTobox");
		var entries = childElements(root, "Entry");
		var oldVoucher = voucherNodes()[0];
		var party = oldVoucher && childElements(oldVoucher, "Party")[0];
		var insertBefore = entries[0] || null;
		var index;
		var option;
		var parts;
		var voucher;
		if (!root || !oldVoucher || !selected) {
			return;
		}
		voucherNodes().forEach(function (node) {
			root.removeChild(node);
		});
		for (index = 0; index < selected.options.length; index += 1) {
			option = selected.options[index];
			parts = String(option.text || "").split("--");
			voucher = createNode("DetData", "voucher");
			setAttr(voucher, "UnitNo", attr(oldVoucher, "UnitNo") || attr(oldVoucher, 0));
			setAttr(voucher, "UnitName", attr(oldVoucher, "UnitName") || attr(oldVoucher, 1));
			setAttr(voucher, "BookNo", attr(oldVoucher, "BookNo") || attr(oldVoucher, 2));
			setAttr(voucher, "BookName", attr(oldVoucher, "BookName") || attr(oldVoucher, 3));
			setAttr(voucher, "VouDate", attr(oldVoucher, "VouDate") || attr(oldVoucher, 4));
			setAttr(voucher, "Approver", attr(oldVoucher, "Approver") || attr(oldVoucher, 5));
			setAttr(voucher, "SalTransNo", option.value);
			setAttr(voucher, "SalVouNo", parts[0] || "");
			setAttr(voucher, "SalVouDate", parts[1] || "");
			setAttr(voucher, "CrTransNo", attr(oldVoucher, "CrTransNo") || attr(oldVoucher, 9));
			setAttr(voucher, "CrVoucherNo", attr(oldVoucher, "CrVoucherNo") || attr(oldVoucher, 10));
			setAttr(voucher, "TransNo", attr(oldVoucher, "TransNo") || attr(oldVoucher, 11));
			setAttr(voucher, "VoucherNo", attr(oldVoucher, "VoucherNo") || attr(oldVoucher, 12));
			if (party) {
				voucher.appendChild(root.ownerDocument.importNode ? root.ownerDocument.importNode(party, true) : party.cloneNode(true));
			}
			root.insertBefore(voucher, insertBefore);
		}
		syncOldInvoiceCodes();
	};

	window.UpdateNarrAmt = function () {
		var selected = field("selTobox");
		var narration = "For Invoice ";
		var amount = 0;
		var index;
		var parts;
		if (!selected) {
			return;
		}
		for (index = 0; index < selected.options.length; index += 1) {
			parts = String(selected.options[index].text || "").split("--");
			narration += (parts[0] || "") + " " + (parts[1] || "") + ", ";
			amount += toNumber(parts[2]);
		}
		narration = trim(narration).replace(/,\s*$/, "");
		setValue("txtNarration", narration);
		setValue("txtAmount", amount.toFixed(2));
	};

	var baseAddEntry = window.AddEntry;
	window.AddEntry = function (flag) {
		if (trim(valueOf("hEditEntry")) === "1") {
			window.UpdateXml();
			setMoveButtonsDisabled(true);
		}
		return baseAddEntry ? baseAddEntry(flag) : false;
	};

	var baseEditEntry = window.EditEntry;
	window.EditEntry = function (entryNo) {
		var result = baseEditEntry ? baseEditEntry(entryNo) : false;
		setMoveButtonsDisabled(String(entryNo) !== "1");
		return result;
	};

	var baseDisplayVoucher = window.DisplayVoucher;
	window.DisplayVoucher = function (callType) {
		if (String(callType || "").toUpperCase() === "B") {
			syncPartyCode();
			window.PopulateInvoices();
			syncOldInvoiceCodes();
		}
		return baseDisplayVoucher ? baseDisplayVoucher() : undefined;
	};

	window.InitVouCNCommissionAmd = function () {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		setDateLimits();
		window.DisplayVoucher("B");
	};
}(window, document));
