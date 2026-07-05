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
		var lower;
		var index;
		if (!frm || !frm.elements || !name) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		lower = String(name).toLowerCase();
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === lower) {
				return frm.elements[index];
			}
		}
		return null;
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

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
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

	function loadResponseXml(islandName, xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml(islandName, serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml(islandName, xhr.responseText);
		}
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName || "").toLowerCase() === wanted);
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
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
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
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (index = 0; index < found.snapshotLength; index += 1) {
			nodes.push(found.snapshotItem(index));
		}
		return nodes;
	}

	function firstNode(context, expression) {
		var nodes = selectNodes(context, expression);
		return nodes.length ? nodes[0] : null;
	}

	function selectOptionByValue(select, value) {
		if (!select) {
			return false;
		}
		return Array.prototype.some.call(select.options, function (option, index) {
			if (String(option.value) === String(value)) {
				select.selectedIndex = index;
				return true;
			}
			return false;
		});
	}

	function addOption(select, text, value) {
		if (select) {
			select.add(new Option(text == null ? "" : String(text), value == null ? "" : String(value)));
		}
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body || "");
		return xhr;
	}

	function parseDate(value) {
		var text = trim(value);
		var parts;
		var monthNames = {
			jan: 0,
			feb: 1,
			mar: 2,
			apr: 3,
			may: 4,
			jun: 5,
			jul: 6,
			aug: 7,
			sep: 8,
			oct: 9,
			nov: 10,
			dec: 11
		};
		if (!text) {
			return null;
		}
		if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
			parts = text.split("-");
			return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
		}
		parts = text.split(/[\/\-]/);
		if (parts.length >= 3) {
			return new Date(
				Number(parts[2]),
				isNaN(Number(parts[1])) ? monthNames[String(parts[1]).substring(0, 3).toLowerCase()] : Number(parts[1]) - 1,
				Number(parts[0])
			);
		}
		return null;
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

	function openDialog(url, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features || "", callback || function () {});
			return;
		}
		window.open(url, "_blank", "height=500,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function runStringDialog(page, query, features, done) {
		openDialog(page + "?" + query, features, function (outValue) {
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
			done(parts);
		});
	}

	function orgSelect() {
		return field("selUnitId");
	}

	function orgId() {
		var select = orgSelect();
		if (select && select.selectedIndex >= 0) {
			return select.value;
		}
		return valueOf("horgID", valueOf("hOrgID"));
	}

	function orgName() {
		var select = orgSelect();
		if (select && select.selectedIndex >= 0) {
			return selectedText(select);
		}
		return valueOf("horgName", valueOf("hOrgName"));
	}

	function lastAccHeadNode() {
		var nodes = childElements(xmlRoot("AccHeadData"));
		return nodes.length ? nodes[nodes.length - 1] : null;
	}

	window.popPartyType = function () {
		var select = field("selPartyType");
		var xhr;
		if (!select) {
			return true;
		}
		select.options.length = 1;
		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(orgId()) + "&sCallTy=P");
		loadResponseXml("OutData", xhr);
		childElements(xmlRoot("OutData")).forEach(function (node) {
			addOption(select, trim(node.textContent || node.text || ""), attr(node, 0));
		});
		return true;
	};

	window.DisplayBook = function (objUnit) {
		var select = field("selBook");
		var unitValue = typeof objUnit === "string" ? objUnit : objUnit && objUnit.options ? objUnit.options[objUnit.selectedIndex].value : orgId();
		var xhr;
		if (select) {
			select.options.length = 1;
		}
		if (!unitValue || unitValue === "0") {
			return false;
		}
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=04&orgID=" + encodeURIComponent(unitValue));
		loadResponseXml("UnitBookData", xhr);
		childElements(xmlRoot("UnitBookData")).forEach(function (node) {
			addOption(select, attr(node, 1), attr(node, 0));
		});
		window.popPartyType();
		return true;
	};

	window.validate = function () {
		var unit = orgSelect();
		var book = field("selBook");
		var partyType = field("selPartyType");
		var partyName = field("txtPartyName");
		var invoiceNo = field("txtInvoiceNo");
		var selectedDate = parseDate(getDateControl("ctlDate"));
		var currentDate = parseDate(valueOf("hCurrDate"));
		if (unit && unit.selectedIndex < 1) {
			alert("Select Unit");
			unit.focus();
			return false;
		}
		if (book && book.selectedIndex < 1) {
			alert("Select Purchase Book");
			book.focus();
			return false;
		}
		if (partyType && partyType.selectedIndex < 1) {
			alert("Select Party");
			partyType.focus();
			return false;
		}
		if (partyName && trim(partyName.value) === "") {
			alert("Party Name should not be blank");
			partyName.select();
			return false;
		}
		if (invoiceNo && trim(invoiceNo.value) === "") {
			alert("Invoice No should not be blank");
			invoiceNo.select();
			return false;
		}
		if (selectedDate && currentDate && selectedDate > currentDate) {
			alert("Voucher Date Should be Less than the System Date ");
			return false;
		}
		return true;
	};

	window.VouCreate = function () {
		var root = xmlRoot("OldVouData");
		var node;
		var parts;
		var book = field("selBook");
		var purType = field("selPurType");
		if (!root || !window.validate()) {
			return false;
		}
		node = firstNode(root, "//Organization");
		if (node) {
			setAttr(node, "OrgId", orgId());
			if (node.hasAttribute && node.hasAttribute("AccUnit")) {
				setAttr(node, "AccUnit", orgId());
			}
			node.textContent = orgName();
		}
		node = firstNode(root, "//Book");
		if (node) {
			setAttr(node, "BookId", valueOf("selBook"));
			node.textContent = selectedText(book);
		}
		node = firstNode(root, "//Party");
		parts = trim(valueOf("hPartyCode")).split("?");
		if (node) {
			setAttr(node, "ParType", parts[0] || "");
			setAttr(node, "ParSubType", parts[1] || "");
			setAttr(node, "ParCode", parts[3] || "");
			node.textContent = valueOf("txtPartyName");
		}
		node = firstNode(root, "//PurInvoice");
		if (node) {
			setAttr(node, "PurInvNo", valueOf("txtInvoiceNo"));
			setAttr(node, "PurInvDate", getDateControl("ctlDate"));
		}
		node = firstNode(root, "//PurCategory");
		if (node) {
			setAttr(node, "Code", valueOf("selPurCat"));
			node.textContent = valueOf("selPurCat");
		}
		node = firstNode(root, "//PurchaseType");
		if (node) {
			setAttr(node, "PurTypeId", valueOf("selPurType"));
			node.textContent = selectedText(purType);
		}
		window.SaveXML();
		return false;
	};

	window.selAccountHead = function (objAcc) {
		var select = objAcc || field("selPartyType");
		var partyType;
		if (!select || select.selectedIndex <= 0) {
			setValue("txtPartyName", "");
			setValue("hPartyCode", "0");
			return false;
		}
		partyType = select.value + "?" + selectedText(select).replace(/&/g, " and ");
		runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(orgId()) + "&Party=" + encodeURIComponent(partyType), "dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (parts) {
			var node;
			if (parts.length <= 2) {
				select.selectedIndex = 0;
				setValue("txtPartyName", "");
				setValue("hPartyCode", "0");
				return;
			}
			if (typeof window.GetPartyHeadXml === "function") {
				window.GetPartyHeadXml(parts[1] || "", parts[0] || "", "0:0:0");
			}
			node = lastAccHeadNode();
			if (node) {
				setValue("hPartyCode", partyType + "?" + attr(node, 0));
				setValue("txtPartyName", attr(node, 3));
			} else {
				select.selectedIndex = 0;
			}
		});
		return false;
	};

	window.SaveXML = function () {
		var xhr = syncPost("XMLSave.asp?Mod=PUR&Name=Voucher AMD", serializeXml("OldVouData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else {
			setValue("hInvDate", getDateControl("ctlDate"));
			window.DisButt();
			form().submit();
		}
		return false;
	};

	window.DisButt = function () {
		var button = field("btnCreate");
		if (button) {
			button.disabled = true;
		}
	};

	window.DispOldVal = function () {
		var root = xmlRoot("OldVouData");
		var node;
		var value;
		var partType;
		var partSubType;
		var partCode;
		var select;
		if (!root) {
			return false;
		}
		node = firstNode(root, "//Organization");
		if (node) {
			value = attr(node, "OrgId");
			select = orgSelect();
			if (select) {
				selectOptionByValue(select, value);
				window.DisplayBook(select);
			} else {
				window.DisplayBook(valueOf("horgID", valueOf("hOrgID")));
			}
		}
		node = firstNode(root, "//Book");
		if (node) {
			setValue("hBkAccHead", attr(node, "BKAccHead"));
			selectOptionByValue(field("selBook"), attr(node, "BookId"));
		}
		node = firstNode(root, "//Party");
		if (node) {
			partType = attr(node, "ParType");
			partSubType = attr(node, "ParSubType");
			partCode = attr(node, "ParCode");
			setValue("txtPartyName", node.textContent || "");
			value = partType + "?" + partSubType;
			selectOptionByValue(field("selPartyType"), trim(value));
			setValue("hPartyCode", value + "?" + (node.textContent || "") + "?" + partCode);
		}
		node = firstNode(root, "//PurInvoice");
		if (node) {
			setValue("txtInvoiceNo", attr(node, "PurInvNo"));
			setDateControl("ctlDate", attr(node, "PurInvDate"));
		}
		node = firstNode(root, "//PurCategory");
		if (node) {
			selectOptionByValue(field("selPurCat"), attr(node, "Code"));
		}
		node = firstNode(root, "//PurchaseType");
		if (node) {
			selectOptionByValue(field("selPurType"), attr(node, "PurTypeId"));
		}
		return true;
	};

	window.ITMSPurchaseAmdBookSelectionCompat = {
		init: ensureCompat
	};

	ensureCompat();
})(window, document);
