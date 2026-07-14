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
		var result = [];
		var frm = form();
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

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
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

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (context.selectNodes) {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		try {
			found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
			for (var index = 0; index < found.snapshotLength; index += 1) {
				nodes.push(found.snapshotItem(index));
			}
		} catch (ignore) {}
		return nodes;
	}

	function firstNode(context, expression) {
		var nodes = selectNodes(context, expression);
		return nodes.length ? nodes[0] : null;
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
		if (body != null) {
			try {
				xhr.setRequestHeader("Content-Type", "text/xml");
			} catch (ignore) {}
		}
		xhr.send(body == null ? null : body);
		return xhr;
	}

	function loadResponseXml(islandName, xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml(islandName, serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml(islandName, xhr.responseText);
		}
	}

	function createXmlElement(xmlName, nodeName) {
		var doc = xmlDocument(xmlName);
		return doc ? doc.createElement(nodeName) : document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function importFor(root, node) {
		var doc = root && root.ownerDocument;
		if (!doc || !node) {
			return node;
		}
		return node.ownerDocument !== doc && doc.importNode ? doc.importNode(node, true) : node;
	}

	function populateSelectFromXml(select, root, textIndex, valueIndex) {
		if (!select) {
			return;
		}
		childElements(root).forEach(function (node) {
			select.add(new Option(attr(node, textIndex), attr(node, valueIndex)));
		});
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

	function setMinMaxDate(name, minDate, maxDate) {
		var control = field(name) || byId(name);
		if (!control) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(minDate);
		} else if (typeof control.setMinDate === "function") {
			control.setMinDate(minDate);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.min = window.ITMSModernCompat.toIsoDate(minDate);
		}
		if (typeof control.SetMaxDate === "function") {
			control.SetMaxDate(maxDate);
		} else if (typeof control.setMaxDate === "function") {
			control.setMaxDate(maxDate);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.max = window.ITMSModernCompat.toIsoDate(maxDate);
		}
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		alert("The compatibility script is still loading. Please try again.");
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

	function lastAccHeadNode() {
		var nodes = childElements(xmlRoot("AccHeadData"));
		return nodes.length ? nodes[nodes.length - 1] : null;
	}

	function addPartyHeadXml(code, name, value2) {
		var parts = String(value2 || "").split(":");
		var root = xmlRoot("AccHeadData");
		var elem;
		if (typeof window.GetPartyHeadXml === "function") {
			window.GetPartyHeadXml(code, name, value2);
			return;
		}
		elem = createXmlElement("AccHeadData", "AccHead");
		setAttr(elem, "No", trim(code));
		setAttr(elem, "Pay", trim(parts[0]));
		setAttr(elem, "Rec", trim(parts[1]));
		setAttr(elem, "Name", name);
		setAttr(elem, "Type", "P");
		setAttr(elem, "Adv", trim(parts[2]));
		if (root) {
			root.appendChild(elem);
		}
	}

	function removeAgentDetails(root) {
		childElements(root, "AgentDetails").forEach(function (node) {
			root.removeChild(node);
		});
	}

	function updateAgentDisplay(root) {
		var agent = firstNode(root, "//Agent");
		var name = agent ? attr(agent, 1) || attr(agent, "Agentname") : "";
		setText("spAgentName", name);
		setValue("hCommName", name);
	}

	window.popPartyType = function () {
		var select = field("selPartyType");
		var orgId = valueOf("selUnitId");
		var xhr;
		if (!select) {
			return true;
		}
		select.options.length = 1;
		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(orgId) + "&sCallTy=P");
		loadResponseXml("OutData", xhr);
		childElements(xmlRoot("OutData")).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, 0)));
		});
		return true;
	};

	window.DisplayBook = function (objUnit) {
		var unit = objUnit || field("selUnitId");
		var select = field("selBook");
		var orgId;
		var xhr;
		if (select) {
			select.options.length = 1;
		}
		if (!unit || unit.selectedIndex === 0 || String(unit.selectedIndex) === "0") {
			return false;
		}
		orgId = unit.options[unit.selectedIndex].value;
		xhr = syncGet("XMLGetOrgBook.asp?BkCode=04&orgID=" + encodeURIComponent(orgId));
		loadResponseXml("UnitBookData", xhr);
		populateSelectFromXml(select, xmlRoot("UnitBookData"), 1, 0);
		window.popPartyType();
		return true;
	};

	window.validate = function () {
		var unit = field("selUnitId");
		var book = field("selBook");
		var purType = field("selPurType");
		var partyType = field("selPartyType");
		var partyName = field("txtPartyName");
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
		if (purType && purType.selectedIndex < 1) {
			alert("Select Purchase type");
			purType.focus();
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
		return true;
	};

	window.VouCreate = function () {
		var root = xmlRoot("VoucherData");
		var selectedDateText = getDateControl("ctlDate");
		var currentDate = parseDate(valueOf("hCurrDate"));
		var selectedDate = parseDate(selectedDateText);
		var invoiceNo = field("txtInvoiceNo");
		var unit = field("selUnitId");
		var book = field("selBook");
		var purType = field("selPurType");
		var partyType = field("selPartyType");
		var header;
		var node;
		var partyParts;
		var xhr;
		if (!root || !window.validate()) {
			return false;
		}
		if (!invoiceNo || trim(invoiceNo.value) === "") {
			alert("Enter Invoice Number ");
			if (invoiceNo) {
				invoiceNo.focus();
			}
			return false;
		}
		if (selectedDate && currentDate && selectedDate > currentDate) {
			alert("Voucher Date Should be Less than the System Date ");
			return false;
		}
		xhr = syncGet("CheckInvCreate.asp?sValue=" + encodeURIComponent(valueOf("hPartyCode") + "?" + invoiceNo.value + "?" + selectedDateText + "?04?" + valueOf("selUnitId")));
		if (String(trim(xhr.responseText)) !== "C") {
			alert("Purchase Voucher already Created for this Party,InvoiceNo and Invoice Date ");
			return false;
		}

		header = createXmlElement("VoucherData", "Header");
		root.appendChild(header);
		setValue("hInvDate", selectedDateText);
		setValue("hOrgName", selectedText(unit));

		node = createXmlElement("VoucherData", "Organization");
		setAttr(node, "OrgId", valueOf("selUnitId"));
		node.textContent = selectedText(unit);
		header.appendChild(node);

		childElements(xmlRoot("UnitBookData")).forEach(function (bookNode) {
			var bookXml;
			if (!book || attr(bookNode, 0) !== book.value) {
				return;
			}
			setValue("hBkAccHead", attr(bookNode, 2));
			bookXml = createXmlElement("VoucherData", "Book");
			setAttr(bookXml, "BookId", book.value);
			setAttr(bookXml, "BKAccHead", attr(bookNode, 2));
			setAttr(bookXml, "BKOtherUnits", attr(bookNode, 3));
			bookXml.textContent = selectedText(book);
			header.appendChild(bookXml);
		});

		node = createXmlElement("VoucherData", "PurchaseType");
		setAttr(node, "PurTypeId", valueOf("selPurType"));
		node.textContent = selectedText(purType);
		header.appendChild(node);

		node = createXmlElement("VoucherData", "PurInvoice");
		setAttr(node, "PurInvNo", valueOf("txtInvoiceNo"));
		setAttr(node, "PurInvDate", selectedDateText);
		header.appendChild(node);

		node = createXmlElement("VoucherData", "Party");
		partyParts = trim(valueOf("hPartyCode")).split("?");
		setAttr(node, "ParType", partyParts[0] || "");
		setAttr(node, "ParSubType", partyParts[1] || "");
		setAttr(node, "ParSubTypeName", selectedText(partyType));
		setAttr(node, "ParCode", partyParts[3] || "");
		node.textContent = valueOf("txtPartyName");
		header.appendChild(node);

		window.SaveXML();
		return false;
	};

	window.selAccountHead = function (objAcc) {
		var select = objAcc || field("selPartyType");
		var orgId = valueOf("selUnitId");
		var partyType;
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		partyType = select.value + "?" + selectedText(select);
		clearChildren(xmlRoot("AccHeadData"));
		runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType), "dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue, parts) {
			var node;
			if (parts.length <= 2 || parts[0] === "-1") {
				select.selectedIndex = 0;
				return;
			}
			addPartyHeadXml(parts[1] || "", parts[0] || "", "0:0:0");
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

	window.showAgent = function (flag) {
		var root = xmlRoot("VoucherData");
		var noAgent = fields("optAgentExist")[2];
		if (field("selUnitId") && field("selUnitId").selectedIndex < 1) {
			alert("Select Unit");
			field("selUnitId").focus();
			if (noAgent) {
				noAgent.checked = true;
			}
			return false;
		}
		if (!root) {
			return false;
		}
		removeAgentDetails(root);
		if (String(flag) === "N") {
			setText("spAgentName", "");
			setValue("hCommName", "");
			return false;
		}
		openDialog("AgentCommisionEntry.asp?OrgID=" + encodeURIComponent(valueOf("selUnitId")) + "&AgentType=" + encodeURIComponent(flag), xmlObject("OutData"), "dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (returnValue) {
			var returnedRoot = xmlRoot(returnValue) || returnValue && returnValue.nodeType === 1 && returnValue || null;
			if (returnedRoot && childElements(returnedRoot).length) {
				root.appendChild(importFor(root, returnedRoot));
				updateAgentDisplay(root);
			} else if (noAgent) {
				noAgent.checked = true;
			}
		});
		return false;
	};

	window.SaveXML = function () {
		var xhr = syncPost("XMLSave.asp?Mod=PUR&Name=Voucher Entry", serializeXml("VoucherData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else if (form()) {
			setValue("hInvDate", getDateControl("ctlDate"));
			window.DisButt();
			form().submit();
		}
		return false;
	};

	window.popVoucherNo = function (sVal) {
		var frm = form();
		var width;
		var url;
		if (typeof window.validateForDel === "function" && !window.validateForDel()) {
			return false;
		}
		if (!frm) {
			return false;
		}
		width = String(sVal) !== "A" ? "450" : "500";
		url = "VouchSelForSalPur.asp?flag=" + encodeURIComponent(sVal) +
			"&orgId=" + encodeURIComponent(valueOf("selUnitId")) +
			"&BookCode=04" +
			"&BookNo=" + encodeURIComponent(valueOf("selBook")) +
			"&TransType=PJR" +
			"&sPurTy=" + encodeURIComponent(valueOf("selPurType")) +
			"&sParTy=" + encodeURIComponent(valueOf("selPartyType")) +
			"&iParCode=" + encodeURIComponent(valueOf("hPartyCode"));
		openDialog(url, "", "dialogHeight:400px;dialogWidth:" + width + "px;center:Yes;help:No;resizable:No;status:No", function (sTemp) {
			var parts;
			var purInvValue;
			var purInvDate;
			var purInvNo;
			if (String(sTemp || "0") === "0") {
				return;
			}
			parts = String(sTemp).split("~");
			setValue("hTransNo", parts[0] || "");
			purInvValue = parts[1] || "";
			purInvDate = purInvValue.length >= 10 ? trim(purInvValue.slice(-10)) : "";
			purInvNo = purInvDate ? purInvValue.slice(0, Math.max(0, purInvValue.length - 11)) : purInvValue;
			setValue("txtInvoiceNo", trim(purInvNo));
			setValue("hInvDate", parts[3] || "");
			setValue("hVouDate", purInvDate);
			if (purInvDate) {
				setDateControl("ctlDate", purInvDate);
			}
			if (String(sVal) === "A") {
				if (field("btnAmend")) {
					field("btnAmend").disabled = true;
				}
				if (field("btnDel")) {
					field("btnDel").disabled = true;
				}
			} else {
				if (field("btnAmend")) {
					field("btnAmend").disabled = false;
				}
				if (field("btnDel")) {
					field("btnDel").disabled = false;
				}
			}
		});
		return false;
	};

	window.validateForDel = function () {
		var unit = field("selUnitId");
		var book = field("selBook");
		if (unit && unit.selectedIndex < 1) {
			alert("Select Unit ");
			unit.focus();
			return false;
		}
		if (book && book.selectedIndex < 1) {
			alert("Select Purchase Book");
			book.focus();
			return false;
		}
		return true;
	};

	window.VouView = function () {
		if (!window.validateForDel()) {
			return false;
		}
		if (trim(valueOf("hTransNo")) === "") {
			alert("Not a Valid Voucher Number ");
			return false;
		}
		form().action = "VouPurView.asp";
		window.DisButt();
		form().submit();
		return true;
	};

	window.VouDel = function () {
		if (!window.validateForDel()) {
			return false;
		}
		form().action = "VouPurDelView.asp";
		window.DisButt();
		form().submit();
		return true;
	};

	window.VouAmend = function () {
		if (!window.validateForDel()) {
			return false;
		}
		setValue("hInvDate", getDateControl("ctlDate"));
		form().action = "VouPURAmdDetails.asp";
		window.DisButt();
		form().submit();
		return true;
	};

	window.DisButt = function () {
		["btnAmend", "btnCreate", "btnDel", "btnView"].forEach(function (name) {
			if (field(name)) {
				field(name).disabled = true;
			}
		});
	};

	window.SetDate = function () {
		setMinMaxDate("ctlDate", "01/04/" + trim(valueOf("hFromYr")), "31/03/" + valueOf("hToYr"));
		return true;
	};

	window.PopulateSalTy = function () {
		return true;
	};
}(window, document));
