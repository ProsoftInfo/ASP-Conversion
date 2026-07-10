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

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
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
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var result = [];
		var i;
		if (!context) {
			return result;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return result;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (i = 0; i < found.snapshotLength; i += 1) {
			result.push(found.snapshotItem(i));
		}
		return result;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body);
		return xhr;
	}

	function importFor(parent, node) {
		if (!parent || !node) {
			return node;
		}
		if (node.ownerDocument !== parent.ownerDocument && parent.ownerDocument.importNode) {
			return parent.ownerDocument.importNode(node, true);
		}
		return node;
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
		return window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate ? window.ITMSModernCompat.toDisplayDate(control.value) : control.value || "";
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function rootFromDialog(value) {
		return xmlRoot(value) || value && value.nodeType === 9 && value.documentElement || value && value.nodeType === 1 && value || null;
	}

	function resetEntryRoot() {
		var entry = createNode("EntryData", "Entry");
		setAttr(entry, "No", "1");
		setAttr(entry, "CRDR", "D");
		setAttr(entry, "Payto", "0");
		setAttr(entry, "Amount", "0");
		setAttr(entry, "AccUnit", "0");
		setAttr(entry, "AccName", "");
		window.EntryRoot = entry;
		return entry;
	}

	function initState() {
		window.VouRoot = xmlRoot("VoucherData");
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		window.RequestRoot = window.RequestRoot && window.RequestRoot.nodeType ? window.RequestRoot : xmlRoot("RequestData");
		if (!window.EntryRoot || String(window.EntryRoot.nodeName).toLowerCase() !== "entry") {
			resetEntryRoot();
		}
		return {
			voucherRoot: window.VouRoot,
			entryRoot: window.EntryRoot,
			requestRoot: window.RequestRoot
		};
	}

	function updateAnalysisAmounts(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node) {
				var code = attr(node, "No") || attr(node, 0);
				var groupCode = attr(node, "GroupCode");
				var ratioField;
				var amountField;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				}
				if (header.nodeName === "Analytical") {
					ratioField = field("txtANALRatio" + code + "Z" + groupCode) || field("txtANALRatio" + code);
					amountField = field("txtANALAmount" + code + "Z" + groupCode) || field("txtANALAmount" + code);
				}
				if (header.nodeName === "PayRec") {
					amountField = field("txtDocAmount" + code);
				}
				if (ratioField) {
					setAttr(node, "Ratio", ratioField.value);
				}
				if (amountField) {
					setAttr(node, "Amount", amountField.value);
					if (node.attributes && node.attributes.length > 5) {
						node.attributes.item(5).nodeValue = amountField.value;
					}
				}
			});
		});
	}

	function clearGeneratedVoucherChildren(voucherRoot) {
		childElements(voucherRoot, "Entry").forEach(function (node) {
			voucherRoot.removeChild(node);
		});
		childElements(voucherRoot, "RequestDetails").forEach(function (node) {
			voucherRoot.removeChild(node);
		});
	}

	function clearGeneratedEntryChildren(entryRoot) {
		childElements(entryRoot, "AccHead").forEach(function (node) {
			entryRoot.removeChild(node);
		});
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
	}

	function instrumentTypeCode(typeValue) {
		var text = trim(typeValue).toLowerCase();
		if (text === "c" || text === "cheque") {
			return "C";
		}
		if (text === "d" || text === "demand draft") {
			return "D";
		}
		if (text === "b" || text === "bankers cheque") {
			return "B";
		}
		if (text === "t" || text === "rtgs") {
			return "T";
		}
		return "W";
	}

	function instrumentPrefix(typeValue) {
		var code = instrumentTypeCode(typeValue);
		if (code === "C") {
			return "C";
		}
		if (code === "D") {
			return "D";
		}
		if (code === "B") {
			return "B";
		}
		if (code === "T") {
			return "T";
		}
		return "Cash";
	}

	window.checkFileds = function () {
		var amount = valueOf("txtAmount");
		if (trim(valueOf("txtNarration")) === "") {
			alert("Enter Narration");
			if (field("txtNarration")) {
				field("txtNarration").select();
			}
			return false;
		}
		if (trim(amount) === "" || !isFinite(Number(String(amount).replace(/,/g, ""))) || toNumber(amount) === 0) {
			alert("Select atleast one term in Request Details Section");
			return false;
		}
		return true;
	};

	window.AddEntry = function () {
		var state = initState();
		var book = field("selBookId");
		var bookParts;
		var accHead;
		var narration;
		if (toNumber(valueOf("hAmount")) < toNumber(valueOf("txtAmount"))) {
			alert("Amount Should be Less then or equal to PayAmount");
			return false;
		}
		if (!book || book.selectedIndex === 0) {
			alert("Select Book");
			if (book) {
				book.focus();
			}
			return false;
		}
		if (!window.checkFileds()) {
			return false;
		}
		bookParts = String(book.value || "").split("?");
		setAttr(state.voucherRoot, "BookNo", bookParts[0] || "");
		setAttr(state.voucherRoot, "BookName", selectedText(book));
		setAttr(state.voucherRoot, "VouDate", getDateControl("ctlDate"));

		setAttr(state.entryRoot, "No", "1");
		setAttr(state.entryRoot, "CRDR", "D");
		setAttr(state.entryRoot, "Payto", valueOf("hParName"));
		setAttr(state.entryRoot, "Amount", valueOf("txtAmount"));
		setAttr(state.entryRoot, "AccUnit", valueOf("hUnitId"));
		setAttr(state.entryRoot, "AccName", valueOf("hOrgName"));
		clearGeneratedEntryChildren(state.entryRoot);

		accHead = createNode("EntryData", "AccHead");
		setAttr(accHead, "No", valueOf("hParValue"));
		setAttr(accHead, "Pay", "0");
		setAttr(accHead, "Rec", "0");
		setAttr(accHead, "Name", valueOf("hParName"));
		setAttr(accHead, "Type", "P");
		setAttr(accHead, "Adv", "0");
		state.entryRoot.appendChild(accHead);

		narration = createNode("EntryData", "Narration");
		narration.textContent = valueOf("txtNarration");
		state.entryRoot.appendChild(narration);

		updateAnalysisAmounts(state.entryRoot);
		clearGeneratedVoucherChildren(state.voucherRoot);
		state.voucherRoot.appendChild(importFor(state.voucherRoot, state.entryRoot));
		if (state.requestRoot) {
			state.voucherRoot.appendChild(importFor(state.voucherRoot, state.requestRoot));
		}
		return window.SaveXML();
	};

	window.SaveXML = function () {
		var xhr = syncPost("XMLSave.asp?Name=Payment%20Request&Mod=" + encodeURIComponent(valueOf("hVouName")), serializeXml("VoucherData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		form().submit();
		return true;
	};

	window.PopInsDet = function () {
		var tempValues = [
			valueOf("hVouType"),
			"",
			valueOf("hVouCode"),
			valueOf("hUnitId"),
			valueOf("hTransNo"),
			valueOf("hVouName")
		].join(":");
		openDialog("BankInsDetails.asp?sTemp=" + encodeURIComponent(tempValues), xmlObject("VoucherData"), "dialogHeight:250px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No", function (value) {
			var root = rootFromDialog(value);
			var node = selectNodes(root, "//BankInstrumentDet")[0];
			var insType;
			var insNo;
			var insDate;
			var payAt;
			var drawnOn;
			var prefix;
			if (!node) {
				return;
			}
			insType = attr(node, "InsType") || attr(node, 2);
			insNo = attr(node, "InsNo") || attr(node, 1);
			insDate = attr(node, "InsDate") || attr(node, 3);
			payAt = attr(node, "PayAt") || attr(node, 4);
			drawnOn = attr(node, "DrawnOn") || attr(node, 5);
			prefix = instrumentPrefix(insType);
			setText("spInsNo", prefix + ": " + insNo);
			setText("spInsDate", insDate);
			setValue("hInsDet", prefix + ": " + insNo + ":" + insType + ":" + payAt + ":" + drawnOn + ":" + insDate);
		});
		return false;
	};

	window.setPayableDisplay = function (flag) {
		var visible = Number(flag) !== 0;
		var additional = byId("Disaddtional");
		var payable = byId("DisPayable");
		if (additional) {
			additional.style.height = visible ? "115px" : "1px";
			additional.style.visibility = visible ? "visible" : "hidden";
		}
		if (payable) {
			payable.style.height = visible ? "110px" : "1px";
			payable.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.setAnalDisplay = function (display, flag) {
		var visible = Number(flag) !== 0;
		var item = byId(String(display) === "A" ? "DisAnal" : "DisCost");
		if (item) {
			item.style.height = visible ? "100px" : "1px";
			item.style.width = visible ? "280px" : "1px";
			item.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.setADDDisplay = function (flag) {
		var visible = Number(flag) !== 0;
		var additional = byId("Disaddtional");
		var ccanl = byId("DisCCANL");
		if (additional) {
			additional.style.height = visible ? "115px" : "1px";
			additional.style.visibility = visible ? "visible" : "hidden";
		}
		if (ccanl) {
			ccanl.style.height = visible ? "114px" : "1px";
			ccanl.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.clearXML = resetEntryRoot;

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.InitPmtGenRegularChqEntryForParty = function () {
		ensureCompat();
		initState();
	};
}(window, document));
