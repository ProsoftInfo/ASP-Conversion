(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var lowerName;
		var i;
		if (!frm || !name) {
			return null;
		}
		if (frm.elements[name] || frm[name]) {
			return frm.elements[name] || frm[name];
		}
		lowerName = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === lowerName) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function firstField(item) {
		return item && item.length && !item.tagName ? item[0] : item;
	}

	function valueOf(name, fallback) {
		var item = firstField(field(name));
		return item ? item.value || "" : fallback == null ? "" : fallback;
	}

	function setValue(name, value) {
		var item = firstField(field(name));
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || null;
	}

	function setVisible(id, visible, height, width) {
		var element = byId(id);
		if (!element) {
			return;
		}
		element.style.height = visible ? height || "" : "1px";
		if (width != null) {
			element.style.width = visible ? width : "1px";
		}
		element.style.visibility = visible ? "visible" : "hidden";
	}

	function requestText(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr.responseText || "";
	}

	function parseXml(text) {
		if (!window.DOMParser || trim(text) === "") {
			return null;
		}
		return new DOMParser().parseFromString(text, "text/xml");
	}

	function xmlObject(name) {
		var element;
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		element = byId(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var element;
		var doc;
		if (object && object.XMLDocument) {
			return object.XMLDocument;
		}
		if (object && object._doc) {
			return object._doc;
		}
		if (object && object.nodeType === 9) {
			return object;
		}
		if (object && object.ownerDocument && String(object.tagName || "").toLowerCase() === "xml") {
			element = object;
			doc = parseXml(element.innerHTML || element.textContent || "");
			if (!doc && element.getAttribute("src")) {
				doc = parseXml(requestText(element.getAttribute("src")));
			}
			if (doc) {
				element._doc = doc;
				return doc;
			}
		}
		return document.implementation && document.implementation.createDocument ? document.implementation.createDocument("", "", null) : null;
	}

	function xmlRoot(nameOrNode) {
		var doc;
		if (typeof nameOrNode === "string") {
			doc = xmlDocument(nameOrNode);
			return doc && doc.documentElement || null;
		}
		if (!nameOrNode) {
			return null;
		}
		return nameOrNode.documentElement || nameOrNode;
	}

	function entryRoot() {
		return window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
	}

	function entryDoc() {
		var root = entryRoot();
		return root && root.ownerDocument || xmlDocument("EntryData");
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toLowerCase() : "";
		var i;
		if (!node) {
			return result;
		}
		for (i = 0; i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
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

	function attr(node, nameOrIndex) {
		var lowerName;
		var i;
		var value;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			return node.attributes.item(nameOrIndex) ? node.attributes.item(nameOrIndex).value || "" : "";
		}
		value = node.getAttribute && node.getAttribute(nameOrIndex);
		if (value != null) {
			return value;
		}
		lowerName = String(nameOrIndex).toLowerCase();
		for (i = 0; i < node.attributes.length; i += 1) {
			if (String(node.attributes[i].name || "").toLowerCase() === lowerName) {
				return node.attributes[i].value || "";
			}
		}
		return "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function appendElement(doc, parent, name) {
		var node = doc.createElement(name);
		parent.appendChild(node);
		return node;
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function clearTableRows(tableId, startIndex, keepCount) {
		var table = byId(tableId);
		var start = Math.max(0, Number(startIndex) || 0);
		var keep = Math.max(0, Number(keepCount) || 0);
		while (table && table.rows.length > start + keep) {
			table.deleteRow(start);
		}
	}

	function insertCell(row, type, name, value, className, align, maxLength, size, extraStyle) {
		var cell = row.insertCell(-1);
		var input;
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		if (String(type) === "2") {
			input = document.createElement("input");
			input.type = "text";
			input.name = name || "";
			input.id = name || "";
			input.value = value == null ? "" : String(value);
			input.className = "FormElem";
			if (size) {
				input.size = size;
			}
			if (maxLength) {
				input.maxLength = maxLength;
			}
			if (extraStyle) {
				input.setAttribute("style", String(extraStyle).replace(/^style=/i, "").replace(/^"|"$/g, ""));
			}
			cell.appendChild(input);
		} else {
			cell.innerHTML = value == null ? "" : String(value);
		}
		return cell;
	}

	function addExcelCell(row, type, name, value, className, align, maxLength, size, extraStyle) {
		if (typeof window.InsertCell === "function") {
			window.InsertCell(row, type, name || "", value == null ? "" : String(value), className || "", align || "", "", size || 0, maxLength || 0, 0, 0, extraStyle || "");
			return;
		}
		insertCell(row, type, name, value, className, align, maxLength, size, extraStyle);
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return null;
		}
		window.open(url, "_blank", "height=500,width=700,resizable=yes,status=no,scrollbars=yes");
		return null;
	}

	function receiveDialogResult(result, callback) {
		if (result != null && callback) {
			callback(result);
		}
	}

	function selectHead(sAccHead, sType, objHead, iHeadCount) {
		var head = objHead || firstField(field("selAccHead"));
		var splitHead;
		var i;
		var value;
		if (!head) {
			return;
		}
		if (String(sType) === "G") {
			for (i = 0; i < head.length; i += 1) {
				value = String(head.options[i].value || "").split("?")[0];
				if (value === String(sAccHead)) {
					head.selectedIndex = i;
					return;
				}
			}
			for (i = 0; i < head.length; i += 1) {
				if (String(head.options[i].value) === String(sAccHead)) {
					head.selectedIndex = i;
					return;
				}
			}
			head.selectedIndex = Number(iHeadCount || 0) + 1;
			return;
		}
		splitHead = String(sAccHead || "").split("?");
		sAccHead = trim(splitHead[0] || "") + "?" + trim(splitHead[1] || "");
		for (i = 0; i < head.length; i += 1) {
			if (String(head.options[i].value) === String(sAccHead)) {
				head.selectedIndex = i;
				return;
			}
		}
		head.selectedIndex = Number(iHeadCount || 0) + 1;
	}

	function checkAccHead(root, accHead) {
		return selectNodes(root, "//AccHead[@No='" + String(accHead).replace(/'/g, "&apos;") + "']").length > 0;
	}

	function validateAmount(amount, name, from, to) {
		var label = name || "Amount";
		var min = from == null ? 0 : toNumber(from);
		var max = to == null ? 9999999999.99 : toNumber(to);
		if (trim(amount) === "") {
			alert(label + " Cannot be blank");
			return false;
		}
		if (isNaN(Number(String(amount).replace(/,/g, "")))) {
			alert("Enter Numeric values for " + label);
			return false;
		}
		if (toNumber(amount) <= min || toNumber(amount) > max) {
			alert(label + " should be >" + min + " and < " + max);
			return false;
		}
		return true;
	}

	function distributeAmount() {
		var root = entryRoot();
		var amount;
		var children;
		var ratio;
		var eachAmount;
		var total;
		var ratioTotal;
		var i;
		if (typeof window.checkFileds === "function" && !window.checkFileds()) {
			setValue("txtAmount", "");
			return;
		}
		childElements(root).forEach(function (node) {
			children = childElements(node);
			if (!children.length || !/^(CostCenter|Analytical)$/i.test(node.nodeName)) {
				return;
			}
			amount = toNumber(valueOf("txtAmount"));
			total = amount;
			ratioTotal = 0;
			ratio = Math.round((100 / children.length) * 100) / 100;
			eachAmount = Math.round(((ratio * amount) / 100) * 100) / 100;
			for (i = 0; i < children.length; i += 1) {
				var code = attr(children[i], "No");
				var groupCode = attr(children[i], "GroupCode");
				var ratioField = node.nodeName === "CostCenter" ? "txtCCRatio" + code : "txtANALRatio" + code + "Z" + groupCode;
				var amountField = node.nodeName === "CostCenter" ? "txtCCAmount" + code : "txtANALAmount" + code + "Z" + groupCode;
				if (i < children.length - 1) {
					setValue(ratioField, ratio);
					setValue(amountField, eachAmount);
					setAttr(children[i], "Ratio", ratio);
					setAttr(children[i], "Amount", eachAmount);
					total -= eachAmount;
					ratioTotal += ratio;
				} else {
					setValue(ratioField, 100 - ratioTotal);
					setValue(amountField, total);
					setAttr(children[i], "Ratio", 100 - ratioTotal);
					setAttr(children[i], "Amount", total);
				}
			}
		});
		childElements(root, "PayRec").forEach(function (node) {
			childElements(node).forEach(function (payNode, index) {
				var docNo = attr(payNode, "No");
				var payableNo = trim(attr(payNode, "PayableNo"));
				var adjust = toNumber(attr(payNode, "TransAmount")) - toNumber(attr(payNode, "AmtAdjusted")) - toNumber(attr(payNode, "AmtToAccount"));
				setValue("txtDocAmount" + docNo + "Z" + payableNo + "Z" + (index + 1), formatNumber(adjust, 2));
			});
		});
	}

	function clearXml() {
		var doc = entryDoc();
		var root = doc && doc.createElement("Entry");
		if (!root) {
			return null;
		}
		setAttr(root, "No", window.iEntryNo || "1");
		setAttr(root, "CRDR", "0");
		setAttr(root, "Payto", "0");
		setAttr(root, "Amount", "0");
		setAttr(root, "AccUnit", "0");
		setAttr(root, "AccName", "");
		setAttr(root, "TdsAmount", "0");
		setAttr(root, "TDSElgi", "0");
		setAttr(root, "TdsPercentage", "0");
		setAttr(root, "PayRecAmount", "0");
		window.EntryRoot = root;
		return root;
	}

	function showNarration(bookCode) {
		var book = String(valueOf("selBook")).split("-")[0];
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")) + "&BookCode=" + encodeURIComponent(bookCode + "?" + book);
		receiveDialogResult(openDialog(url, "", "", function (value) {
			if (trim(value) !== "") {
				setValue("txtNarration", value);
			}
		}), function (value) {
			if (trim(value) !== "") {
				setValue("txtNarration", value);
			}
		});
	}

	function popCostCenter(headerNode) {
		var rows = childElements(headerNode);
		var table = byId("tblCost");
		var row;
		var i;
		if (!rows.length || !table) {
			setAnalDisplay("C", 0);
			return;
		}
		setAnalDisplay("C", 1);
		clearTableRows("tblCost", 1, 1);
		for (i = 0; i < rows.length; i += 1) {
			row = table.insertRow(i + 1);
			addExcelCell(row, 1, "", i + 1, "ExcelSerial", "Center");
			addExcelCell(row, 1, "", attr(rows[i], "ShortName") || attr(rows[i], 2), "ExcelDisplayCell", "left");
			addExcelCell(row, 2, "txtCCRatio" + attr(rows[i], "No"), attr(rows[i], "Ratio") || attr(rows[i], 3), "ExcelInputCell", "", 5, 6);
			addExcelCell(row, 2, "txtCCAmount" + attr(rows[i], "No"), attr(rows[i], "Amount") || attr(rows[i], 4), "ExcelInputCell", "", 10, 12);
		}
	}

	function popAnalytical(headerNode) {
		var rows = childElements(headerNode);
		var table = byId("tblAnal");
		var row;
		var i;
		if (!rows.length || !table) {
			setAnalDisplay("A", 0);
			return;
		}
		setAnalDisplay("A", 1);
		clearTableRows("tblAnal", 1, 1);
		for (i = 0; i < rows.length; i += 1) {
			row = table.insertRow(i + 1);
			addExcelCell(row, 1, "", i + 1, "ExcelSerial", "Center");
			addExcelCell(row, 1, "", attr(rows[i], "ShortName") || attr(rows[i], 2), "ExcelDisplayCell");
			addExcelCell(row, 2, "txtANALRatio" + attr(rows[i], "No") + "Z" + attr(rows[i], "GroupCode"), attr(rows[i], "Ratio") || attr(rows[i], 3), "ExcelInputCell", "", 5, 6);
			addExcelCell(row, 2, "txtANALAmount" + attr(rows[i], "No") + "Z" + attr(rows[i], "GroupCode"), attr(rows[i], "Amount") || attr(rows[i], 4), "ExcelInputCell", "", 10, 12);
		}
	}

	function appendImported(parent, node) {
		var doc = parent && parent.ownerDocument;
		if (!parent || !node) {
			return;
		}
		parent.appendChild(doc.importNode ? doc.importNode(node, true) : node.cloneNode(true));
	}

	function processCcAnalResult(root) {
		var entry = entryRoot();
		var accepted = root && (attr(root, 0) === "1" || attr(root, "RetVal") === "1" || attr(root, "Value") === "1");
		if (!accepted) {
			setADDDisplay(0);
			return;
		}
		setADDDisplay(1);
		childElements(root).forEach(function (node) {
			appendImported(entry, node);
			if (String(node.nodeName).toLowerCase() === "costcenter") {
				popCostCenter(node);
			} else if (String(node.nodeName).toLowerCase() === "analytical") {
				popAnalytical(node);
			}
		});
	}

	function showCCAnal(orgId, accCode, costCenter, analytical) {
		var transNo = valueOf("hTransNo");
		var entNo = valueOf("hEntryNo");
		var url;
		var result;
		if (Number(costCenter) !== 1 && Number(analytical) !== 1) {
			setADDDisplay(0);
			return;
		}
		url = "CCAnalysisSelection.asp?orgId=" + encodeURIComponent(orgId) + "&AccCode=" + encodeURIComponent(accCode) +
			"&TransNo=" + encodeURIComponent(transNo) + "&EntNo=" + encodeURIComponent(entNo);
		result = openDialog(url, "", "", function (value) {
			processCcAnalResult(xmlRoot(value));
		});
		if (result) {
			processCcAnalResult(xmlRoot(result));
		}
	}

	function popPayRec(headerNode) {
		var rows = childElements(headerNode);
		var table = byId("tblPayable");
		var row;
		var i;
		if (!rows.length || !table) {
			return;
		}
		setPayableDisplay(1);
		clearTableRows("tblPayable", 2, 1);
		for (i = 0; i < rows.length; i += 1) {
			var remaining = toNumber(attr(rows[i], "TransAmount")) - toNumber(attr(rows[i], "AmtAdjusted")) - toNumber(attr(rows[i], "AmtToAccount"));
			row = table.insertRow(i + 2);
			addExcelCell(row, 1, "", i + 1, "ExcelSerial", "Center");
			addExcelCell(row, 1, "", attr(rows[i], "InvNo"), "ExcelDisplayCell");
			addExcelCell(row, 1, "", attr(rows[i], "InvDate"), "ExcelDisplayCell");
			addExcelCell(row, 1, "", formatNumber(attr(rows[i], "TransAmount"), 2), "ExcelDisplayCell", "Right");
			addExcelCell(row, 1, "", formatNumber(attr(rows[i], "AmtAdjusted"), 2), "ExcelDisplayCell", "Right");
			addExcelCell(row, 1, "", formatNumber(attr(rows[i], "AmtToAccount"), 2), "ExcelDisplayCell", "Right");
			addExcelCell(row, 1, "", formatNumber(remaining, 2), "ExcelDisplayCell", "Right");
			addExcelCell(row, 2, "txtDocAmount" + attr(rows[i], "No") + "Z" + trim(attr(rows[i], "PayableNo")) + "Z" + (i + 1), formatNumber(attr(rows[i], "AmtToAdjust"), 2), "ExcelInputCell", "right", 10, 12, 'style="text-align:right"');
		}
	}

	function setPayableDisplay(flag) {
		setVisible("Disaddtional", Number(flag) !== 0, "115px");
		setVisible("DisPayable", Number(flag) !== 0, "110px");
	}

	function setAnalDisplay(display, flag) {
		setVisible(String(display) === "A" ? "DisAnal" : "DisCost", Number(flag) !== 0, "100px", "280px");
	}

	function setADDDisplay(flag) {
		setVisible("Disaddtional", Number(flag) !== 0, "115px");
		if (Number(flag) === 2) {
			return;
		}
		setVisible("DisCCANL", Number(flag) !== 0, "114px");
	}

	function getGlHeadXml(value) {
		var parts = String(value || "").split(":");
		var root = xmlRoot("AccHeadData");
		var node;
		if (!root) {
			return;
		}
		node = appendElement(root.ownerDocument, root, "AccHead");
		setAttr(node, "No", trim(parts[0]));
		setAttr(node, "CostCenter", trim(parts[2]));
		setAttr(node, "Analytical", trim(parts[3]));
		setAttr(node, "Name", trim(parts[5]));
		setAttr(node, "Type", "G");
		setAttr(node, "TransFlag", trim(parts[4]));
	}

	function getGlHeadXmlForSalAcc() {
		var root = xmlRoot("AccHeadData");
		var node;
		if (!root) {
			return;
		}
		node = appendElement(root.ownerDocument, root, "AccHead");
		setAttr(node, "No", valueOf("hSalAccCode"));
		setAttr(node, "CostCenter", "0");
		setAttr(node, "Analytical", "0");
		setAttr(node, "Name", valueOf("hSalAccName"));
		setAttr(node, "Type", "G");
		setAttr(node, "TransFlag", "A");
	}

	function getPartyHeadXml(code, name, value) {
		var parts = String(value || "").split(":");
		var root = xmlRoot("AccHeadData");
		var node;
		if (!root) {
			return;
		}
		node = appendElement(root.ownerDocument, root, "AccHead");
		setAttr(node, "No", trim(code));
		setAttr(node, "Pay", trim(parts[0]));
		setAttr(node, "Rec", trim(parts[1]));
		setAttr(node, "Name", name);
		setAttr(node, "Type", "P");
		setAttr(node, "Adv", trim(parts[2]));
	}

	function setApp(type) {
		var approver = firstField(field("selUserId"));
		if (approver) {
			approver.disabled = String(type) !== "Y";
		}
	}

	function checkApp() {
		var approvals = field("optApprove");
		var narration = valueOf("txtNarration");
		if (approvals && approvals.length && approvals[0].checked && firstField(field("selUserId")) && firstField(field("selUserId")).selectedIndex === 0) {
			alert("Select Approver ");
			firstField(field("selUserId")).focus();
			return false;
		}
		if (narration.length > 300) {
			alert("Narration Should be Less than 300 Characters ");
			return false;
		}
		if (field("selAccUnitId") && field("selUnitId") && String(valueOf("selAccUnitId")) !== String(valueOf("selUnitId"))) {
			alert("Created Unit and Accounting Unit is different!!");
			return false;
		}
		return true;
	}

	function voucherDate() {
		var dateField = firstField(field("ctlDate")) || byId("ctlDate");
		if (dateField && typeof dateField.GetDate === "function") {
			return dateField.GetDate();
		}
		if (dateField && typeof dateField.getDate === "function") {
			return dateField.getDate();
		}
		return dateField ? dateField.value : "";
	}

	function checkFinDate() {
		var from = trim(valueOf("hFinFrm"));
		var to = trim(valueOf("hFinTo"));
		var parts;
		var period;
		var message;
		if (!from || !to) {
			return true;
		}
		parts = trim(voucherDate()).split(/[\/.-]/);
		message = "Voucher Date Should Be Between 01/04/" + from.substring(0, 4) + " To 31/03/" + to.substring(0, 4);
		if (parts.length < 3) {
			alert(message);
			return false;
		}
		period = parts[0].length === 4 ? toNumber(parts[0] + parts[1]) : toNumber(parts[2] + parts[1]);
		if (period < toNumber(from) || period > toNumber(to)) {
			alert(message);
			return false;
		}
		return true;
	}

	function popPayRecAmd(headerNode) {
		popPayRec(headerNode);
		setPayableDisplayAmd(1);
	}

	function setPayableDisplayAmd(flag) {
		setVisible("Disaddtional", Number(flag) !== 0, "115px");
		setVisible("DisPayable", Number(flag) !== 0, "90px");
	}

	function chkEnter(event) {
		var evt = event || null;
		if (evt && evt.keyCode === 13) {
			if (evt.preventDefault) {
				evt.preventDefault();
			}
			return false;
		}
		return true;
	}

	window.SelectHead = selectHead;
	window.CheckAccHead = checkAccHead;
	window.popAddAmount = distributeAmount;
	window.clearXML = clearXml;
	window.ValidateAmount = validateAmount;
	window.showNarration = showNarration;
	window.popCostCenter = popCostCenter;
	window.popAnalytical = popAnalytical;
	window.showCCAnal = showCCAnal;
	window.popPayRec = popPayRec;
	window.setPayableDisplay = setPayableDisplay;
	window.setAnalDisplay = setAnalDisplay;
	window.setADDDisplay = setADDDisplay;
	window.GetGlHeadXml = getGlHeadXml;
	window.GetGlHeadXmlForSalAcc = getGlHeadXmlForSalAcc;
	window.GetPartyHeadXml = getPartyHeadXml;
	window.SetApp = setApp;
	window.CheckApp = checkApp;
	window.CheckFinDate = checkFinDate;
	window.popPayRecAmd = popPayRecAmd;
	window.setPayableDisplayAmd = setPayableDisplayAmd;
	window.ChkEnter = chkEnter;
}(window, document));
