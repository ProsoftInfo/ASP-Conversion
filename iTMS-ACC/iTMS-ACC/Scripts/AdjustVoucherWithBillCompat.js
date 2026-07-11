(function (window, document) {
	"use strict";

	var dTotal = 0;
	var dDrTotal = 0;
	var dCrTotal = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.frm1 || document.forms["frm1"] || document.frm1 || document.forms[0] || null;
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
			if (String(elements[index].name || "").toLowerCase() === target) {
				return elements[index];
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

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
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

	function withItemAliases(nodes) {
		nodes.Item = function (index) {
			return this[index];
		};
		nodes.item = nodes.Item;
		return nodes;
	}

	function selectNodes(context, expression) {
		var doc;
		var result;
		var nodes = [];
		var index;
		if (!context) {
			return withItemAliases(nodes);
		}
		if (typeof context.selectNodes === "function") {
			return withItemAliases(Array.prototype.slice.call(context.selectNodes(expression)));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return withItemAliases(nodes);
		}
		try {
			result = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
			for (index = 0; index < result.snapshotLength; index += 1) {
				nodes.push(result.snapshotItem(index));
			}
		} catch (ignore) {}
		return withItemAliases(nodes);
	}

	function firstNode(context, expression) {
		var nodes = selectNodes(context, expression);
		return nodes.length ? nodes[0] : null;
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

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function importNodeFor(doc, node) {
		return node && node.ownerDocument !== doc && doc.importNode ? doc.importNode(node, true) : node;
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
		return xhr.responseText || "";
	}

	function checkboxesNamed(namePart) {
		var frm = form();
		var matches = [];
		var elements;
		var index;
		if (!frm || !frm.elements) {
			return matches;
		}
		elements = frm.elements;
		for (index = 0; index < elements.length; index += 1) {
			if (String(elements[index].type || "").toLowerCase() === "checkbox" && String(elements[index].name || "").indexOf(namePart) !== -1) {
				matches.push(elements[index]);
			}
		}
		return matches;
	}

	function updateDrCrSpans() {
		var partyType = trim(textOf("spParTy"));
		var diffAmount = partyType === "DR" ? dDrTotal - dCrTotal : dCrTotal - dDrTotal;
		setText("spDRAmt", formatNumber(dDrTotal));
		setText("spCRAmt", formatNumber(dCrTotal));
		setText("spDiffAmt", formatNumber(diffAmount));
	}

	function updateRunningTotal(adjType, amount, checked) {
		var sign = checked ? 1 : -1;
		switch (String(adjType)) {
			case "I":
			case "D":
			case "P":
				dTotal -= sign * toNumber(amount);
				break;
			case "PI":
			case "C":
				dTotal += sign * toNumber(amount);
				break;
			default:
				break;
		}
	}

	function updateDrCrTotal(adjType, amount, checked) {
		var delta = (checked ? 1 : -1) * toNumber(amount);
		switch (String(adjType)) {
			case "I":
			case "D":
			case "R":
				dDrTotal += delta;
				break;
			case "PI":
			case "C":
			case "P":
				dCrTotal += delta;
				break;
			default:
				break;
		}
	}

	function ensurePayRecRoot(accRoot) {
		var accDoc = xmlDocument("AccHeadData");
		var payRecRoot = childElements(accRoot, "PayRec")[0];
		if (!payRecRoot && accDoc) {
			payRecRoot = accDoc.createElement("PayRec");
			accRoot.insertBefore(payRecRoot, accRoot.firstChild);
		}
		return payRecRoot;
	}

	function appendSelectedDocument(payRecRoot, checkbox) {
		var accDoc = xmlDocument("AccHeadData");
		var docNo = checkbox.value;
		var docParts = String(docNo || "").split("Z");
		var source = valueOf("hDoc" + docNo).split("?");
		var node;
		if (!accDoc || !payRecRoot) {
			return;
		}
		node = accDoc.createElement("Doc");
		setAttr(node, "No", docParts[0]);
		setAttr(node, "InvNo", trim(source[0]));
		setAttr(node, "InvDate", trim(source[1]));
		setAttr(node, "TransAmount", formatNumber(source[2]));
		setAttr(node, "AmtAdjusted", formatNumber(source[3]));
		setAttr(node, "AmtToAdjust", valueOf("TxtToAdjust" + docNo, "0.00"));
		setAttr(node, "DocType", valueOf("hVouType"));
		setAttr(node, "AmtToAccount", formatNumber(source[4]));
		setAttr(node, "PayableNo", trim(source[5]));
		setAttr(node, "AdjType", trim(source[6]));
		setAttr(node, "CheckAdjust", docParts[4] || "");
		setAttr(node, "Check", "SEL");
		payRecRoot.appendChild(node);
	}

	function appendExistingAdjustmentDocs(payRecRoot) {
		var accDoc = xmlDocument("AccHeadData");
		var adjRoot = xmlRoot("AdjDet");
		if (!accDoc || !adjRoot) {
			return;
		}
		selectNodes(adjRoot, "//voucher/Entry/PayRec/Doc").forEach(function (docNode) {
			if (attr(docNode, "Check") === "") {
				payRecRoot.appendChild(importNodeFor(accDoc, docNode));
			}
		});
	}

	function appendVoucherAccountHead(payRecRoot) {
		var accDoc = xmlDocument("AccHeadData");
		var adjRoot = xmlRoot("AdjDet");
		var accHead = firstNode(adjRoot, "//voucher/Entry/AccHead");
		if (accDoc && accHead) {
			payRecRoot.appendChild(importNodeFor(accDoc, accHead));
		}
	}

	function selectedAdjustmentAmount() {
		var total = 0;
		checkboxesNamed("chkDocument").forEach(function (checkbox) {
			var value;
			var transAmount;
			var amount;
			if (!checkbox.checked) {
				return;
			}
			value = checkbox.value;
			transAmount = toNumber(valueOf("txtToAdk" + value));
			amount = toNumber(valueOf("TxtToAdjust" + value));
			if (transAmount >= amount) {
				total += amount;
			}
		});
		return total;
	}

	function finaldone() {
		var frm = form();
		var accRoot;
		var payRecRoot;
		var recCountNode;
		var selectedDocuments;
		var adjustedAmount;
		if (!frm) {
			return false;
		}
		ensureCompat();
		if (String(valueOf("hCrTrNo")) !== "0") {
			adjustedAmount = toNumber(valueOf("TxtTotToAdjustedAmount")) + toNumber(valueOf("TxtTotAdjustedAmount"));
			if (toNumber(textOf("spVouAmt")) < adjustedAmount) {
				alert("Adjusted Amount is greater than the VoucherAmount, Adjustments Cannot be Done ");
				return false;
			}
		} else if (toNumber(textOf("spDRAmt")) > 0 || toNumber(textOf("spCRAmt")) > 0) {
			if (toNumber(textOf("spDiffAmt")) !== 0) {
				alert("Both Debit and Credit Amount Should be Equal ");
				return false;
			}
			setValue("hDrCrAmt", textOf("spDRAmt"));
		} else {
			alert("Select Any Bills ");
			return false;
		}

		accRoot = xmlRoot("AccHeadData");
		if (!accRoot) {
			alert("Adjustment data is not available.");
			return false;
		}
		payRecRoot = ensurePayRecRoot(accRoot);
		if (!payRecRoot) {
			alert("Adjustment data is not available.");
			return false;
		}
		clearChildren(payRecRoot);

		selectedDocuments = checkboxesNamed("chkDocument").filter(function (checkbox) {
			return checkbox.checked;
		});
		selectedDocuments.forEach(function (checkbox) {
			appendSelectedDocument(payRecRoot, checkbox);
		});

		setAttr(accRoot, "No", selectedDocuments.length ? "1" : "0");
		recCountNode = childElements(accRoot, "RecCount")[0] || firstNode(accRoot, "//RecCount");
		if (recCountNode) {
			setAttr(recCountNode, "Val", valueOf("hRecCount"));
		}

		appendExistingAdjustmentDocs(payRecRoot);
		appendVoucherAccountHead(payRecRoot);

		syncPost("XMLSave.asp?Name=Bill Closing&Mod=Det", serializeXml("AccHeadData"));
		syncPost("XMLSave.asp?Name=Adjust&Mod=Det", serializeXml("AccHeadData"));
		frm.action = "AdjustVouWithBillUpdate.asp";
		frm.submit();
		return true;
	}

	function finalcancel() {
		var frm = form();
		if (!frm) {
			return false;
		}
		frm.action = "MANAGEBANKVOUCHERS.ASP";
		frm.submit();
		return true;
	}

	function DisplayAmt(sPassValue) {
		var boxes = checkboxesNamed("chkDocument");
		var value;
		var transAmount;
		var amount;
		var docValue;
		var source;
		var index;
		var totalAmount;

		if (String(valueOf("hCrTrNo")) !== "0") {
			totalAmount = selectedAdjustmentAmount();
			for (index = 0; index < boxes.length; index += 1) {
				if (!boxes[index].checked) {
					continue;
				}
				value = boxes[index].value;
				transAmount = toNumber(valueOf("txtToAdk" + value));
				amount = toNumber(valueOf("TxtToAdjust" + value));
				if (transAmount < amount) {
					alert("Enter the amount which is less than or equal to bill Amount");
					setValue("TxtToAdjust" + value, "0.00");
					setValue("TxtTotToAdjustedAmount", "0.00");
					return false;
				}
			}
			setValue("TxtTotToAdjustedAmount", formatNumber(totalAmount));
		}

		for (index = 0; index < boxes.length; index += 1) {
			value = boxes[index].value;
			if (trim(sPassValue) !== trim(value)) {
				continue;
			}
			docValue = valueOf("hDoc" + value);
			source = docValue.split("?");
			amount = toNumber(valueOf("TxtToAdjust" + value));
			transAmount = toNumber(valueOf("txtToAdk" + value));
			if (String(valueOf("hCrTrNo")) !== "0") {
				if (boxes[index].checked && transAmount < amount) {
					alert("Enter the amount which is less than or equal to bill Amount");
					return false;
				}
				updateRunningTotal(source[6], amount, boxes[index].checked);
			} else {
				updateDrCrTotal(source[6], amount, boxes[index].checked);
				updateDrCrSpans();
			}
			break;
		}
		return true;
	}

	function UpdateAmt(sourceCheckbox) {
		var value = sourceCheckbox && sourceCheckbox.value || "";
		var docValue = valueOf("hDoc" + value);
		var amountField = field("TxtToAdjust" + value);
		var parts = docValue.split("?");
		var amount;
		if (!value) {
			return false;
		}
		if (String(valueOf("hCrTrNo")) !== "0") {
			if (amountField) {
				amountField.readOnly = false;
				amountField.removeAttribute("readonly");
			}
			amount = toNumber(valueOf("TxtToAdjust" + value));
			updateRunningTotal(parts[6], amount, !!sourceCheckbox.checked);
		} else {
			amount = toNumber(valueOf("txtToAdk" + value));
			updateDrCrTotal(parts[6], amount, !!sourceCheckbox.checked);
			updateDrCrSpans();
		}
		return true;
	}

	function SetTotalVal(iAdvNo) {
		var adjRoot;
		if (String(iAdvNo) === "0") {
			dTotal = 0;
		}
		ensureCompat();
		if (valueOf("hType") === "") {
			adjRoot = xmlRoot("AdjDet");
			selectNodes(adjRoot, "//voucher/Entry/PayRec/Doc").forEach(function (docNode) {
				setAttr(docNode, "Check", "");
			});
		}
		return true;
	}

	function selectedInvoiceNumber(value) {
		return trim(String(value || "").split(":")[0]);
	}

	function documentInvoiceNumber(docNode) {
		var invoiceNumber = attr(docNode, "InvNo");
		var parts;
		var marker;
		if (invoiceNumber.indexOf(":") !== -1) {
			parts = invoiceNumber.split(":");
			invoiceNumber = parts.length > 1 ? parts[1] : parts[0];
		}
		marker = invoiceNumber.indexOf("Dt");
		if (marker !== -1) {
			invoiceNumber = invoiceNumber.substring(0, marker);
		}
		return trim(invoiceNumber);
	}

	function DeleteData() {
		var adjRoot;
		var selectedCount = 0;
		if (!confirm("Do You Want To Delete Bill")) {
			return false;
		}
		ensureCompat();
		adjRoot = xmlRoot("AdjDet");
		checkboxesNamed("ChkBillDet").forEach(function (checkbox) {
			var invNo;
			if (!checkbox.checked) {
				return;
			}
			selectedCount += 1;
			checkbox.disabled = false;
			invNo = selectedInvoiceNumber(checkbox.value);
			selectNodes(adjRoot, "//voucher/Entry/PayRec/Doc").forEach(function (docNode) {
				if (invNo === documentInvoiceNumber(docNode)) {
					setAttr(docNode, "Check", "DEL");
				}
			});
		});
		if (selectedCount === 0) {
			alert("Select any one record for Delete");
			return false;
		}
		syncPost("XMLSave.asp?Name=Adjust&Mod=Det", serializeXml("AdjDet"));
		setValue("hType", "A");
		form().submit();
		return true;
	}

	function ShowVouch(iCrTransNo) {
		var url = "BankVouchView_San.asp?TransNo=" + encodeURIComponent(iCrTransNo);
		var features = "dialogHeight:450px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, function () {});
		} else {
			window.open(url, "_blank", "height=450,width=670,resizable=no,status=no,scrollbars=yes");
		}
		return false;
	}

	function document_onkeypress(evt) {
		var eventObject = evt || null;
		if (eventObject && (eventObject.key === "Escape" || eventObject.keyCode === 27)) {
			if (eventObject.preventDefault) {
				eventObject.preventDefault();
			}
			finalcancel();
			return false;
		}
		return true;
	}

	document.addEventListener("keydown", document_onkeypress);

	window.finaldone = finaldone;
	window.finalcancel = finalcancel;
	window.DisplayAmt = DisplayAmt;
	window.UpdateAmt = UpdateAmt;
	window.SetTotalVal = SetTotalVal;
	window.DeleteData = DeleteData;
	window.ShowVouch = ShowVouch;
	window.document_onkeypress = document_onkeypress;
	window.window_onunload = function () {
		return true;
	};
}(window, document));
