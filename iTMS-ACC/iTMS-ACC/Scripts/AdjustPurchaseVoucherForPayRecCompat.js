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
			for (var index = 0; index < result.snapshotLength; index += 1) {
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

	function createPayRecDoc(doc) {
		return doc.createElement("Doc");
	}

	function appendSelectedDocument(payRecRoot, checkbox) {
		var accDoc = xmlDocument("AccHeadData");
		var docNo = checkbox.value;
		var docParts = String(docNo || "").split("Z");
		var source = valueOf("hDoc" + docNo).split("?");
		var node;
		if (!accDoc) {
			return;
		}
		node = createPayRecDoc(accDoc);
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
		setAttr(node, "CheckAdjust", docParts[3]);
		setAttr(node, "Check", "SEL");
		setAttr(node, "CrTrNo", source[7]);
		payRecRoot.appendChild(node);
	}

	function alreadyInPayRec(payRecRoot, advance) {
		var children = childElements(payRecRoot, "Doc");
		var createdTransNo = attr(advance, "CreatedTransNo");
		var advNo = trim(attr(advance, "AdvNo"));
		return children.some(function (node) {
			return attr(node, "CrTrNo") === createdTransNo && trim(attr(node, "PayableNo")) === advNo;
		});
	}

	function appendAdvanceDocument(payRecRoot, advance, voucherType) {
		var accDoc = xmlDocument("AccHeadData");
		var node;
		var advanceAdjType = trim(attr(advance, "AdjType"));
		if (!accDoc || alreadyInPayRec(payRecRoot, advance)) {
			return;
		}
		node = createPayRecDoc(accDoc);
		setAttr(node, "No", attr(advance, "TransNo"));
		setAttr(node, "InvNo", attr(advance, "VoucherNo"));
		setAttr(node, "InvDate", attr(advance, "VoucherDate"));
		setAttr(node, "TransAmount", formatNumber(attr(advance, "AmountRec")));
		setAttr(node, "AmtAdjusted", formatNumber(attr(advance, "AmountAdj")));
		setAttr(node, "AmtToAdjust", formatNumber(attr(advance, "AmountToAdj")));
		setAttr(node, "DocType", valueOf("hVouType"));
		setAttr(node, "AmtToAccount", formatNumber(attr(advance, "ToAccount")));
		setAttr(node, "PayableNo", trim(attr(advance, "AdvNo")));
		if (trim(voucherType) === "PAY") {
			setAttr(node, "AdjType", advanceAdjType === "D" ? advanceAdjType : "P");
		} else {
			setAttr(node, "AdjType", advanceAdjType);
		}
		if (attr(advance, "Check") === "DEL") {
			setAttr(node, "Check", "SEL");
			setAttr(node, "CheckAdjust", "NN");
		} else {
			setAttr(node, "Check", "");
		}
		setAttr(node, "CrTrNo", attr(advance, "CreatedTransNo"));
		payRecRoot.appendChild(node);
	}

	function appendPendingAdvanceDocuments(payRecRoot, voucherType) {
		var adjRoot = xmlRoot("AdjDet");
		var advances = selectNodes(adjRoot, "//AdvanceDetails/Advance");
		advances.forEach(function (advance) {
			var check = attr(advance, "Check");
			if (check === "" || check === "DEL") {
				appendAdvanceDocument(payRecRoot, advance, voucherType);
			}
		});
	}

	function appendParty(payRecRoot) {
		var accDoc = xmlDocument("AccHeadData");
		var adjRoot = xmlRoot("AdjDet");
		var party = firstNode(adjRoot, "//Header/Party");
		if (accDoc && party) {
			payRecRoot.appendChild(importNodeFor(accDoc, party));
		}
	}

	function finaldone() {
		var frm = form();
		var accRoot;
		var payRecRoot;
		var recCountNode;
		var voucherType = trim(valueOf("hCallFrom"));
		var selectedDocuments;
		var adjustedAmount;
		if (!frm) {
			return false;
		}
		ensureCompat();
		if (String(valueOf("hCrTrNo")) !== "0") {
			adjustedAmount = toNumber(valueOf("TxtTotToAdjustedAmount")) + toNumber(valueOf("TxtTotDebitNoteAdj")) + toNumber(valueOf("TxtTotAdjustedAmount"));
			if (toNumber(textOf("spVouAmt")) < adjustedAmount) {
				alert("Adjusted Amount is greater than the VoucherAmount,Adjustments Cannot be Done ");
				return false;
			}
		} else {
			if (toNumber(textOf("spDRAmt")) > 0 || toNumber(textOf("spCRAmt")) > 0) {
				if (toNumber(textOf("spDiffAmt")) !== 0) {
					alert("Both Debit and Credit Amount Should be Equal ");
					return false;
				}
			} else {
				alert("Select Any Bills ");
				return false;
			}
			setValue("hDrCrAmt", textOf("spDRAmt"));
		}

		accRoot = xmlRoot("AccHeadData");
		if (!accRoot) {
			alert("Adjustment data is not available.");
			return false;
		}
		payRecRoot = childElements(accRoot, "PayRec")[0];
		if (!payRecRoot) {
			payRecRoot = xmlDocument("AccHeadData").createElement("PayRec");
			accRoot.insertBefore(payRecRoot, accRoot.firstChild);
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

		appendPendingAdvanceDocuments(payRecRoot, voucherType);
		appendParty(payRecRoot);

		if (selectNodes(xmlDocument("AccHeadData"), "//PayRec/Doc").length === 0) {
			alert("No Adujstment can done");
			return false;
		}

		syncPost("XMLSave.asp?Name=Bill Closing&Mod=Det", serializeXml("AccHeadData"));
		frm.action = "AdjustVouWithBillUpdateNew.asp";
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

	function DisplayAmt111(sPassValue) {
		var boxes = checkboxesNamed("chkDocument");
		var index;
		var checkbox;
		var value;
		var docValue;
		var source;
		var amount;
		var transAmount;
		for (index = 0; index < boxes.length; index += 1) {
			checkbox = boxes[index];
			value = checkbox.value;
			if (trim(sPassValue) !== trim(value)) {
				continue;
			}
			docValue = valueOf("hDoc" + value);
			transAmount = toNumber(valueOf("txtToAdk" + value));
			amount = toNumber(valueOf("TxtToAdjust" + value));
			source = docValue.split("?");
			if (String(valueOf("hCrTrNo")) !== "0") {
				if (checkbox.checked && transAmount < amount) {
					alert("Enter the amount which is less than or equal to bill Amount");
					return false;
				}
				updateRunningTotal(source[6], amount, checkbox.checked);
			} else {
				updateDrCrTotal(source[6], amount, checkbox.checked);
				updateDrCrSpans();
			}
			break;
		}
		return true;
	}

	function DisplayAmt() {
		var boxes = checkboxesNamed("chkDocument");
		var debitNoteTotal = 0;
		var adjustmentTotal = 0;
		var index;
		var value;
		var transAmount;
		var amount;
		var checkType;
		if (String(valueOf("hCrTrNo")) === "0") {
			return true;
		}
		for (index = 0; index < boxes.length; index += 1) {
			if (!boxes[index].checked) {
				continue;
			}
			value = boxes[index].value;
			transAmount = toNumber(valueOf("txtToAdk" + value));
			amount = toNumber(valueOf("TxtToAdjust" + value));
			checkType = valueOf("hCheckType" + value);
			if (transAmount >= amount) {
				if (checkType === "D" || checkType === "C") {
					debitNoteTotal += amount;
				} else {
					adjustmentTotal += amount;
				}
			} else {
				alert("Enter the amount which is less than or equal to bill Amount");
				setValue("TxtToAdjust" + value, "0.00");
				if (checkType === "D" || checkType === "C") {
					setValue("TxtTotDebitNoteAdj", "0.00");
				} else {
					setValue("TxtTotToAdjustedAmount", "0.00");
				}
				return false;
			}
		}
		setValue("TxtTotDebitNoteAdj", formatNumber(debitNoteTotal));
		setValue("TxtTotToAdjustedAmount", formatNumber(adjustmentTotal));
		return true;
	}

	function UpdateAmt(sourceCheckbox) {
		var value = sourceCheckbox && sourceCheckbox.value || "";
		var docValue = valueOf("hDoc" + value);
		var amountField = field("TxtToAdjust" + value);
		var amount;
		var parts;
		if (!value) {
			return false;
		}
		parts = docValue.split("?");
		if (String(valueOf("hCrTrNo")) !== "0") {
			if (amountField) {
				amountField.readOnly = false;
				amountField.removeAttribute("readonly");
			}
			amount = toNumber(valueOf("TxtToAdjust" + value));
			updateRunningTotal(parts[6], amount, !!sourceCheckbox.checked);
			DisplayAmt(value);
		} else {
			amount = toNumber(valueOf("txtToAdk" + value));
			updateDrCrTotal(parts[6], amount, !!sourceCheckbox.checked);
			updateDrCrSpans();
		}
		return true;
	}

	function SetTotalVal(iAdvNo) {
		var adjRoot;
		var advances;
		if (String(iAdvNo) === "0") {
			dTotal = 0;
		}
		ensureCompat();
		if (valueOf("hType") === "") {
			adjRoot = xmlRoot("AdjDet");
			advances = selectNodes(adjRoot, "//AdvanceDetails/Advance");
			advances.forEach(function (advance) {
				setAttr(advance, "Check", "");
			});
		}
		return true;
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
			var selectedInvNo;
			var selectedTransNo;
			var parts;
			if (!checkbox.checked) {
				return;
			}
			selectedCount += 1;
			checkbox.disabled = false;
			parts = String(checkbox.value || "").split(":");
			selectedInvNo = trim(parts[0]);
			selectedTransNo = trim(parts[2]);
			selectNodes(adjRoot, "//AdvanceDetails/Advance").forEach(function (advance) {
				if (selectedInvNo === trim(attr(advance, "VoucherNo")) && selectedTransNo === trim(attr(advance, "TransNo"))) {
					setAttr(advance, "Check", "DEL");
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
		var url = "PurchaseVouchView_San.asp?TransNo=" + encodeURIComponent(iCrTransNo);
		var features = "dialogHeight:490px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, function () {});
		} else {
			window.open(url, "_blank", "height=490,width=670,resizable=no,status=no,scrollbars=yes");
		}
		return false;
	}

	function document_onkeypress(evt) {
		var eventObject = evt || window.event;
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
	window.DisplayAmt111 = DisplayAmt111;
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
