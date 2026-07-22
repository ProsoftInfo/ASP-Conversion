(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootNode = null;
	var classCode = "";
	var itemCode = "";
	var entryNo = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function toInt(value) {
		var parsed = parseInt(String(value == null ? "" : value).replace(/,/g, ""), 10);
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function focusOrSelect(control) {
		if (!control) {
			return;
		}
		if (typeof control.select === "function") {
			control.select();
		} else if (typeof control.focus === "function") {
			control.focus();
		}
	}

	function xmlDocument(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function outDataElement() {
		ensureCompat();
		return document.getElementById("OutData") || window.OutData || null;
	}

	function outDataDocument() {
		return xmlDocument(outDataElement());
	}

	function outDataRoot() {
		return xmlRoot(outDataElement());
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function importForDocument(doc, node) {
		if (!doc || !node) {
			return null;
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function findParentItem() {
		var items = elementChildren(rootNode, "ITEM");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "CLACODE") === trim(classCode) && getAttr(items[i], "ITMCODE") === trim(itemCode) && getAttr(items[i], "ENTRYNO") === trim(entryNo)) {
				return items[i];
			}
		}
		return null;
	}

	function removePickChildren(itemNode) {
		var picks = elementChildren(itemNode, "Pick");
		for (var i = 0; i < picks.length; i += 1) {
			itemNode.removeChild(picks[i]);
		}
	}

	function findPick(root, storeCode, binCode, lotNo) {
		var picks = elementChildren(root, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (getAttr(picks[i], "LOTNO") === trim(lotNo) && getAttr(picks[i], "LOC") === trim(storeCode) && getAttr(picks[i], "BIN") === trim(binCode)) {
				return picks[i];
			}
		}
		return null;
	}

	function rowSuffix(name, prefix) {
		return String(name || "").indexOf(prefix) === 0 ? String(name).substring(prefix.length) : "";
	}

	function issueRows() {
		var frm = form();
		var rows = [];
		var elements = frm && frm.elements || [];
		var suffix;
		for (var i = 0; i < elements.length; i += 1) {
			if (elements[i].type === "text" && String(elements[i].name || "").indexOf("txtQty") === 0) {
				suffix = rowSuffix(elements[i].name, "txtQty");
				if (field("txtIss" + suffix)) {
					rows.push({
						index: suffix,
						stock: elements[i],
						issue: field("txtIss" + suffix),
						pack: field("txtTotPackZ" + suffix)
					});
				}
			}
		}
		return rows;
	}

	function pendingQuantity() {
		var quantityElement = document.getElementById("idQty");
		var text = trim(quantityElement ? quantityElement.textContent || quantityElement.innerText : "");
		return text === "" ? null : toNumber(text);
	}

	function setReturnValue() {
		if (!rootNode) {
			rootNode = xmlRoot(objTemp);
		}
		if (!rootNode) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(rootNode);
		} else {
			window["return" + "Value"] = rootNode;
			window.returnvalue = rootNode;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function showTooltip(text) {
		var tooltip = document.getElementById("itms-pick-tooltip");
		var parts = String(text || "").split("|");
		var eventObj = window.event || null;
		if (!tooltip) {
			tooltip = document.createElement("div");
			tooltip.id = "itms-pick-tooltip";
			tooltip.style.cssText = "position:absolute;z-index:9999;display:none;max-width:260px;padding:6px 8px;border:1px solid #777;background:#ffffe1;color:#111;font:11px Verdana,Arial,sans-serif;white-space:pre-line;";
			document.body.appendChild(tooltip);
		}
		tooltip.textContent = "Passed For Count From : " + (parts[0] || "") + "\nPassed For Count To      : " + (parts[1] || "");
		tooltip.style.left = ((eventObj && eventObj.clientX || 10) + 12 + (window.pageXOffset || document.documentElement.scrollLeft || 0)) + "px";
		tooltip.style.top = ((eventObj && eventObj.clientY || 10) + 12 + (window.pageYOffset || document.documentElement.scrollTop || 0)) + "px";
		tooltip.style.display = "block";
	}

	function hideTooltip() {
		var tooltip = document.getElementById("itms-pick-tooltip");
		if (tooltip) {
			tooltip.style.display = "none";
		}
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function paginationParts(root, storeCode, binCode, lotNo) {
		var pick = findPick(root, storeCode, binCode, lotNo);
		var pagination = elementChildren(pick, "Pagination")[0];
		var value = pagination && pagination.attributes && pagination.attributes.item(0) ? pagination.attributes.item(0).nodeValue : "";
		if (!value) {
			return null;
		}
		return {
			raw: value,
			parts: String(value).split("```")
		};
	}

	function updateLotRow(lineNo, storeCode, binCode, lotNo) {
		var root = outDataRoot();
		var pick = findPick(root, storeCode, binCode, lotNo);
		var qtyField = field("txtIss" + lineNo);
		var packField = field("txtTotPackZ" + lineNo);
		if (!pick) {
			return;
		}
		if (qtyField) {
			qtyField.value = getAttr(pick, "QTYISS");
		}
		if (packField) {
			packField.value = getAttr(root, "NoofPack");
		}
	}

	function continueLotDialog(lineNo, storeCode, binCode, lotNo, root) {
		var page = paginationParts(root, storeCode, binCode, lotNo);
		if (page && page.parts.length === 1) {
			openModal(
				"mrsPickDetailLotPop.asp?sTemp=" + page.raw.replace(/\|/g, "&"),
				outDataDocument(),
				"dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No",
				function () {
					continueLotDialog(lineNo, storeCode, binCode, lotNo, outDataRoot());
				}
			);
			return;
		}
		if (page && page.parts.length > 1 && trim(page.parts[1]).substring(0, 2) === "NO") {
			return;
		}
		updateLotRow(lineNo, storeCode, binCode, lotNo);
	}

	window.ShowDet = function (desc) {
		if (desc === "-") {
			hideTooltip();
			return;
		}
		showTooltip(desc);
	};

	window.fnInit = function (sItem, sClass, sEntNo) {
		var outRoot;
		var parentItem;
		var picks;
		var incomingPicks;
		var issueField;
		var packField;
		var i;

		itemCode = sItem;
		classCode = sClass;
		entryNo = sEntNo;
		objTemp = modalArgs();
		rootNode = xmlRoot(objTemp);
		outRoot = outDataRoot();
		parentItem = findParentItem();

		if (!parentItem || !outRoot) {
			return;
		}

		picks = elementChildren(parentItem, "Pick");
		for (i = 0; i < picks.length; i += 1) {
			setAttr(outRoot, "TOT", getAttr(picks[i], "TOT"));
			incomingPicks = elementChildren(picks[i], "PICK");
			for (var p = 0; p < incomingPicks.length; p += 1) {
				if (getAttr(incomingPicks[p], "QTYISS") !== "") {
					issueField = field("txtIss" + (p + 1));
					packField = field("txtTotPackZ" + (p + 1));
					if (issueField) {
						issueField.value = getAttr(incomingPicks[p], "QTYISS");
					}
					if (packField) {
						packField.value = getAttr(incomingPicks[p], "NoofPack");
					}
				}
				replaceOutDataPick(outRoot, incomingPicks[p]);
			}
			parentItem.removeChild(picks[i]);
		}
	};

	function replaceOutDataPick(outRoot, incomingPick) {
		var existing = findPick(outRoot, getAttr(incomingPick, "LOC"), getAttr(incomingPick, "BIN"), getAttr(incomingPick, "LOTNO"));
		var outDoc = outDataDocument();
		if (existing) {
			outRoot.removeChild(existing);
		}
		outRoot.appendChild(importForDocument(outDoc, incomingPick));
	}

	window.CheckSubmit = function () {
		var outRoot = outDataRoot();
		var outDoc = outDataDocument();
		var rows = issueRows();
		var totalQty = 0;
		var totalPacks = 0;
		var parentItem;
		var picks;
		var row;
		var pick;
		var pending;
		var i;

		if (!outRoot) {
			return;
		}

		for (i = 0; i < rows.length; i += 1) {
			row = rows[i];
			if (trim(row.issue.value) === "") {
				alert("Enter Quantity Issue");
				focusOrSelect(row.issue);
				return;
			}
			if (!checkNumbers(row.issue.value)) {
				alert("Enter Numerals Only");
				focusOrSelect(row.issue);
				return;
			}
			if (toNumber(row.issue.value) > toNumber(row.stock.value)) {
				alert("Quantity Issue should be equal to or less than Stock Quantity " + row.stock.value);
				focusOrSelect(row.issue);
				return;
			}
			totalQty += toNumber(row.issue.value);
		}

		pending = pendingQuantity();
		if (pending != null && totalQty > pending) {
			alert("Quantity Issue should be equal to or less than Quantity Pending " + pending);
			return;
		}

		parentItem = findParentItem();
		if (!parentItem) {
			return;
		}

		totalPacks = 0;
		picks = elementChildren(outRoot, "PICK");
		for (i = 0; i < picks.length; i += 1) {
			row = rows[i];
			if (!row) {
				break;
			}
			setAttr(picks[i], "QTYISS", trim(row.issue.value));
			setAttr(picks[i], "NoofPack", row.pack ? row.pack.value : "");
			totalPacks += toInt(row.pack && row.pack.value);
		}

		setAttr(outRoot, "TOT", totalQty !== 0 ? totalQty : "0");
		setAttr(outRoot, "NoofPack", totalPacks);
		removePickChildren(parentItem);

		if (totalQty !== 0) {
			setAttr(parentItem, "ISSQTY", "");
			parentItem.appendChild(importForDocument(xmlDocument(objTemp), outRoot));
		}

		closeWithReturn();
	};

	window.CheckLot = function (obj, entNo, classValue, itemValue, org, optName, invRecNo) {
		var parts = String(obj && obj.name || "").split(":");
		var line = parts[1] || "";
		var lotNo = parts[2] || "";
		var storeCode = parts[3] || "";
		var binCode = parts[4] || "";
		var tempValues = String(obj && obj.name || "") + ":" + classValue + ":" + itemValue + ":" + org + ":" + invRecNo + ":" + entNo + ":" + optName;
		var root = outDataRoot();

		openModal(
			"mrsPickDetailLotPop.asp?sTemp=" + tempValues,
			outDataDocument(),
			"dialogHeight:490px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No",
			function () {
				continueLotDialog(line, storeCode, binCode, lotNo, root || outDataRoot());
			}
		);
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
