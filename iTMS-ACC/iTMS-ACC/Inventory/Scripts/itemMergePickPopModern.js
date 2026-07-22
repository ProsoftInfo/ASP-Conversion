(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var outRoot = null;
	var outDoc = null;
	var classCode = "";
	var itemCode = "";
	var entryNo = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var value = frm && frm.elements ? frm.elements[name] : null;
		if (value && value.length && !value.tagName) {
			return value[0] || null;
		}
		return value || null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot === "0" || lot.toUpperCase() === "NULL" ? "N/A" : lot;
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
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

	function importNode(targetDoc, node) {
		if (!targetDoc || !node) {
			return node ? node.cloneNode(true) : null;
		}
		return targetDoc.importNode ? targetDoc.importNode(node, true) : node.cloneNode(true);
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function sameItem(node) {
		return String(node && node.nodeName).toLowerCase() === "item" &&
			getAttr(node, "CLASSCODE") === classCode &&
			getAttr(node, "ITEMCODE") === itemCode &&
			getAttr(node, "ENTRYNO") === entryNo;
	}

	function selectedItemNode() {
		var items = elementChildren(root, "Item");
		for (var i = 0; i < items.length; i += 1) {
			if (sameItem(items[i])) {
				return items[i];
			}
		}
		return null;
	}

	function removePickChildren(item) {
		elementChildren(item, "Pick").forEach(function (pick) {
			item.removeChild(pick);
		});
	}

	function matchingOutPick(sourcePick) {
		var picks = elementChildren(outRoot, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (normalizeLot(getAttr(picks[i], "LOTNO")) === normalizeLot(getAttr(sourcePick, "LOTNO")) &&
					getAttr(picks[i], "LOC") === getAttr(sourcePick, "LOC") &&
					getAttr(picks[i], "BIN") === getAttr(sourcePick, "BIN")) {
				return picks[i];
			}
		}
		return null;
	}

	function replaceOutPick(sourcePick) {
		var target = matchingOutPick(sourcePick);
		var imported = importNode(outDoc, sourcePick);
		if (!outRoot || !imported) {
			return;
		}
		if (target) {
			outRoot.replaceChild(imported, target);
		} else {
			outRoot.appendChild(imported);
		}
	}

	function loadExistingSelection() {
		var item = selectedItemNode();
		var pick;
		var picks;
		var line = 0;
		if (!item) {
			return;
		}
		pick = elementChildren(item, "Pick")[0];
		if (!pick) {
			return;
		}
		setAttr(outRoot, "TOT", getAttr(pick, "TOT"));
		picks = elementChildren(pick, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (getAttr(picks[i], "QTYISS") !== "") {
				line += 1;
				if (field("txtIss" + line)) {
					field("txtIss" + line).value = getAttr(picks[i], "QTYISS");
				}
				if (field("txtTotPackZ" + line)) {
					field("txtTotPackZ" + line).value = getAttr(picks[i], "NoofPack");
				}
			}
			replaceOutPick(picks[i]);
		}
		removePickChildren(item);
	}

	function paginationNode(lot, loc, bin) {
		var picks = elementChildren(outRoot, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (normalizeLot(getAttr(picks[i], "LOTNO")) === normalizeLot(lot) &&
					getAttr(picks[i], "LOC") === loc &&
					getAttr(picks[i], "BIN") === bin) {
				return elementChildren(picks[i], "Pagination")[0] || null;
			}
		}
		return null;
	}

	function updateLineFromOutData(lineNo, lot, loc, bin) {
		var picks = elementChildren(outRoot, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (normalizeLot(getAttr(picks[i], "LOTNO")) === normalizeLot(lot) &&
					getAttr(picks[i], "LOC") === loc &&
					getAttr(picks[i], "BIN") === bin) {
				if (field("txtIss" + lineNo)) {
					field("txtIss" + lineNo).value = getAttr(picks[i], "QTYISS");
				}
				if (field("txtTotPackZ" + lineNo)) {
					field("txtTotPackZ" + lineNo).value = getAttr(outRoot, "NoofPack") || "0";
				}
				return;
			}
		}
	}

	function openBagSelector(tempValue, lineNo, lot, loc, bin) {
		openDialog("ItemMergeBagSelPop.asp?sTemp=" + encodeURIComponent(tempValue), xmlIsland("OutData"), "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No", function () {
			var page = paginationNode(lot, loc, bin);
			var raw = page ? getAttr(page, "Value") || getAttr(page, "VALUE") || (page.attributes[0] && page.attributes[0].nodeValue) : "";
			var parts = raw.split("```");
			if (raw && parts.length === 1) {
				openBagSelector(raw.replace(/\|/g, "&"), lineNo, lot, loc, bin);
				return;
			}
			if (parts.length > 1 && trim(parts[1]).substr(0, 2).toUpperCase() === "NO") {
				return;
			}
			updateLineFromOutData(lineNo, lot, loc, bin);
		});
	}

	function numericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function pendingQty() {
		var idQty = byId("idQty");
		return idQty ? toNumber(idQty.textContent || idQty.innerText) : null;
	}

	function selectedLines() {
		var lines = [];
		var elements = form() && form().elements || [];
		for (var i = 0; i < elements.length; i += 1) {
			var match = elements[i].name && elements[i].name.match(/^txtQty(\d+)$/);
			if (match && field("txtIss" + match[1])) {
				lines.push({
					no: match[1],
					stock: elements[i],
					issue: field("txtIss" + match[1]),
					pack: field("txtTotPackZ" + match[1])
				});
			}
		}
		return lines;
	}

	function validateLines() {
		var lines = selectedLines();
		var total = 0;
		for (var i = 0; i < lines.length; i += 1) {
			if (trim(lines[i].issue.value) === "") {
				alert("Enter Quantity Issue");
				lines[i].issue.select();
				return false;
			}
			if (!numericText(lines[i].issue.value)) {
				alert("Enter Numerals Only");
				lines[i].issue.select();
				return false;
			}
			if (toNumber(lines[i].issue.value) > toNumber(lines[i].stock.value)) {
				alert("Quantity Issue should be equal to or less than Stock Quantity " + lines[i].stock.value);
				lines[i].issue.select();
				return false;
			}
			total += toNumber(lines[i].issue.value);
		}
		if (pendingQty() !== null && total > pendingQty()) {
			alert("Quantity Issue should be equal to or less than Quantity Pending " + pendingQty());
			return false;
		}
		return {
			lines: lines,
			total: total
		};
	}

	function applyLineQuantities(lines) {
		var picks = elementChildren(outRoot, "PICK");
		var totalPack = 0;
		for (var i = 0; i < picks.length; i += 1) {
			var line = lines[i];
			var packValue;
			if (!line) {
				continue;
			}
			packValue = line.pack ? line.pack.value : "0";
			setAttr(picks[i], "QTYISS", trim(line.issue.value));
			setAttr(picks[i], "NoofPack", packValue);
			totalPack += toNumber(packValue);
		}
		setAttr(outRoot, "NoofPack", totalPack);
	}

	function appendSelectionToParent(total) {
		var item = selectedItemNode();
		var targetDoc = xmlDocument(objTemp) || (item && item.ownerDocument);
		if (!item) {
			return;
		}
		removePickChildren(item);
		if (total !== 0) {
			setAttr(outRoot, "TOT", total);
			setAttr(item, "QTY", "");
			item.appendChild(importNode(targetDoc, outRoot));
		} else {
			setAttr(outRoot, "TOT", "0");
		}
	}

	function setReturnValue() {
		if (root) {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(root);
			} else {
				window["return" + "Value"] = root;
				window.returnvalue = root;
			}
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function (sItem, sClass, sEntNo) {
		itemCode = String(sItem || "");
		classCode = String(sClass || "");
		entryNo = String(sEntNo || "");
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		outRoot = xmlRoot(xmlIsland("OutData"));
		outDoc = xmlDocument(xmlIsland("OutData")) || (outRoot && outRoot.ownerDocument);
		if (root) {
			setAttr(root, "DONE", "NO");
		}
		loadExistingSelection();
		return false;
	};

	window.CheckLot = function (obj, iEntNo, iClass, iItem, sOrg, sOptName, sAttID, sAttList) {
		var parts = String(obj && obj.name || "").split(":");
		var lineNo = parts[1] || "";
		var lot = parts[2] || "";
		var loc = parts[3] || "";
		var bin = parts[4] || "";
		var stock = parts[5] || "";
		var tempValue = [
			obj.name,
			iClass,
			iItem,
			sOrg,
			iEntNo,
			sOptName,
			sAttID,
			sAttList
		].join(":");
		openBagSelector(tempValue, lineNo, lot, loc, bin, stock);
		return false;
	};

	window.ShowDet = function (description) {
		var text = String(description || "");
		if (text === "-") {
			window.status = "";
			return false;
		}
		window.status = "Passed For Count: " + text.replace(/\|/g, " - ");
		return false;
	};

	window.checkNumbers = numericText;

	window.CheckSubmit = function () {
		var result;
		if (!root) {
			return false;
		}
		setAttr(root, "DONE", "YES");
		result = validateLines();
		if (result === false) {
			return false;
		}
		if (result.total !== 0) {
			applyLineQuantities(result.lines);
		}
		appendSelectionToParent(result.total);
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
