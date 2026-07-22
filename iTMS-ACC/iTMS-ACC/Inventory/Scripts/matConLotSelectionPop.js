(function (window, document) {
	"use strict";

	var parentRoot = null;
	var parentDoc = null;
	var localRoot = null;
	var localDoc = null;
	var currentLotNo = "";
	var currentIssueDate = "";

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
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot.toUpperCase() === "NULL" || lot.toUpperCase() === "N/A" ? "0" : lot;
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
		return window[name] || document[name] || document.getElementById(name);
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

	function createElement(doc, name) {
		if (doc && doc.createElement) {
			return doc.createElement(name);
		}
		return document.implementation.createDocument("", name, null).documentElement;
	}

	function cloneIntoDoc(node, doc) {
		if (!node) {
			return null;
		}
		if (doc && doc.importNode) {
			return doc.importNode(node, true);
		}
		return node.cloneNode(true);
	}

	function removeChildren(parent, name) {
		elementChildren(parent, name).forEach(function (node) {
			parent.removeChild(node);
		});
	}

	function fieldValue(name) {
		var item = field(name);
		return trim(item ? item.value : "");
	}

	function rowCount() {
		return toNumber(fieldValue("hiCtr"));
	}

	function context() {
		var parts = fieldValue("hObjVal").split(":");
		return {
			issueNo: parts[1] || "",
			item: parts[2] || "",
			classCode: parts[3] || "",
			callFrom: parts[9] || fieldValue("hCallFrom")
		};
	}

	function currentItemDet(createIfMissing) {
		var ctx = context();
		var nodes = elementChildren(parentRoot, "ItemDet");
		var node = null;
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "IssEntryNo") !== ctx.issueNo) {
				continue;
			}
			if (ctx.callFrom === "ISSNO" && getAttr(nodes[i], "Item") === "" && getAttr(nodes[i], "Class") === "") {
				node = nodes[i];
				break;
			}
			if (ctx.callFrom !== "ISSNO" && getAttr(nodes[i], "Item") === ctx.item && getAttr(nodes[i], "Class") === ctx.classCode) {
				node = nodes[i];
				break;
			}
		}
		if (!node && ctx.callFrom === "ISSNO") {
			for (i = 0; i < nodes.length; i += 1) {
				if (getAttr(nodes[i], "IssEntryNo") === ctx.issueNo) {
					node = nodes[i];
					break;
				}
			}
		}
		if (!node && createIfMissing && parentRoot && parentDoc) {
			node = createElement(parentDoc, "ItemDet");
			setAttr(node, "Item", ctx.callFrom === "ISSNO" ? "" : ctx.item);
			setAttr(node, "Class", ctx.callFrom === "ISSNO" ? "" : ctx.classCode);
			setAttr(node, "IssEntryNo", ctx.issueNo);
			setAttr(node, "Remarks", "");
			setAttr(node, "Qty", "0");
			setAttr(node, "AttributeList", "");
			parentRoot.appendChild(node);
		}
		return node;
	}

	function localLotNode(lotNo, issueDate) {
		var normalized = normalizeLot(lotNo);
		var nodes = elementChildren(localRoot, "LotDet");
		for (var i = 0; i < nodes.length; i += 1) {
			if (normalizeLot(getAttr(nodes[i], "LotNo")) === normalized && getAttr(nodes[i], "IssueDate") === trim(issueDate)) {
				return nodes[i];
			}
		}
		return null;
	}

	function hydrateSelectionsFromParent() {
		var itemDet = currentItemDet(false);
		var parentLots = elementChildren(itemDet, "LotDet");
		var localLot;
		var qty;
		var index;
		for (var i = 0; i < parentLots.length; i += 1) {
			localLot = localLotNode(getAttr(parentLots[i], "LotNo"), getAttr(parentLots[i], "IssueDate"));
			if (!localLot || !localRoot) {
				continue;
			}
			setAttr(localLot, "QtyRet", getAttr(parentLots[i], "QtyRet"));
			removeChildren(localLot, "SerialDetails");
			elementChildren(parentLots[i], "SerialDetails").forEach(function (serialNode) {
				localLot.appendChild(cloneIntoDoc(serialNode, localDoc));
			});
			index = elementChildren(localRoot, "LotDet").indexOf(localLot) + 1;
			qty = field("txtQty" + index);
			if (qty) {
				qty.value = getAttr(parentLots[i], "QtyRet") || "0";
			}
		}
	}

	function setReturnValue() {
		if (parentRoot) {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(parentRoot);
			} else {
				window["return" + "Value"] = parentRoot;
				window.returnvalue = parentRoot;
			}
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function lotQuantityFromSerials(lotNode) {
		var total = 0;
		elementChildren(lotNode, "SerialDetails").forEach(function (serialNode) {
			total += toNumber(getAttr(serialNode, "Qty"));
		});
		return total;
	}

	function validateRows() {
		var qty;
		var remaining;
		var totalQty = toNumber(fieldValue("idQty"));
		for (var i = 1; i <= rowCount(); i += 1) {
			qty = field("txtQty" + i);
			remaining = field("txtRemQty" + i);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity to Return");
				if (qty) {
					qty.select();
				}
				return false;
			}
			if (!isNumericText(qty.value)) {
				alert("Enter Numerals Only");
				qty.select();
				return false;
			}
			if (toNumber(qty.value) > toNumber(remaining && remaining.value) || toNumber(qty.value) > totalQty) {
				alert("Return Quantity should be equal to or less than Remaining Quantity");
				qty.select();
				return false;
			}
		}
		return true;
	}

	window.checkNumbers = function (value) {
		return isNumericText(value);
	};

	window.fnInit = function () {
		var args = modalArgs();
		parentRoot = xmlRoot(args);
		parentDoc = xmlDocument(args) || parentRoot && parentRoot.ownerDocument;
		localRoot = xmlRoot(xmlIsland("ItemData"));
		localDoc = xmlDocument(xmlIsland("ItemData")) || localRoot && localRoot.ownerDocument;
		hydrateSelectionsFromParent();
		return false;
	};

	window.CheckLot = function (obj, index) {
		var parts = String(obj && obj.name || "").split("`");
		var lotNo = normalizeLot(parts[5] || "");
		var issueDate = parts[9] || "";
		openModal("MatConSerialPop.asp?sTemp=" + encodeURIComponent(obj && obj.name || ""), xmlIsland("ItemData"), "dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {
			var lotNode = localLotNode(lotNo, issueDate);
			var qty = field("txtQty" + index);
			if (lotNode && qty) {
				qty.value = lotQuantityFromSerials(lotNode);
				setAttr(lotNode, "QtyRet", qty.value);
			}
		});
		return false;
	};

	window.CheckSubmit = function () {
		var itemDet;
		var localLots;
		var qty;
		if (!rowCount()) {
			return false;
		}
		if (!validateRows()) {
			return false;
		}
		itemDet = currentItemDet(true);
		if (!itemDet) {
			return false;
		}
		localLots = elementChildren(localRoot, "LotDet");
		for (var i = 0; i < localLots.length; i += 1) {
			qty = field("txtQty" + (i + 1));
			setAttr(localLots[i], "QtyRet", qty ? qty.value : "0");
		}
		removeChildren(itemDet, "LotDet");
		localLots.forEach(function (lotNode) {
			itemDet.appendChild(cloneIntoDoc(lotNode, parentDoc));
		});
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
