(function (window, document) {
	"use strict";

	var objTemp = null;
	var parentRoot = null;
	var parentDoc = null;
	var classCode = "";
	var itemCode = "";
	var itemEntryNo = "";

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

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function textOf(id) {
		var item = document.getElementById(id);
		return trim(item ? item.innerText || item.textContent || "" : "");
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

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function pickDataDocument() {
		return xmlDocument(xmlIsland("PickData"));
	}

	function pickDataRoot() {
		return xmlRoot(xmlIsland("PickData"));
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

	function setPickDataRoot(node) {
		var island = xmlIsland("PickData");
		var doc = pickDataDocument();
		var serialized;
		var clone;
		if (!node || !doc) {
			return;
		}
		if (island && typeof island.loadXML === "function" && window.XMLSerializer) {
			serialized = new window.XMLSerializer().serializeToString(node);
			island.loadXML(serialized);
			return;
		}
		clone = importForDocument(doc, node);
		while (doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		doc.appendChild(clone);
	}

	function findParentItem(classValue, itemValue, entryNo) {
		var items = elementChildren(parentRoot, "ITM");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "CLACODE") === trim(classValue) &&
					getAttr(items[i], "ITMCODE") === trim(itemValue) &&
					(getAttr(items[i], "ItemEntNo") === trim(entryNo) || getAttr(items[i], "PICKNO") === trim(entryNo))) {
				return items[i];
			}
		}
		return null;
	}

	function findParentPickDet(classValue, itemValue, entryNo) {
		var itemNode = findParentItem(classValue, itemValue, entryNo);
		return elementChildren(itemNode, "PickDet")[0] || null;
	}

	function findPick(root, predicate) {
		var picks = elementChildren(root, "PICK");
		for (var i = 0; i < picks.length; i += 1) {
			if (predicate(picks[i])) {
				return picks[i];
			}
		}
		return null;
	}

	function findPickBySerial(root, serialNo) {
		return findPick(root, function (pick) {
			return getAttr(pick, "SNO") === trim(serialNo);
		});
	}

	function findPickForPagination(root, storeCode, binCode, entryNo, lotNo) {
		return findPick(root, function (pick) {
			return getAttr(pick, "LOC") === trim(storeCode) &&
				getAttr(pick, "BIN") === trim(binCode) &&
				getAttr(pick, "ITEMENTRYNO") === trim(entryNo) &&
				(!lotNo || getAttr(pick, "LOTNO") === trim(lotNo));
		}) || findPick(root, function (pick) {
			return getAttr(pick, "LOC") === trim(storeCode) &&
				getAttr(pick, "BIN") === trim(binCode) &&
				getAttr(pick, "ITEMENTRYNO") === trim(entryNo);
		});
	}

	function issueRows() {
		var frm = form();
		var elements = frm && frm.elements || [];
		var rows = [];
		var name;
		var suffix;
		for (var i = 0; i < elements.length; i += 1) {
			name = String(elements[i].name || "");
			if (String(elements[i].type || "").toLowerCase() === "text" && name.indexOf("txtQty") === 0) {
				suffix = name.substring("txtQty".length);
				if (field("txtIss" + suffix)) {
					rows.push({
						suffix: suffix,
						marked: elements[i],
						stock: field("hStock" + suffix),
						issue: field("txtIss" + suffix)
					});
				}
			}
		}
		return rows;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function setReturnValue() {
		if (!parentRoot) {
			parentRoot = xmlRoot(objTemp);
		}
		if (!parentRoot) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(parentRoot);
		} else {
			window["return" + "Value"] = parentRoot;
			window.returnvalue = parentRoot;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function paginationDetails(root, storeCode, binCode, entryNo, lotNo) {
		var pick = findPickForPagination(root, storeCode, binCode, entryNo, lotNo);
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

	function refreshIssueQuantity(lineNo) {
		var root = pickDataRoot();
		var pick = findPickBySerial(root, lineNo);
		var quantity = field("txtIss" + lineNo);
		if (pick && quantity && elementChildren(pick).length) {
			quantity.value = getAttr(pick, "ISSQTY");
		}
	}

	function continueLotDialog(lineNo, storeCode, binCode, entryNo, issueEntryNo, lotNo) {
		var root = pickDataRoot();
		var page = paginationDetails(root, storeCode, binCode, entryNo, lotNo);
		var url;
		if (page && page.parts.length === 1) {
			url = "IssuePickLotPop.asp?sTemp=" + page.raw.replace(/\|/g, "&") + "&IssNo=" + encodeURIComponent(issueEntryNo);
			openModal(url, pickDataDocument(), "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No", function () {
				continueLotDialog(lineNo, storeCode, binCode, entryNo, issueEntryNo, lotNo);
			});
			return;
		}
		if (page && page.parts.length > 1 && trim(page.parts[1]).substring(0, 2) === "NO") {
			return;
		}
		refreshIssueQuantity(lineNo);
	}

	function updatePickRootFromRows(root) {
		var picks = elementChildren(root, "PICK");
		var quantity;
		for (var i = 0; i < picks.length; i += 1) {
			quantity = field("txtIss" + getAttr(picks[i], "SNO"));
			if (quantity) {
				setAttr(picks[i], "ISSQTY", trim(quantity.value) === "" ? "" : toNumber(quantity.value));
			}
		}
	}

	function mergePickRootIntoParent(root) {
		var itemNode = findParentItem(classCode, itemCode, itemEntryNo);
		var oldPickDet = elementChildren(itemNode, "PickDet")[0];
		var imported;
		if (!itemNode || !parentDoc || !root) {
			return;
		}
		imported = importForDocument(parentDoc, root);
		if (oldPickDet) {
			itemNode.replaceChild(imported, oldPickDet);
		} else {
			itemNode.appendChild(imported);
		}
	}

	window.fnInit = function (sClass, sItm, iEntryNo) {
		var parentPickDet;
		var picks;
		var quantity;
		objTemp = modalArgs();
		parentRoot = xmlRoot(objTemp);
		parentDoc = xmlDocument(objTemp);
		classCode = sClass;
		itemCode = sItm;
		itemEntryNo = iEntryNo;
		if (!parentDoc && parentRoot) {
			parentDoc = parentRoot.ownerDocument;
		}
		if (!field("hiCtr") || trim(field("hiCtr").value) === "") {
			return;
		}
		parentPickDet = findParentPickDet(sClass, sItm, iEntryNo);
		if (parentPickDet) {
			setPickDataRoot(parentPickDet);
			picks = elementChildren(parentPickDet, "PICK");
			for (var i = 0; i < picks.length; i += 1) {
				quantity = field("txtIss" + getAttr(picks[i], "SNO"));
				if (quantity) {
					quantity.value = getAttr(picks[i], "ISSQTY");
				}
			}
		}
	};

	window.CheckLot = function (obj, lineNo, issueQty, itemValue, classValue, attId) {
		var parts = String(obj && obj.name || "").split("`");
		var entryNo = parts[1] || "";
		var issueEntryNo = parts[2] || "";
		var lotNo = parts[3] || "";
		var storeCode = parts[4] || "";
		var binCode = parts[5] || "";
		var orgId = field("hOrgID") ? field("hOrgID").value : "";
		var parentPickDet = findParentPickDet(classValue, itemValue, entryNo);
		var tempValues;
		var url;

		if (parentPickDet) {
			setPickDataRoot(parentPickDet);
		}

		tempValues = "hValueZ:" + lineNo + ":" + lotNo + ":" + storeCode + ":" + binCode + "::" + classValue + ":" + itemValue + ":" + orgId + "::" + entryNo + "::" + attId + ":";
		url = "IssuePickLotPop.asp?sTemp=" + encodeURIComponent(tempValues) + "&IssNo=" + encodeURIComponent(issueEntryNo);
		openModal(url, pickDataDocument(), "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No", function () {
			continueLotDialog(lineNo, storeCode, binCode, entryNo, issueEntryNo, lotNo);
		});
	};

	window.CheckSubmit = function (todaysDate, itemValue, classValue, entryNo) {
		var rows = issueRows();
		var totalWithLot = 0;
		var totalWithoutLot = 0;
		var total = 0;
		var row;
		var issueValue;
		var root = pickDataRoot();

		classCode = classValue;
		itemCode = itemValue;
		itemEntryNo = entryNo;

		for (var i = 0; i < rows.length; i += 1) {
			row = rows[i];
			if (trim(row.issue.value) === "") {
				issueValue = 0;
			} else if (!checkNumbers(row.issue.value)) {
				alert("Enter Numerals Only");
				focusOrSelect(row.issue);
				return;
			} else {
				issueValue = toNumber(row.issue.value);
			}

			if (trim(row.marked.value) !== "-") {
				totalWithLot += issueValue;
				if (issueValue > toNumber(row.marked.value)) {
					alert("Quantity Issue should be less than or equal to Quantity Marked");
					focusOrSelect(row.issue);
					return;
				}
				if (issueValue > toNumber(row.stock && row.stock.value)) {
					alert("Quantity Issue should be less than or equal to Stock Quantity");
					focusOrSelect(row.issue);
					return;
				}
			} else {
				totalWithoutLot += issueValue;
				if (issueValue > toNumber(row.stock && row.stock.value)) {
					alert("Quantity Issue should be less than or equal to Stock Quantity");
					focusOrSelect(row.issue);
					return;
				}
			}
			total += issueValue;
		}

		if (total === 0) {
			alert("Enter Quantity Issue or select from Serial number(s)");
			return;
		}
		if (totalWithLot > toNumber(textOf("idWQty"))) {
			alert("Quantity Issue from Lot Selection should be less than or equal to (" + textOf("idWQty") + ")");
			return;
		}
		if (totalWithoutLot > toNumber(textOf("idWOQty"))) {
			alert("Quantity Issue from without Lot Selection should be less than or equal to (" + textOf("idWOQty") + ")");
			return;
		}
		if (total > toNumber(textOf("idQty"))) {
			alert("Quantity Issue should be less than or equal to total quantity available for Pick (" + textOf("idQty") + ")");
			return;
		}

		setAttr(root, "TOT", total);
		updatePickRootFromRows(root);
		mergePickRootIntoParent(root);
		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
