(function (window, document) {
	"use strict";

	var objTemp = null;
	var parentRoot = null;
	var parentDoc = null;
	var itemCode = "";
	var classCode = "";
	var entryNo = "";
	var attributeList = "";

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
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function rowField(prefix, row) {
		return field(prefix + row) || field(prefix.charAt(0).toUpperCase() + prefix.slice(1) + row);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function toInteger(value) {
		var parsed = parseInt(String(value == null ? "" : value).replace(/,/g, ""), 10);
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
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

	function outDataDocument() {
		return xmlDocument(xmlIsland("OutData"));
	}

	function outDataRoot() {
		return xmlRoot(xmlIsland("OutData"));
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

	function encodedAttribute(value) {
		return trim(value).replace(/'/g, "~~").replace(/#/g, "$").replace(/:/g, "@");
	}

	function attrMatches(nodeValue, wanted) {
		var nodeText = trim(nodeValue);
		var wantedText = trim(wanted);
		return nodeText === wantedText || encodedAttribute(nodeText) === wantedText;
	}

	function itemAttr(node, oldName, newName) {
		return getAttr(node, oldName) || getAttr(node, newName);
	}

	function itemMatches(node, requireEntry) {
		var name = String(node && node.nodeName || "").toLowerCase();
		if (name !== "itemdetails" && name !== "item") {
			return false;
		}
		if (itemAttr(node, "CLASSCODE", "CLACODE") !== trim(classCode) || itemAttr(node, "ITEMCODE", "ITMCODE") !== trim(itemCode)) {
			return false;
		}
		if (requireEntry && itemAttr(node, "ENTRYNO", "ItemEntNo") !== trim(entryNo)) {
			return false;
		}
		return attrMatches(getAttr(node, "ATTRIBUTELIST"), attributeList);
	}

	function findItemNode(requireEntry) {
		var nodes = elementChildren(parentRoot);
		for (var i = 0; i < nodes.length; i += 1) {
			if (itemMatches(nodes[i], requireEntry)) {
				return nodes[i];
			}
		}
		return null;
	}

	function matchingOutNode(root, loc, bin, lot) {
		var nodes = elementChildren(root);
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "LOTNO") === trim(lot) && getAttr(nodes[i], "LOC") === trim(loc) && getAttr(nodes[i], "BIN") === trim(bin)) {
				return nodes[i];
			}
		}
		return null;
	}

	function replaceOutNode(outRoot, incoming) {
		var existing = matchingOutNode(outRoot, getAttr(incoming, "LOC"), getAttr(incoming, "BIN"), getAttr(incoming, "LOTNO"));
		var imported = importForDocument(outDataDocument(), incoming);
		if (!imported) {
			return;
		}
		if (existing) {
			outRoot.replaceChild(imported, existing);
		} else {
			outRoot.appendChild(imported);
		}
	}

	function removeChildrenByName(node, name) {
		var nodes = elementChildren(node, name);
		for (var i = 0; i < nodes.length; i += 1) {
			node.removeChild(nodes[i]);
		}
	}

	function removeAllElementChildren(node) {
		var nodes = elementChildren(node);
		for (var i = 0; i < nodes.length; i += 1) {
			node.removeChild(nodes[i]);
		}
	}

	function issueRows() {
		var frm = form();
		var elements = frm && frm.elements || [];
		var rows = [];
		var match;
		for (var i = 0; i < elements.length; i += 1) {
			match = String(elements[i].name || "").match(/^txtQtyZ(\d+)$/i);
			if (match) {
				rows.push({
					index: match[1],
					stock: elements[i],
					issue: rowField("txtIssZ", match[1]),
					pack: rowField("txtTotPackZ", match[1])
				});
			}
		}
		return rows;
	}

	function textOf(id) {
		var item = document.getElementById(id);
		return trim(item ? item.textContent || item.innerText || "" : "");
	}

	function selectedLotPackValue() {
		var radio = field("radLotPack");
		if (!radio) {
			return "";
		}
		if (radio.length !== undefined) {
			for (var i = 0; i < radio.length; i += 1) {
				if (radio[i].checked) {
					return radio[i].value;
				}
			}
			return radio[0] ? radio[0].value : "";
		}
		return radio.checked ? radio.value : "";
	}

	function setLotPackValue(value) {
		var radio = field("radLotPack");
		if (!radio || radio.length === undefined) {
			return;
		}
		for (var i = 0; i < radio.length; i += 1) {
			radio[i].checked = radio[i].value === value;
		}
	}

	function setReturnValue() {
		if (!parentRoot) {
			parentRoot = xmlRoot(objTemp);
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

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function paginationFor(root, loc, bin, lot) {
		var pick = matchingOutNode(root, loc, bin, lot);
		var pagination = elementChildren(pick, "Pagination")[0];
		var attr = pagination && pagination.attributes && pagination.attributes.item(0);
		var value = attr ? attr.nodeValue || attr.value || "" : "";
		return value ? {
			raw: value,
			parts: String(value).split("```")
		} : null;
	}

	function refreshLotRow(lineNo, loc, bin, lot) {
		var root = outDataRoot();
		var node = matchingOutNode(root, loc, bin, lot);
		var issue = rowField("txtIssZ", lineNo);
		var pack = rowField("txtTotPackZ", lineNo);
		var issueQty;
		var noOfPack;

		if (!node) {
			return;
		}
		issueQty = getAttr(node, "QTYISS") || "0";
		if (issue) {
			issue.value = issueQty;
		}
		noOfPack = getAttr(root, "NoofPack");
		if (pack) {
			pack.value = noOfPack && noOfPack !== "0" && issueQty !== "0" ? noOfPack : "0";
		}
	}

	function continueLotDialog(lineNo, loc, bin, lot) {
		var page = paginationFor(outDataRoot(), loc, bin, lot);
		if (page && page.parts.length === 1) {
			openModal("mrsPickDetailLotPop.asp?sTemp=" + page.raw.replace(/\|/g, "&"), xmlIsland("OutData"), "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No", function () {
				continueLotDialog(lineNo, loc, bin, lot);
			});
			return;
		}
		if (page && page.parts.length > 1 && trim(page.parts[1]).substring(0, 2) === "NO") {
			return;
		}
		if (trim(field("hType") && field("hType").value) === "M" &&
				trim(field("hPickPackFlag") && field("hPickPackFlag").value) === "N" &&
				trim(field("hRcptNumbering") && field("hRcptNumbering").value) === "LS") {
			setLotPackValue("P");
		}
		refreshLotRow(lineNo, loc, bin, lot);
	}

	window.ShowDet = function (description) {
		var tooltip = document.getElementById("itms-pick-tooltip");
		var eventObject = window.event || null;
		var parts;
		if (description === "-") {
			if (tooltip) {
				tooltip.style.display = "none";
			}
			return;
		}
		if (!tooltip) {
			tooltip = document.createElement("div");
			tooltip.id = "itms-pick-tooltip";
			tooltip.style.cssText = "position:absolute;z-index:9999;display:none;max-width:260px;padding:6px 8px;border:1px solid #777;background:#ffffe1;color:#111;font:11px Verdana,Arial,sans-serif;white-space:pre-line;";
			document.body.appendChild(tooltip);
		}
		parts = String(description || "").split("|");
		tooltip.textContent = "Passed For Count From : " + (parts[0] || "") + "\nPassed For Count To      : " + (parts[1] || "");
		tooltip.style.left = ((eventObject && eventObject.clientX || 10) + 12 + (window.pageXOffset || document.documentElement.scrollLeft || 0)) + "px";
		tooltip.style.top = ((eventObject && eventObject.clientY || 10) + 12 + (window.pageYOffset || document.documentElement.scrollTop || 0)) + "px";
		tooltip.style.display = "block";
	};

	window.EnableLot = function () {
		var root = outDataRoot();
		var rows = toInteger(field("hCtr") && field("hCtr").value);
		var lotPackValue;
		var pickNodes;
		var button;
		var issue;

		if (trim(field("hType") && field("hType").value) !== "M" ||
				trim(field("hPickPackFlag") && field("hPickPackFlag").value) !== "N" ||
				trim(field("hRcptNumbering") && field("hRcptNumbering").value) !== "LS") {
			return;
		}

		lotPackValue = selectedLotPackValue();
		if (lotPackValue === "L" && root) {
			pickNodes = elementChildren(root, "PICK");
			for (var p = 0; p < pickNodes.length; p += 1) {
				removeAllElementChildren(pickNodes[p]);
			}
		}

		for (var i = 1; i <= rows; i += 1) {
			button = rowField("btnSerialZ", i);
			issue = rowField("txtIssZ", i);
			if (button) {
				button.disabled = lotPackValue === "L";
			}
			if (issue) {
				issue.disabled = lotPackValue !== "L";
			}
		}
	};

	window.fnInit = function (sItem, sClass, sEntNo, sAttriList) {
		var outRoot;
		var itemNode;
		var picks;
		var pickChildren;
		var outMatch;
		var rowNo;

		itemCode = sItem;
		classCode = sClass;
		entryNo = sEntNo;
		attributeList = sAttriList;
		objTemp = modalArgs();
		parentRoot = xmlRoot(objTemp);
		parentDoc = xmlDocument(objTemp);
		if (!parentDoc && parentRoot) {
			parentDoc = parentRoot.ownerDocument;
		}
		if (parentRoot) {
			setAttr(parentRoot, "DONE", "NO");
		}
		outRoot = outDataRoot();
		itemNode = findItemNode(true);
		if (!itemNode || !outRoot) {
			return;
		}

		picks = elementChildren(itemNode, "Pick");
		for (var i = 0; i < picks.length; i += 1) {
			setAttr(outRoot, "TOT", getAttr(picks[i], "TOT"));
			pickChildren = elementChildren(picks[i]);
			for (var p = 0; p < pickChildren.length; p += 1) {
				outMatch = matchingOutNode(outRoot, getAttr(pickChildren[p], "LOC"), getAttr(pickChildren[p], "BIN"), getAttr(pickChildren[p], "LOTNO"));
				if (outMatch) {
					rowNo = getAttr(outMatch, "Count");
					if (rowField("txtIssZ", rowNo)) {
						rowField("txtIssZ", rowNo).value = getAttr(pickChildren[p], "QTYISS");
					}
					if (rowField("txtTotPackZ", rowNo)) {
						rowField("txtTotPackZ", rowNo).value = getAttr(pickChildren[p], "NoofPack");
					}
				}
				replaceOutNode(outRoot, pickChildren[p]);
			}
			itemNode.removeChild(picks[i]);
		}

		if (trim(field("hType") && field("hType").value) === "M" && trim(field("hPickPackFlag") && field("hPickPackFlag").value) === "N") {
			if (elementChildren(outRoot, "PICK").length <= 1 && elementChildren(outRoot).length > 0) {
				window.CheckLot("1", entryNo, classCode, itemCode, field("hOrgID") ? field("hOrgID").value : "", "", field("hOptName") ? field("hOptName").value : "", field("hAttID") ? field("hAttID").value : "", field("hAttList") ? field("hAttList").value : "");
			}
		}
	};

	window.CheckLot = function (row, rowEntryNo, rowClass, rowItem, orgId, invRecNo, optionName, attrId, attrList) {
		var holder = rowField("hValueZ", row);
		var parts;
		var lineNo;
		var lot;
		var loc;
		var bin;
		var tempValues;

		if (!holder) {
			return false;
		}
		parts = String(holder.value || "").split(":");
		lineNo = parts[1] || row;
		lot = parts[2] || "";
		loc = parts[3] || "";
		bin = parts[4] || "";
		tempValues = holder.value + ":" + rowClass + ":" + rowItem + ":" + orgId + ":" + invRecNo + ":" + rowEntryNo + ":" + optionName + ":" + attrId + ":" + attrList;

		openModal("mrsPickDetailLotPop.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData"), "dialogHeight:580px;dialogWidth:580px;center:Yes;help:No;resizable:No;status:No", function () {
			continueLotDialog(lineNo, loc, bin, lot);
		});
		return false;
	};

	window.CheckSubmit = function () {
		var outRoot = outDataRoot();
		var rows = issueRows();
		var totalQty = 0;
		var totalPacks = 0;
		var parentItem;
		var outChildren;
		var row;
		var issueValue;
		var pendingText;
		var count;

		if (parentRoot) {
			setAttr(parentRoot, "DONE", "YES");
		}
		if (!outRoot) {
			return false;
		}

		for (var i = 0; i < rows.length; i += 1) {
			row = rows[i];
			if (!row.issue || trim(row.issue.value) === "") {
				alert("Enter Quantity Issue");
				focusOrSelect(row.issue);
				return false;
			}
			if (!isNumericText(row.issue.value)) {
				alert("Enter Numerals Only");
				focusOrSelect(row.issue);
				return false;
			}
			issueValue = toNumber(row.issue.value);
			if (issueValue > toNumber(row.stock && row.stock.value)) {
				alert("Quantity Issue should be equal to or less than Stock Quantity " + (row.stock ? row.stock.value : ""));
				focusOrSelect(row.issue);
				return false;
			}
			totalQty += issueValue;
		}

		pendingText = textOf("idQty");
		if (pendingText !== "" && totalQty > toNumber(pendingText)) {
			alert("Quantity Issue should be equal to or less than Quantity Pending " + pendingText);
			return false;
		}

		if (totalQty !== 0) {
			outChildren = elementChildren(outRoot);
			for (i = 0; i < outChildren.length; i += 1) {
				count = getAttr(outChildren[i], "Count");
				setAttr(outChildren[i], "QTYISS", rowField("txtIssZ", count) ? rowField("txtIssZ", count).value : "");
				setAttr(outChildren[i], "NoofPack", rowField("txtTotPackZ", count) ? rowField("txtTotPackZ", count).value : "0");
				totalPacks += toInteger(rowField("txtTotPackZ", count) && rowField("txtTotPackZ", count).value);
			}

			setAttr(outRoot, "TOT", totalQty);
			if (trim(field("hType") && field("hType").value) === "M" &&
					trim(field("hPickPackFlag") && field("hPickPackFlag").value) === "N" &&
					trim(field("hRcptNumbering") && field("hRcptNumbering").value) === "LS") {
				setAttr(outRoot, "ONLYLOT", selectedLotPackValue());
			} else {
				setAttr(outRoot, "ONLYLOT", "");
			}
			setAttr(outRoot, "NoofPack", totalPacks);

			parentItem = findItemNode(true) || findItemNode(false);
			if (parentItem) {
				removeChildrenByName(parentItem, "Pick");
				setAttr(parentItem, "QTY", "");
				parentItem.appendChild(importForDocument(parentDoc, outRoot));
			}
		} else {
			setAttr(outRoot, "TOT", "0");
			parentItem = findItemNode(true) || findItemNode(false);
			if (parentItem) {
				removeChildrenByName(parentItem, "Pick");
			}
		}

		setAttr(outRoot, "NoofPack", totalPacks);
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
