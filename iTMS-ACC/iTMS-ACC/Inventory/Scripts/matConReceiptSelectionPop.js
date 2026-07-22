(function (window, document) {
	"use strict";

	var parentRoot = null;
	var parentDoc = null;
	var localRoot = null;
	var localDoc = null;

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

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(trim(value));
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

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = value;
			element.innerText = value;
		}
	}

	function textOf(id) {
		var element = byId(id);
		return trim(element ? element.textContent || element.innerText || "" : "");
	}

	function fieldValue(name) {
		var item = field(name);
		return trim(item ? item.value : "");
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value;
		}
	}

	function selectedDate(name, hiddenName) {
		var input = field(name);
		if (input) {
			if (typeof input.GetDate === "function") {
				return input.GetDate();
			}
			if (typeof input.getDate === "function") {
				return input.getDate();
			}
			if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
				return window.ITMSModernCompat.toDisplayDate(input.value) || input.value;
			}
			return input.value;
		}
		return fieldValue(hiddenName);
	}

	function setDate(name, value) {
		var input = field(name);
		if (!input || !value) {
			return;
		}
		if (typeof input.SetDate === "function") {
			input.SetDate(value);
		} else if (typeof input.setDate === "function") {
			input.setDate(value);
		} else {
			input.value = value;
		}
	}

	function context() {
		var parts = fieldValue("hObjVal").split(":");
		return {
			type: parts[0] || "",
			issueNo: parts[1] || "",
			item: parts[2] || "",
			classCode: parts[3] || "",
			orgId: parts[4] || fieldValue("hOrgID"),
			callFrom: parts[8] || fieldValue("hCallFrom")
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

	function parseReceipt(value) {
		var parts = String(value || "").split(":");
		return {
			no: trim(parts[0]),
			item: trim(parts[1]),
			attId: trim(parts[2]),
			receiptNumber: trim(parts[3])
		};
	}

	function sameReceipt(node, receipt) {
		return getAttr(node, "No") === receipt.no &&
			getAttr(node, "Item") === receipt.item &&
			getAttr(node, "AttID") === receipt.attId &&
			getAttr(node, "RNumbering") === receipt.receiptNumber;
	}

	function localReceiptNode(receipt) {
		var nodes = elementChildren(localRoot, "RcptItem");
		for (var i = 0; i < nodes.length; i += 1) {
			if (sameReceipt(nodes[i], receipt)) {
				return nodes[i];
			}
		}
		return null;
	}

	function ensureLocalReceiptNode(receipt) {
		var node = localReceiptNode(receipt);
		if (!node && localRoot) {
			node = createElement(localDoc, "RcptItem");
			setAttr(node, "No", receipt.no);
			setAttr(node, "Item", receipt.item);
			setAttr(node, "AttID", receipt.attId);
			setAttr(node, "RNumbering", receipt.receiptNumber);
			setAttr(node, "Qty", "0");
			setAttr(node, "ByProduct", "N");
			localRoot.appendChild(node);
		}
		return node;
	}

	function removeLocalReceiptNode(receipt) {
		var node = localReceiptNode(receipt);
		if (node && node.parentNode) {
			node.parentNode.removeChild(node);
		}
	}

	function rowCount() {
		return toNumber(fieldValue("hCtr"));
	}

	function qtyField(receipt) {
		return field("txtOPQtyZ" + receipt.no + "Z" + receipt.item + "Z" + receipt.attId) ||
			field("txtOPQtyZ" + receipt.no + "Z" + receipt.item + "Z" + receipt.attId + "Z" + receipt.receiptNumber);
	}

	function updateLocalNodeFromRow(index) {
		var checkbox = field("ChkZ" + index);
		var byProduct = field("ChkBPZ" + index);
		var receipt;
		var node;
		var qty;
		if (!checkbox) {
			return;
		}
		receipt = parseReceipt(checkbox.value);
		if (!checkbox.checked) {
			removeLocalReceiptNode(receipt);
			return;
		}
		node = ensureLocalReceiptNode(receipt);
		qty = qtyField(receipt);
		setAttr(node, "Qty", qty ? qty.value : getAttr(node, "Qty") || "0");
		setAttr(node, "ByProduct", byProduct && byProduct.checked ? "Y" : "N");
	}

	function findRowForReceipt(receipt) {
		var checkbox;
		for (var i = 1; i <= rowCount(); i += 1) {
			checkbox = field("ChkZ" + i);
			if (checkbox && sameReceipt({
				getAttribute: function (name) {
					return parseReceipt(checkbox.value)[{
						No: "no",
						Item: "item",
						AttID: "attId",
						RNumbering: "receiptNumber"
					}[name]];
				}
			}, receipt)) {
				return i;
			}
		}
		return 0;
	}

	function hydrateSelectionsFromParent() {
		var itemDet = currentItemDet(false);
		var receipts = elementChildren(itemDet, "RcptItem");
		var receipt;
		var rowIndex;
		var checkbox;
		var byProduct;
		var qty;
		var localNode;
		for (var i = 0; i < receipts.length; i += 1) {
			receipt = {
				no: getAttr(receipts[i], "No"),
				item: getAttr(receipts[i], "Item"),
				attId: getAttr(receipts[i], "AttID"),
				receiptNumber: getAttr(receipts[i], "RNumbering")
			};
			if (!localReceiptNode(receipt) && localRoot) {
				localRoot.appendChild(cloneIntoDoc(receipts[i], localDoc));
			}
			localNode = localReceiptNode(receipt) || receipts[i];
			rowIndex = findRowForReceipt(receipt);
			if (rowIndex) {
				checkbox = field("ChkZ" + rowIndex);
				byProduct = field("ChkBPZ" + rowIndex);
				qty = qtyField(receipt);
				if (checkbox) {
					checkbox.checked = true;
				}
				if (byProduct) {
					byProduct.checked = getAttr(localNode, "ByProduct") === "Y";
				}
				if (qty) {
					qty.value = getAttr(localNode, "Qty") || "0";
				}
			}
		}
	}

	function buildPaginationDetails(status, page) {
		function param(name, value) {
			return name + "=" + encodeURIComponent(value == null ? "" : String(value));
		}
		return encodeURIComponent(fieldValue("hObjVal")) +
			"|" + param("FromDate", selectedDate("ctlFrom", "hFromDate")) +
			"|" + param("ToDate", selectedDate("ctlTo", "hToDate")) +
			"|" + param("ItemCode", fieldValue("hItemCode")) +
			"|" + param("hCurrentPage", fieldValue("hCurrentPage")) +
			"|hWho=" + status +
			"|" + param("hSubmit", page || "") +
			"|" + param("hRowCount", fieldValue("hRowCount")) +
			"|" + param("ClassCode", fieldValue("hClassCode")) +
			"|" + param("CatCode", fieldValue("hCatCode")) +
			"|" + param("ItemName", textOf("txtItem")) +
			"|" + param("ClassName", textOf("txtClass"));
	}

	function setPagination(status, page) {
		var itemDet = currentItemDet(true);
		var pagination;
		if (!itemDet) {
			return;
		}
		removeChildren(itemDet, "Pagination");
		pagination = createElement(parentDoc, "Pagination");
		setAttr(pagination, "Details", buildPaginationDetails(status, page));
		itemDet.appendChild(pagination);
	}

	function syncRowsToLocal() {
		for (var i = 1; i <= rowCount(); i += 1) {
			updateLocalNodeFromRow(i);
		}
	}

	function persistSelections() {
		var itemDet = currentItemDet(true);
		if (!itemDet) {
			return;
		}
		syncRowsToLocal();
		removeChildren(itemDet, "RcptItem");
		elementChildren(localRoot, "RcptItem").forEach(function (node) {
			itemDet.appendChild(cloneIntoDoc(node, parentDoc));
		});
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

	function buildQuery(values) {
		var parts = [];
		Object.keys(values).forEach(function (name) {
			if (values[name] !== undefined && values[name] !== null && values[name] !== "") {
				parts.push(encodeURIComponent(name) + "=" + encodeURIComponent(values[name]));
			}
		});
		return parts.join("&");
	}

	function parseClassificationReturn(value) {
		var text = trim(value);
		var result = { classes: [], categories: [], names: [] };
		var parts;
		var codes;
		var names;
		if (!text || text === "-1") {
			return result;
		}
		parts = text.split("*****");
		if (parts[0] === "-1") {
			return result;
		}
		codes = (parts[0] || "").split("|");
		codes.forEach(function (entry) {
			var path = entry.split(":");
			if (path.length > 1) {
				result.classes.push(path[path.length - 1]);
				result.categories.push(path[1]);
			} else if (path[0]) {
				result.categories.push(path[0].substr(3));
			}
		});
		names = (parts[1] || "").split("|||");
		names.forEach(function (entry) {
			var path = entry.split(":");
			if (path[path.length - 1]) {
				result.names.push(path[path.length - 1]);
			}
		});
		return result;
	}

	function selectedItemNodes(value) {
		var root = xmlRoot(value);
		var nodes = elementChildren(root);
		if (root && (getAttr(root, "ItemCode") || getAttr(root, "RetField1"))) {
			nodes.unshift(root);
		}
		return nodes;
	}

	window.checkNumbers = checkNumbers;

	window.GenXML = function (rowIndex) {
		updateLocalNodeFromRow(rowIndex);
		return false;
	};

	window.SelectClassifcation = function () {
		openModal("/include/ClassificationSelectPop.asp?" + buildQuery({
			sIType: "1",
			sOrgID: fieldValue("hOrgID"),
			sITypename: "",
			SelMode: "M"
		}), "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (returnData) {
			var selected = parseClassificationReturn(returnData);
			setText("txtClass", selected.names.join(","));
			setFieldValue("hClassCode", selected.classes.join(","));
			setFieldValue("hCatCode", selected.categories.join(","));
		});
		return false;
	};

	window.SelectItem = function () {
		var size = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1").split(":") : [];
		var program = size[0] || "ItemSelectCommon.asp";
		var height = size[1] || "500";
		var width = size[2] || "850";
		var url = "../../Common/" + program + "?" + buildQuery({
			orgID: fieldValue("hOrgID"),
			Stock: "Y",
			hSelectMode: "M",
			hDispButt: "N",
			hClassCodes: fieldValue("hClassCode")
		});
		openModal(url, xmlIsland("ItemXML"), "dialogHeight:" + height + "px;dialogWidth:" + width + "px;Status:No", function (outValue) {
			var itemCodes = [];
			var itemNames = [];
			selectedItemNodes(outValue).forEach(function (node) {
				var itemCode = getAttr(node, "ItemCode") || getAttr(node, "RetField1");
				var itemName = getAttr(node, "ItemName") || getAttr(node, "RetField4");
				if (itemCode) {
					itemCodes.push(itemCode);
				}
				if (itemName) {
					itemNames.push(itemName);
				}
			});
			if (itemCodes.length) {
				setText("txtItem", itemNames.join(","));
				setFieldValue("hItemCode", itemCodes.join(","));
			}
		});
		return false;
	};

	window.fnInit = function () {
		var args = modalArgs();
		parentRoot = xmlRoot(args);
		parentDoc = xmlDocument(args) || parentRoot && parentRoot.ownerDocument;
		localRoot = xmlRoot(xmlIsland("ItemData"));
		localDoc = xmlDocument(xmlIsland("ItemData")) || localRoot && localRoot.ownerDocument;
		setDate("ctlFrom", fieldValue("hFromDate"));
		setDate("ctlTo", fieldValue("hToDate"));
		hydrateSelectionsFromParent();
		return false;
	};

	window.PackDisplay = function (receiptNo, itemCode, attList, quantity, receiptNumber, rowIndex) {
		var checkbox = field("ChkZ" + rowIndex);
		var receipt = {
			no: trim(receiptNo),
			item: trim(itemCode),
			attId: trim(attList),
			receiptNumber: trim(receiptNumber)
		};
		if (!checkbox || !checkbox.checked) {
			alert("Please select the Receipt Number to select the Bag Details");
			if (checkbox) {
				checkbox.focus();
			}
			return false;
		}
		ensureLocalReceiptNode(receipt);
		openModal("MatConOPackSelPop.asp?sTemp=" + encodeURIComponent([receiptNo, itemCode, attList, quantity, receiptNumber].join(":")),
			xmlIsland("ItemData"),
			"dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No",
			function () {
				var node = localReceiptNode(receipt);
				var qty = qtyField(receipt);
				if (node && qty) {
					qty.value = getAttr(node, "Qty") || "0";
				}
			});
		return false;
	};

	window.CheckLot = function (obj, index) {
		openModal("MatConSerialPop.asp?sTemp=" + encodeURIComponent(obj && obj.name || ""), xmlIsland("ItemData"), "dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {
			var qty = field("txtOPQtyZ" + index);
			if (qty && localRoot) {
				qty.value = getAttr(localRoot, "SerQtyRet") || qty.value;
			}
		});
		return false;
	};

	window.NextSelection = function (page) {
		persistSelections();
		setPagination("NEXT", page);
		closeWithReturn();
		return false;
	};

	window.CheckSubmit = function () {
		var rows = rowCount();
		if (!rows) {
			closeWithReturn();
			return false;
		}
		for (var i = 1; i <= rows; i += 1) {
			updateLocalNodeFromRow(i);
		}
		persistSelections();
		setPagination("DONE```YES", "");
		closeWithReturn();
		return false;
	};

	window.Func_Close = function () {
		setPagination("DONE```NO", "");
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
