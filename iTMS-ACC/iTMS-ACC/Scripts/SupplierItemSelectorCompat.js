(function (window, document) {
	"use strict";

	var buttonPressed = "";

	function loadCompat() {
		var loader;
		if (window.ITMSModernCompat || document.querySelector('script[src*="itms-modern-compat.js"]')) {
			return;
		}
		loader = document.createElement("script");
		loader.type = "text/javascript";
		loader.src = "../scripts/itms-modern-compat.js";
		(document.head || document.documentElement).appendChild(loader);
	}

	loadCompat();

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.FormName || document.forms.formname || document.forms.FormName || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function asArray(collection) {
		return Array.prototype.slice.call(collection || []);
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return asArray(node && node.childNodes).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function xmlDocument(object) {
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function dialogArgs() {
		var args = window.dialogArguments;
		var match;
		var id;
		if (!args) {
			match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
			id = match ? decodeURIComponent(match[1]) : "";
			if (id && window.opener && window.opener.__itmsDialogArgs) {
				args = window.opener.__itmsDialogArgs[id];
				window.dialogArguments = args;
			}
		}
		return args;
	}

	function root() {
		var object = dialogArgs();
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || null;
	}

	function createNode(name) {
		var doc = xmlDocument(dialogArgs());
		if (doc && doc.createElement) {
			return doc.createElement(name);
		}
		return document.implementation.createDocument("", "", null).createElement(name);
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function notifyDialogValue(id, value) {
		if (!id || !window.opener) {
			return;
		}
		try {
			if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
				window.opener.ITMSModernCompat._receiveDialogValue(id, value);
				return;
			}
		} catch (ignoreDirectReturn) {}
		try {
			window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
		} catch (ignoreMessageReturn) {}
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		notifyDialogValue(id, value);
	}

	function closeWithRoot() {
		returnValue(root());
		window.close();
	}

	function selectedItems() {
		return childElements(root(), "Item");
	}

	function selectedMaterials() {
		var result = [];
		childElements(root(), "Materials").forEach(function (materials) {
			result = result.concat(childElements(materials, "Entry"));
		});
		return result;
	}

	function nextEntryNo() {
		var max = 0;
		selectedItems().forEach(function (item) {
			var value = parseInt(attr(item, "EntryNo"), 10);
			if (!isNaN(value) && value > max) {
				max = value;
			}
		});
		return max + 1;
	}

	function selectMode() {
		return trim(fieldValue("hSelectMode")).toUpperCase();
	}

	function removeAllItems() {
		selectedItems().forEach(function (item) {
			root().removeChild(item);
		});
	}

	function hiddenValue(prefix, parts) {
		return fieldValue(prefix + (parts[1] || "") + (parts[2] || ""));
	}

	function selectedAttribute(parts, includeZero) {
		var hasAttributes = hiddenValue("hAttribVal", parts);
		var select = field("SelAttribList" + (parts[1] || "") + (parts[2] || ""));
		var value;
		if (!trim(hasAttributes) || trim(hasAttributes).toUpperCase() === "N" || !select) {
			return "";
		}
		value = select.value || "";
		if (!includeZero && value.indexOf("0:") === 0) {
			return "";
		}
		return value;
	}

	function selectedAttributeText(attributeList) {
		var text = "";
		trim(attributeList).split(",").forEach(function (entry) {
			var parts = entry.split(":");
			var codeParts = String(parts[0] || "").split("#");
			if (trim(codeParts[1]) !== "0" && trim(parts[1])) {
				text += (text ? "," : "") + trim(parts[1]);
			}
		});
		return text ? "[" + text + "]" : "";
	}

	function itemExists(parts, attributeList) {
		return selectedItems().some(function (item) {
			return trim(attr(item, "ItemCode")) === trim(parts[1]) &&
				trim(attr(item, "ClassCode")) === trim(parts[2]) &&
				trim(attr(item, "AttributeList")) === trim(attributeList || "");
		});
	}

	function removeItem(parts, attributeList) {
		selectedItems().forEach(function (item) {
			var matches = trim(attr(item, "ItemCode")) === trim(parts[1]) && trim(attr(item, "ClassCode")) === trim(parts[2]);
			if (attributeList != null) {
				matches = matches && trim(attr(item, "AttributeList")) === trim(attributeList);
			}
			if (matches) {
				root().removeChild(item);
			}
		});
	}

	function valueOrZero(value) {
		return trim(value) || "0";
	}

	function appendItem(parts, attributeList, appendAttributeText) {
		var node;
		if (itemExists(parts, attributeList)) {
			return;
		}
		node = createNode("Item");
		setAttr(node, "EntryNo", nextEntryNo());
		setAttr(node, "CompanyItemCode", parts[0] || "");
		setAttr(node, "ItemCode", parts[1] || "");
		setAttr(node, "ClassCode", parts[2] || "");
		setAttr(node, "ItemName", (parts[4] || "") + (appendAttributeText ? selectedAttributeText(attributeList) : ""));
		setAttr(node, "ClassName", parts[3] || "");
		setAttr(node, "StoresUoM", parts[5] || "");
		setAttr(node, "Decimal", parts[6] || "");
		setAttr(node, "ReceiptNum", parts[7] || "");
		setAttr(node, "AttributeList", attributeList || "");
		setAttr(node, "ItemRate", valueOrZero(hiddenValue("hItemRate", parts)));
		setAttr(node, "ItemStock", valueOrZero(hiddenValue("hItemStock", parts)));
		setAttr(node, "LocAndBinCount", valueOrZero(hiddenValue("hBinAndLocCheck", parts)));
		setAttr(node, "LocNo", valueOrZero(hiddenValue("hLocNo", parts)));
		setAttr(node, "BinNo", valueOrZero(hiddenValue("hBinNo", parts)));
		setAttr(node, "PartyCode", parts[9] || "");
		setAttr(node, "PartyType", parts[10] || "");
		setAttr(node, "PartySubType", parts[11] || "");
		setAttr(node, "SuppItemCode", parts[12] || "");
		setAttr(node, "SuppItemDesc", parts[13] || "");
		setAttr(node, "MarketPrice", valueOrZero(hiddenValue("hMarketPrice", parts)));
		root().appendChild(node);
	}

	function xmlFun(obj) {
		var parts = String(obj && obj.value || "").split(":");
		var attributeList = trim(fieldValue("hDisableBut")).toUpperCase() === "Y" ? selectedAttribute(parts, true) : "";
		if (selectMode() === "M") {
			if (obj && obj.checked) {
				appendItem(parts, attributeList, false);
			} else {
				removeItem(parts, attributeList);
			}
			displayList();
			return;
		}
		removeAllItems();
		if (obj && obj.checked) {
			appendItem(parts, attributeList, false);
		}
		displayList();
	}

	function addFun() {
		var count = parseInt(fieldValue("hChkCount"), 10) || 0;
		var check;
		var parts;
		var attributeList;
		for (var i = 0; i < count; i += 1) {
			check = field("pKey" + i);
			if (check && check.checked) {
				check.checked = false;
				parts = String(check.value || "").split(":");
				attributeList = selectedAttribute(parts, false);
				if (selectMode() === "M") {
					appendItem(parts, attributeList, true);
				} else {
					removeAllItems();
					appendItem(parts, attributeList, false);
				}
			}
		}
		displayAttribList();
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&#39;");
	}

	function renderSelected(includeAttribute) {
		var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
		selectedItems().forEach(function (item) {
			var attrList = includeAttribute ? trim(attr(item, "AttributeList")).replace(/:/g, "*") + ":" : "";
			var value = [
				attr(item, "CompanyItemCode"),
				attr(item, "ItemCode"),
				attr(item, "ClassCode"),
				includeAttribute ? attrList + attr(item, "ItemName") : attr(item, "ClassName") + ":" + attr(item, "ItemName")
			].join(":");
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeHtml(value) + '" checked onclick="RemoveNode(this)">';
			html += '</td><td class="ExcelDisplayCell">' + escapeHtml(attr(item, "CompanyItemCode")) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(item, "ItemName").replace(/~~/g, '"')) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(item, "ClassName")) + '</td></tr>';
		});
		if (!includeAttribute) {
			selectedMaterials().forEach(function (entry) {
				html += '<tr><td class="ExcelDisplayCell">';
				html += '<input type="checkbox" name="chk" value="' + escapeHtml(attr(entry, "SlNo")) + '" checked onclick="RemoveNode(this)">';
				html += '</td><td class="ExcelDisplayCell">--NA--</td>';
				html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(entry, "ItemName")) + '</td>';
				html += '<td class="ExcelDisplayCell">--NA--</td></tr>';
			});
		}
		html += "</table><br>";
		if (byId("idSelList")) {
			byId("idSelList").innerHTML = html;
		}
	}

	function displayAttribList() {
		renderSelected(true);
	}

	function displayList() {
		renderSelected(false);
	}

	function removeMaterial(slNo) {
		childElements(root(), "Materials").forEach(function (materials) {
			childElements(materials, "Entry").forEach(function (entry) {
				if (trim(attr(entry, "SlNo")) === trim(slNo)) {
					materials.removeChild(entry);
				}
			});
		});
	}

	function removeNode(obj) {
		var value = String(obj && obj.value || "");
		var parts = value.split(":");
		var attrList;
		if (obj && obj.checked) {
			return;
		}
		if (parts.length === 1) {
			removeMaterial(value);
			displayList();
			return;
		}
		attrList = parts.length > 4 ? trim(parts[3]).replace(/\*/g, ":") : null;
		removeItem(parts, attrList);
		asArray(document.querySelectorAll('input[type="checkbox"][name^="pKey"]')).forEach(function (check) {
			var checkParts = String(check.value || "").split(":");
			if (trim(parts[1]) === trim(checkParts[1]) && trim(parts[2]) === trim(checkParts[2])) {
				check.checked = false;
			}
		});
		if (attrList) {
			displayAttribList();
		} else {
			displayList();
		}
	}

	function showPage(args) {
		var query = String(args || "") +
			"&Query=" + encodeURIComponent(fieldValue("Query")) +
			"&SearchType=" + encodeURIComponent(fieldValue("SearchType")) +
			"&SearchBy=" + encodeURIComponent(fieldValue("SearchBy"));
		buttonPressed = "Page";
		setAttr(root(), "Action", "Page");
		setAttr(root(), "PassQuery", query);
		closeWithRoot();
	}

	function sendValue() {
		buttonPressed = "Done";
		setAttr(root(), "Action", "Done");
		closeWithRoot();
	}

	function init() {
		buttonPressed = "";
		if (trim(selectMode()) === "S" && field("BUTTON1")) {
			field("BUTTON1").disabled = true;
		}
		displayAttribList();
	}

	function deleteNodes() {
		removeAllItems();
	}

	window.showpage = showPage;
	window.ShowPage = showPage;
	window.SendValue = sendValue;
	window.XmlFun = xmlFun;
	window.XMLFun = xmlFun;
	window.AddFun = addFun;
	window.Init = init;
	window.DeleteNodes = deleteNodes;
	window.RemoveNode = removeNode;
	window.DispAttribList = displayAttribList;
	window.DispList = displayList;
	window.WithOutMat = function () {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("../Purchase/Transaction/SelMaterialNew.asp?orgID=010101&hSelectMode=M", window.dialogArguments, "dialogHeight:600px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (outRoot) {
				childElements(outRoot, "Materials").forEach(function (materials) {
					root().appendChild(materials.cloneNode(true));
				});
				displayList();
			});
		}
	};
	window.window_onunload = function () {
		if (!buttonPressed) {
			setAttr(root(), "Action", "CLOSE");
			removeAllItems();
		}
		setAttr(root(), "ItemType", "");
		returnValue(root());
	};
	window.addEventListener("beforeunload", function () {
		window.window_onunload();
	});
}(window, document));
