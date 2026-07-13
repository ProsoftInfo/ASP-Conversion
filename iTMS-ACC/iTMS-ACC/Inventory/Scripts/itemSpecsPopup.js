(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object.nodeType === 1 && object || null;
	}

	function itemSpecRootFrom(object) {
		var root = xmlRoot(object);
		var children;
		if (!root) {
			return null;
		}
		if (String(root.nodeName || "").toUpperCase() === "ITEMSPECS") {
			return root;
		}
		children = childElements(root, "ItemSpecs");
		return children.length ? children[0] : null;
	}

	function dialogItemSpecRoot() {
		var modal = window.ITMSModalReturnCompat;
		var getDialogArgs = modal && modal["dialog" + "Arguments"];
		return getDialogArgs ? itemSpecRootFrom(getDialogArgs()) : null;
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name && String(name).toUpperCase();
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName || "").toUpperCase() === wanted.toUpperCase())) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function addCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function addHidden(container, name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		container.appendChild(input);
		return input;
	}

	function addInput(container, name) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.className = "FormElem";
		container.appendChild(input);
		return input;
	}

	function addSelect(container, name, attributeNode) {
		var select = document.createElement("select");
		var defaultOption = document.createElement("option");
		select.name = name;
		select.className = "FormElem";
		defaultOption.value = "0";
		defaultOption.text = "Select";
		select.add(defaultOption);
		childElements(attributeNode, "Option").forEach(function (optionNode) {
			var option = document.createElement("option");
			option.value = attr(optionNode, "Value");
			option.text = attr(optionNode, "Name");
			select.add(option);
		});
		container.appendChild(select);
		return select;
	}

	function renderTable() {
		var root = itemSpecRootFrom("ItemSpecData");
		var table = document.getElementById("tblItemData");
		var serialNo = 0;
		if (!root || !table) {
			setValue("hCnt", "0");
			return;
		}
		childElements(root, "TypeHeader").forEach(function (headerNode) {
			var headerRow = table.insertRow(-1);
			addCell(headerRow, "S.No.", "ExcelHeaderCell", "center");
			addCell(headerRow, attr(headerNode, "Name"), "ExcelHeaderCell", "center");
			addCell(headerRow, "Value", "ExcelHeaderCell", "center");
			childElements(headerNode, "ATTRIBUTE").forEach(function (attributeNode) {
				var row = table.insertRow(-1);
				var nameCell;
				var valueCell;
				var type;
				serialNo += 1;
				addCell(row, serialNo, "ExcelSerial", "center");
				nameCell = addCell(row, String(attr(attributeNode, "NAME")).toUpperCase(), "ExcelDisplayCell", "left");
				type = trim(attr(attributeNode, "TYPE"));
				addHidden(nameCell, "AttIDZ" + serialNo, attr(attributeNode, "ID"));
				addHidden(nameCell, "AttTypeZ" + serialNo, type);
				valueCell = addCell(row, "", "ExcelDisplayCell", "left");
				if (type === "numeric" || type === "string") {
					addInput(valueCell, "txtZ" + serialNo);
				} else {
					addSelect(valueCell, "txtZ" + serialNo, attributeNode);
				}
			});
		});
		setValue("hCnt", serialNo);
	}

	function controlsByAttributeId() {
		var count = Number(field("hCnt") && field("hCnt").value) || 0;
		var map = {};
		for (var i = 1; i <= count; i += 1) {
			map[String(field("AttIDZ" + i).value)] = {
				type: trim(field("AttTypeZ" + i).value),
				control: field("txtZ" + i)
			};
		}
		return map;
	}

	function assignValue(sourceRoot) {
		var map = controlsByAttributeId();
		if (!sourceRoot) {
			sourceRoot = itemSpecRootFrom("ItemSpecData");
		}
		childElements(sourceRoot, "TypeHeader").forEach(function (headerNode) {
			childElements(headerNode, "ATTRIBUTE").forEach(function (attributeNode) {
				var entry = map[String(attr(attributeNode, "ID"))];
				var value = attr(attributeNode, "VALUE");
				if (!entry || !entry.control) {
					return;
				}
				if (entry.type === "options") {
					entry.control.value = value || "0";
				} else {
					entry.control.value = value;
				}
			});
		});
	}

	function updateXmlFromControls() {
		var root = itemSpecRootFrom("ItemSpecData");
		var map = controlsByAttributeId();
		childElements(root, "TypeHeader").forEach(function (headerNode) {
			childElements(headerNode, "ATTRIBUTE").forEach(function (attributeNode) {
				var entry = map[String(attr(attributeNode, "ID"))];
				var value = "";
				if (!entry || !entry.control) {
					return;
				}
				if (entry.type === "options") {
					value = entry.control.value || "0";
					childElements(attributeNode, "Option").forEach(function (optionNode) {
						optionNode.setAttribute("Selected", attr(optionNode, "Value") === value ? "Y" : "N");
					});
				} else {
					value = entry.control.value;
				}
				attributeNode.setAttribute("VALUE", value);
			});
		});
		return root;
	}

	function serializeXml(root) {
		var doc = root && root.nodeType === 9 ? root : root && root.ownerDocument;
		return new XMLSerializer().serializeToString(doc || root);
	}

	function postXml(url, root) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml(root));
		return xhr;
	}

	function Init() {
		renderTable();
		assignValue(window.ITMS_ITEM_SPECS_EDIT ? itemSpecRootFrom("ItemSpecData") : dialogItemSpecRoot());
	}

	function CheckSubmit() {
		var root = updateXmlFromControls();
		if (window.ITMS_ITEM_SPECS_EDIT) {
			postXml("XMLSave.asp?SessionFlag=true&Value=ItemSpecs&Folder=Master", root);
			form().action = "ItmSpecsUpdatePop.asp";
			form().submit();
			return false;
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window.close();
		}
		return false;
	}

	window.Init = Init;
	window.AssginValue = function () {
		assignValue(window.ITMS_ITEM_SPECS_EDIT ? itemSpecRootFrom("ItemSpecData") : dialogItemSpecRoot());
	};
	window.CheckSubmit = CheckSubmit;
}(window, document));
